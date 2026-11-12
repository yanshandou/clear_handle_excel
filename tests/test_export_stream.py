"""边导边看（用户 2026-11-12 定）：批量导出 PDF 还在跑时，导好的文件马上能在调整页面看、打开、改；
调整页面点的修改排进导出队列，导出做完手上这一份就让它先做，做完接着导。"""
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import MagicMock, patch

import web_clip_fixer as web
from core import batch, workflow as wf
from tests.test_pdf_review import RegenBase


def export_job(workbooks, completed=(), failed=(), folder=""):
    return {"kind": "export_pdf", "status": "running", "workbooks": [str(p) for p in workbooks],
            "completed": [str(p) for p in completed], "failed": [str(p) for p in failed],
            "options": {"folder": str(folder)}}


class OpenBlockTests(unittest.TestCase):
    def test_done_and_failed_open_current_and_later_wait(self):
        a, b, c, d = (Path(tempfile.gettempdir()) / f"{n}.xlsx" for n in "abcd")
        job = export_job([a, b, c, d], completed=[a], failed=[b])
        self.assertEqual(batch.open_block_reason(job, a), "")
        self.assertEqual(batch.open_block_reason(job, b), "")  # 出错跳过的这次不会再碰：可以打开修
        self.assertIn("正在处理", batch.open_block_reason(job, c))
        self.assertIn("还没轮到", batch.open_block_reason(job, d))
        self.assertEqual(batch.open_block_reason(job, Path(tempfile.gettempdir()) / "x.xlsx"), "")

    def test_clean_sheets_still_waits(self):
        # 精简 Sheet 做完还要接着自动检查，先不放
        a = Path(tempfile.gettempdir()) / "a.xlsx"
        job = {**export_job([a], completed=[a]), "kind": "clean_sheets"}
        self.assertTrue(batch.open_block_reason(job, a))


class OpenRequestTests(unittest.TestCase):
    """按真实请求从入口走到打开那一步（以前拦截改了一处漏一处，用户实测打不开）"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        self.a, self.b = root / "a.xlsx", root / "b.xlsx"
        for p in (self.a, self.b):
            p.write_bytes(b"x")
        self.folder = str(root)
        state = {**web.STATE, "busy": True, "folder": self.folder,
                 "job": export_job([self.a, self.b], completed=[self.a], folder=root)}
        state.pop("maintenance", None)
        for patcher in (patch.dict(web.STATE, state, clear=True), patch.object(web.autofix, "wait_for_writes"),
                        patch.object(web.batch, "note_opened")):
            patcher.start()
            self.addCleanup(patcher.stop)

    def post(self, path, payload):
        handler = object.__new__(web.WebHandler)
        handler.path, sent = path, []
        handler._send_json = lambda data, status=200: sent.append(data)
        handler._handle_post(payload)
        return sent[0]

    def test_exported_file_opens_waiting_one_does_not(self):
        with patch.object(web, "open_workbook_in_excel", return_value="") as opener:
            self.assertTrue(self.post("/api/excel/open", {"file": "a.xlsx", "folder": self.folder})["ok"])
            self.assertEqual(Path(opener.call_args.args[0]).resolve(), self.a.resolve())
            opener.reset_mock()
            with self.assertRaisesRegex(ValueError, "处理完才能打开"):
                self.post("/api/excel/open", {"file": "b.xlsx", "folder": self.folder})
        opener.assert_not_called()


class SideTurnTests(unittest.TestCase):
    def setUp(self):
        self.state = {"busy": False, "job": None}
        patcher = patch.object(wf, "STATE", self.state)
        patcher.start()
        self.addCleanup(patcher.stop)
        for name in ("log_msg",):
            p = patch.object(wf, name)
            p.start()
            self.addCleanup(p.stop)

    def test_no_export_takes_and_releases_task_slot(self):
        with wf.side_turn("pdf_review_regen"):
            self.assertEqual((self.state["busy"], self.state["maintenance"]), (True, "pdf_review_regen"))
        self.assertFalse(self.state["busy"])
        self.assertNotIn("maintenance", self.state)

    def test_waits_for_export_to_hand_over(self):
        self.state.update(busy=True, job=export_job([Path("a.xlsx")]))
        order, waits = [], []

        def side():
            with wf.side_turn("pdf_review_process", waits.append):
                order.append("side")

        worker = threading.Thread(target=side)
        worker.start()
        time.sleep(0.3)
        self.assertEqual(order, [])  # 导出还在做手上这一份：不能同时开 Excel
        self.assertTrue(waits)
        wf.let_side_work_run(MagicMock())
        order.append("export continues")
        worker.join(2)
        self.assertEqual(order, ["side", "export continues"])
        self.assertTrue(self.state["busy"])  # 任务槽还是导出的
        self.assertNotIn("maintenance", self.state)

    def test_nothing_waiting_returns_at_once(self):
        dog = MagicMock()
        wf.let_side_work_run(dog)
        dog.file.assert_not_called()


class RegenDuringExportTests(RegenBase):
    def setUp(self):
        super().setUp()
        self.other = self.tmp / "後.xlsx"
        self.other.write_bytes(b"y")

    def test_exported_file_regenerates_when_export_hands_over(self):
        self.state.update(busy=True, job=export_job([self.book, self.other], completed=[self.book], folder=self.tmp))
        with patch.object(wf, "_export_workbook_pdfs", return_value=[self.pdf]) as export:
            self.assertTrue(pdf_review_regenerate(self.pdf)["started"])
            time.sleep(0.3)
            export.assert_not_called()
            self.assertIn("等正在导出", wf_regen_status()["waiting"])
            wf.let_side_work_run(MagicMock())
            result = self.wait()
        self.assertTrue(result["ok"])
        export.assert_called_once()
        self.assertTrue(self.state["busy"])
        self.assertNotIn("maintenance", self.state)

    def test_file_still_waiting_to_export_is_refused(self):
        self.state.update(busy=True, job=export_job([self.other, self.book], folder=self.tmp))
        with self.assertRaisesRegex(ValueError, "还在导出队列"):
            pdf_review_regenerate(self.pdf)

    def test_list_hides_files_not_exported_yet(self):
        self.state.update(busy=True, job=export_job([self.other, self.book], folder=self.tmp))
        listed = pdf_review_list(self.tmp)
        self.assertEqual(listed["pdfs"], [])
        self.assertEqual((listed["exporting"]["total"], listed["exporting"]["done"], listed["exporting"]["hidden"]), (2, 0, 1))
        self.state["job"]["completed"].append(str(self.book))
        listed = pdf_review_list(self.tmp)
        self.assertEqual([e["pdf"] for e in listed["pdfs"]], [str(self.pdf)])
        self.state.update(busy=False)
        self.assertIsNone(pdf_review_list(self.tmp)["exporting"])


def pdf_review_regenerate(pdf):
    from core import pdf_review
    return pdf_review.regenerate(str(pdf), True)


def pdf_review_list(folder):
    from core import pdf_review
    return pdf_review.list_pdfs(str(folder))


def wf_regen_status():
    from core import pdf_review
    return pdf_review.regen_status()


if __name__ == "__main__":
    unittest.main()
