"""上传时导出同名 PDF 并一起上传（用户 2026-11-10）：终点已有同名 PDF 时默认不覆盖，上传前先列出来让用户选"""
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import distribute, pdf_bundle, storage, upload_prep, work_task, workflow as wf

PREP = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [], "reference": False, "normalized": True}
PDF_ON = {"options": {"include_hidden": False, "skip_sheets": ["変更履歴"], "no_header_footer": True}}


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


def fake_export(excel, wb_path, include_hidden, on_sheet=None, skip_sheets=None, no_header_footer=False):
    target = pdf_bundle.pdf_path_for(wb_path)
    on_sheet and on_sheet(1, 1, "Sheet1")
    target.write_text("PDF of " + read(wb_path), encoding="utf-8")
    return {"pdf": str(target), "sheets": ["Sheet1"], "pages": 2, "skipped": [{"sheet": "変更履歴", "reason": "在跳过名单里"}]}


class UploadPdfTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.company = root / "公司" / "仕様書"
        write(self.company / "a.xlsx", "A")
        write(self.company / "sub" / "b.xlsx", "B")
        for obj, name, value in ((storage, "DATABASE_PATH", root / "test.sqlite3"),
                                 (distribute, "_STATUS", {"running": False})):
            patcher = patch.object(obj, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        wf.STATE["busy"] = False
        self.task = work_task.create(str(self.company), str(root / "handle-excel"))
        distribute.preview_fetch("", "", "", self.task["id"])
        distribute.start()
        self.wait()
        self.work = Path(self.task["work"])
        write(self.work / "a.xlsx", "A fixed")
        write(self.work / "sub" / "b.xlsx", "B fixed")
        self.export = Mock(side_effect=fake_export)
        for p in (patch.object(upload_prep, "prepare", side_effect=lambda get_excel, local, reference, where: dict(PREP)),
                  patch.object(pdf_bundle, "export_workbook", self.export),
                  patch("core.excel_ops.start_background_excel", Mock()), patch("core.excel_ops.quit_background_excel", Mock()),
                  patch("core.file_problem.ask", Mock(return_value="skip"))):
            p.start()
            self.addCleanup(p.stop)

    def wait(self):
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def upload(self, pdf, files=("a.xlsx", "sub/b.xlsx")):
        distribute.start_task_upload(self.task["id"], list(files), {}, pdf)
        return self.wait()

    def test_pdf_uploaded_next_to_excel(self):
        result = self.upload(PDF_ON)
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual(result["summary"]["PDF 已传"], 2)
        self.assertEqual(read(self.company / "a.pdf"), "PDF of A fixed")
        self.assertEqual(read(self.company / "sub" / "b.pdf"), "PDF of B fixed")
        self.assertTrue((self.work / "a.pdf").is_file())  # 本地处理文件夹也留一份（和批量导出同名 PDF 一样）
        self.assertIn("同名 PDF a.pdf 已新建", distribute.markdown(result))
        self.assertFalse(distribute.needs_attention(result))
        self.assertNotIn("_pdf_local", result["files"][0])

    def test_without_option_no_pdf(self):
        result = self.upload(None)
        self.assertEqual(result["summary"]["已复制"], 2)
        self.export.assert_not_called()
        self.assertFalse((self.company / "a.pdf").exists())
        self.assertNotIn("PDF 已传", result["summary"])

    def test_existing_pdf_is_listed_and_kept_by_default(self):
        write(self.company / "a.pdf", "别人放的 PDF")
        check = distribute.pdf_check(self.task["id"], ["a.xlsx", "sub/b.xlsx"])
        existing = {f["rel"]: f for f in check["files"] if f["exists"]}
        self.assertEqual(list(existing), ["a.xlsx"])
        self.assertFalse(existing["a.xlsx"]["ours"])
        result = self.upload(PDF_ON)
        self.assertEqual(read(self.company / "a.pdf"), "别人放的 PDF")
        self.assertEqual(result["files"][0]["pdf"]["status"], "kept")
        self.assertEqual(result["summary"]["PDF 没覆盖"], 1)
        self.assertEqual(result["summary"]["已复制"], 2)  # Excel 照传
        self.assertEqual(read(self.company / "sub" / "b.pdf"), "PDF of B fixed")

    def test_overwrite_when_chosen_and_unchanged(self):
        write(self.company / "a.pdf", "旧 PDF")
        check = distribute.pdf_check(self.task["id"], ["a.xlsx"])
        sha = check["files"][0]["sha256"]
        result = self.upload({**PDF_ON, "overwrite": {"a.xlsx": sha}}, ["a.xlsx"])
        self.assertEqual(result["files"][0]["pdf"]["status"], "ok")
        self.assertEqual(read(self.company / "a.pdf"), "PDF of A fixed")
        # 传过一次之后再查，就认得是上次本工具传的
        self.assertTrue(distribute.pdf_check(self.task["id"], ["a.xlsx"])["files"][0]["ours"])

    def test_chosen_overwrite_but_changed_after_check_is_kept(self):
        write(self.company / "a.pdf", "旧 PDF")
        sha = distribute.pdf_check(self.task["id"], ["a.xlsx"])["files"][0]["sha256"]
        write(self.company / "a.pdf", "检查之后别人又换了")
        result = self.upload({**PDF_ON, "overwrite": {"a.xlsx": sha}}, ["a.xlsx"])
        self.assertEqual(read(self.company / "a.pdf"), "检查之后别人又换了")
        self.assertIn("检查之后又被改过", result["files"][0]["pdf"]["error"])

    def test_export_failure_skips_only_that_file(self):
        self.export.side_effect = lambda excel, wb_path, *a, **k: (_ for _ in ()).throw(RuntimeError("Excel 导不出来")) \
            if wb_path.name == "a.xlsx" else fake_export(excel, wb_path, *a, **k)
        result = self.upload(PDF_ON)
        self.assertEqual(result["summary"]["已复制"], 1)
        self.assertIn("上传前导出同名 PDF没做成", result["files"][0]["error"])
        self.assertEqual(read(self.company / "a.xlsx"), "A")

    def test_free_upload_carries_pdf(self):
        other = Path(self.tmp.name) / "中转"
        write(other / "x.xlsx", "X")
        dest = Path(self.tmp.name) / "终点"
        write(dest / "x.pdf", "终点原来的 PDF")
        preview = distribute.preview_free(str(other), str(dest))
        check = distribute.pdf_check(preview_id=preview["id"])
        self.assertTrue(check["files"][0]["exists"])
        distribute.start_free(preview["id"], False, {**PDF_ON, "overwrite": {"x.xlsx": check["files"][0]["sha256"]}})
        result = self.wait()
        self.assertEqual(result["summary"]["PDF 已传"], 1)
        self.assertEqual(read(dest / "x.pdf"), "PDF of X")
        self.assertFalse((other / "x.pdf").exists())  # 起点不动，PDF 从临时副本导出
        self.assertIn("同名 PDF x.pdf 已覆盖", distribute.markdown(result))


    def test_pdf_only_leaves_excel_alone(self):
        write(self.company / "a.pdf", "别人放的 PDF")
        result = self.upload({**PDF_ON, "only": True})
        self.assertEqual(read(self.company / "a.xlsx"), "A")  # Excel 不传
        self.assertEqual(read(self.company / "a.pdf"), "别人放的 PDF")  # 没选覆盖
        self.assertEqual(read(self.company / "sub" / "b.pdf"), "PDF of B fixed")
        self.assertTrue(all(f["pdf_only"] and f["status"] == "ok" for f in result["files"]))
        self.assertIn("只传 PDF，Excel 没传", distribute.markdown(result))
        files = {f["rel"]: f for f in distribute.task_files(self.task["id"])["files"]}
        self.assertFalse(files["a.xlsx"]["uploaded"])  # Excel 没记成已上传

    def test_free_upload_pdf_only_needs_no_excel_confirm(self):
        other = Path(self.tmp.name) / "中转"
        write(other / "x.xlsx", "X")
        dest = Path(self.tmp.name) / "终点"
        write(dest / "x.xlsx", "终点原来的 Excel")
        preview = distribute.preview_free(str(other), str(dest))
        distribute.start_free(preview["id"], False, {**PDF_ON, "only": True})
        result = self.wait()
        self.assertEqual(read(dest / "x.xlsx"), "终点原来的 Excel")
        self.assertEqual(read(dest / "x.pdf"), "PDF of X")
        self.assertEqual(result["summary"]["PDF 已传"], 1)


    def test_pdf_only_prepares_temp_copy_and_keeps_local_excel(self):
        def prepare(get_excel, local, reference, where):
            local.write_text(read(local) + " +处理", encoding="utf-8")
            return dict(PREP)
        upload_prep.prepare.side_effect = prepare
        result = self.upload({**PDF_ON, "only": True}, ["a.xlsx"])
        self.assertEqual(read(self.work / "a.xlsx"), "A fixed")  # 处理文件夹里的 Excel 没动
        self.assertEqual(read(self.company / "a.pdf"), "PDF of A fixed +处理")  # PDF 是处理过的临时副本导的
        self.assertEqual(read(self.work / "a.pdf"), "PDF of A fixed +处理")  # 处理文件夹也留一份
        self.assertEqual(result["files"][0]["status"], "ok")
        self.assertNotIn("_temp_root", result)
        files = {f["rel"]: f for f in distribute.task_files(self.task["id"])["files"]}
        self.assertTrue(files["a.xlsx"]["pdf_uploaded_at"])
        self.assertFalse(files["sub/b.xlsx"]["pdf_uploaded_at"])

    def test_shared_changed_at_copy_time_marks_prep_and_pdf_stale(self):
        def export(excel, wb_path, *args, **kwargs):
            info = fake_export(excel, wb_path, *args, **kwargs)
            if wb_path.name == "a.xlsx":
                write(self.company / "a.xlsx", "A 别人刚改")  # 处理期间共享被人改了
            return info
        self.export.side_effect = export
        result = self.upload(PDF_ON)
        a, b = result["files"]
        self.assertEqual(a["status"], "review")
        self.assertTrue(a["prep_stale"])
        self.assertEqual(a["pdf"]["status"], "stale")
        self.assertIn("复制前一刻发现共享里的这份被人改过", a["error"])
        self.assertFalse((self.work / "a.pdf").exists())  # 过期 PDF 删掉了
        self.assertEqual(read(self.work / "a.xlsx"), "A 别人刚改")  # 同步了共享最新版
        self.assertEqual(b["status"], "ok")  # 不挡后面的，照常传上去
        self.assertEqual(read(self.company / "sub" / "b.xlsx"), "B fixed")
        self.assertIn("同步前那一版", distribute.markdown(result))


if __name__ == "__main__":
    unittest.main()
