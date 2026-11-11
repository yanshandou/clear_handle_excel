"""修复图片查找进行中，结果里已有的文件点「打开 / 对照打开 / 开 PDF」必须真的走到打开那一步
（用户 2026-11-11 实测两次打不开：整条请求路径上任何一处拦截漏掉都会这样，所以按真实请求从头走一遍）。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import web_clip_fixer as web
from core import image_restore as ir


class OpenDuringScanRequestTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        root = Path(self.temp.name)
        self.new, self.old, self.pdf = root / "新" / "a.xlsx", root / "旧" / "a.xlsx", root / "新" / "a.pdf"
        for path in (self.new, self.old, self.pdf, root / "新" / "b.xlsx"):
            path.parent.mkdir(exist_ok=True)
            path.write_bytes(b"x")
        self.folder = str(root / "新")
        done = {"name": "a.xlsx", "path": str(self.new), "old": str(self.old), "pdf": str(self.pdf),
                "sheets": [{"sheet": "手順", "old_sheet": "手順"}]}
        status = {"running": True, "mode": "scan", "result": {"folder": self.folder, "files": [done]}}
        # 查找进行中的真实状态：busy + maintenance，主任务可能还留着以前的扫描
        state = {**web.STATE, "busy": True, "maintenance": "image_restore", "folder": self.folder,
                 "job": {"kind": "batch_scan", "workbooks": [str(self.new)], "completed": [], "scanned": []}}
        for patcher in (patch.object(ir, "_STATUS", status), patch.dict(web.STATE, state, clear=True),
                        patch.object(web.autofix, "wait_for_writes"), patch.object(web.batch, "note_opened")):
            patcher.start()
            self.addCleanup(patcher.stop)

    def post(self, path, payload):
        handler = object.__new__(web.WebHandler)
        handler.path, sent = path, []
        handler._send_json = lambda data, status=200: sent.append(data)
        handler._handle_post(payload)
        return sent[0]

    def test_finished_file_opens_in_excel(self):
        with patch.object(web, "open_workbook_in_excel", return_value="") as opener:
            self.assertTrue(self.post("/api/excel/open", {"file": "a.xlsx", "folder": self.folder, "sheet": "手順"})["ok"])
        self.assertEqual(Path(opener.call_args.args[0]).resolve(), self.new.resolve())

    def test_finished_file_opens_side_by_side(self):
        with patch.object(web.repair_compare, "open_comparison", return_value={}) as opener:
            self.assertTrue(self.post("/api/excel/compare", {"kind": "images", "folder": self.folder, "file": "a.xlsx",
                                                               "sheet": "手順", "cell": ""})["ok"])
        self.assertEqual(opener.call_args.args[0]["backup"], self.old.resolve())

    def test_finished_file_opens_pdf(self):
        with patch.object(ir.os, "startfile", create=True) as opener:
            self.assertTrue(self.post("/api/image_restore/open_pdf", {"file": "a.xlsx"})["ok"])
        opener.assert_called_once_with(str(self.pdf))

    def test_unfinished_file_and_writing_modes_stay_blocked(self):
        with patch.object(web, "open_workbook_in_excel") as opener:
            with self.assertRaisesRegex(ValueError, "还没查完"):
                self.post("/api/excel/open", {"file": "b.xlsx", "folder": self.folder})
            ir._STATUS["mode"] = "fix"
            with self.assertRaisesRegex(ValueError, "正在修复图片"):
                self.post("/api/excel/open", {"file": "a.xlsx", "folder": self.folder})
        opener.assert_not_called()


if __name__ == "__main__":
    unittest.main()
