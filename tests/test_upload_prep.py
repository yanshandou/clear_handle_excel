"""上传前每个文件自动做：清标记色 → 对照参考恢复残留颜色 → 视图标准化（core/upload_prep.py）。"""
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook

from core import batch, color_compare, distribute, excel_ops, mark_finder, storage, upload_prep, work_task, workflow as wf


def write(path: Path, text: str = "x") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


class PrepareTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.local, self.ref = Path(temp.name) / "work" / "a.xlsx", Path(temp.name) / "ref" / "a.xlsx"
        write(self.local, "local")
        write(self.ref, "ref")
        self.excel = Mock()
        self.book = Mock(ReadOnly=False)
        self.open_book = Mock(return_value=self.book)
        self.order = []

        def clear(excel, book):
            self.order.append("clear")
            return {"leftovers": 3, "leftover_failed": [], "tabs": [{"sheet": "S", "status": "ok", "reason": ""}]}

        cells = [{"sheet": "S", "addr": "A1", "anchor": "A1", "action": "restore"},
                 {"sheet": "S", "addr": "B2", "anchor": "B2", "action": "skip", "reason": "和备份同一位置的内容或公式不同，请人工检查"}]
        self.enterContext(patch.object(excel_ops, "open_workbook_checked", self.open_book))
        self.enterContext(patch.object(batch, "clear_workbook_marks", side_effect=clear))
        self.enterContext(patch.object(color_compare, "plan_xml", side_effect=lambda *a, **kw: self.order.append("compare") or {"cells": cells}))
        self.enterContext(patch.object(mark_finder, "_read_old", side_effect=lambda *a: self.order.append("read_reference") or {}))
        self.enterContext(patch.object(mark_finder, "restore_workbook_colors", side_effect=lambda book, ready, old: self.order.append("apply")
                                       or [{**ready[0], "status": "ok", "reason": "已改回旧版的颜色"}]))
        self.enterContext(patch.object(upload_prep, "_normalize_view", side_effect=lambda *a: self.order.append("normalize")))

    def test_clears_restores_and_normalizes_in_order(self):
        result = upload_prep.prepare(lambda: self.excel, self.local, self.ref)
        self.assertEqual(self.order, ["compare", "read_reference", "clear", "apply", "normalize"])
        self.open_book.assert_called_once_with(self.excel, str(self.local.resolve()), UpdateLinks=0)
        self.book.Save.assert_called_once_with()
        self.book.Close.assert_called_once_with(SaveChanges=False)
        self.assertEqual((result["marks"], result["tabs"], result["colors"], result["color_skipped"]), (3, 1, 1, 1))
        self.assertEqual(result["color_problems"][0]["cell"], "B2")
        self.assertTrue(result["normalized"])
        text = upload_prep.text(result)
        self.assertIn("清掉标记 3 条", text)
        self.assertIn("按参考改回颜色 1 处", text)
        self.assertIn("颜色改不了 1 处", text)

    def test_without_reference_file_colors_are_not_compared(self):
        self.ref.unlink()
        result = upload_prep.prepare(lambda: self.excel, self.local, self.ref)
        self.assertEqual(self.order, ["clear", "normalize"])
        self.open_book.assert_called_once()
        self.book.Save.assert_called_once_with()
        self.book.Close.assert_called_once_with(SaveChanges=False)
        self.assertIn("参考里没有这份文件", upload_prep.text(result))

    def test_failure_tells_which_step(self):
        upload_prep._normalize_view.side_effect = ValueError("不能标准化视图")
        where = {}
        with self.assertRaises(ValueError):
            upload_prep.prepare(lambda: self.excel, self.local, self.ref, where)
        self.assertEqual(where["step"], "视图标准化")
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_marks_that_cannot_be_removed_stop_the_upload(self):
        batch.clear_workbook_marks.side_effect = None
        batch.clear_workbook_marks.return_value = {"leftovers": 1, "leftover_failed": ["工作表受保护"], "tabs": []}
        with self.assertRaises(ValueError):
            upload_prep.prepare(lambda: self.excel, self.local, self.ref)
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)

    def test_reference_changes_during_preparation_prevent_save(self):
        upload_prep._normalize_view.side_effect = lambda *a: write(self.ref, "new reference")
        with self.assertRaisesRegex(ValueError, "当前文件或备份在检查颜色期间发生变化"):
            upload_prep.prepare(lambda: self.excel, self.local, self.ref)
        self.book.Save.assert_not_called()
        self.book.Close.assert_called_once_with(SaveChanges=False)


class ClearedTabPlanTests(unittest.TestCase):
    def test_same_orange_reference_is_planned_for_restoration_only_after_clear(self):
        with tempfile.TemporaryDirectory() as directory:
            local, reference = Path(directory) / "local.xlsx", Path(directory) / "reference.xlsx"
            book = Workbook()
            book.active.sheet_properties.tabColor = batch.TAB_HIGHLIGHT[1:]
            book.save(local)
            book.save(reference)
            book.close()

            self.assertEqual(color_compare.plan_xml(local, reference)["cells"], [])
            plan = color_compare.plan_xml(local, reference, clear_tab_color=batch.TAB_HIGHLIGHT)["cells"]
            self.assertEqual(len(plan), 1)
            self.assertEqual((plan[0]["tab_cleared"], plan[0]["action"], plan[0]["old_fill"]),
                             (True, "restore", batch.TAB_HIGHLIGHT))


class TaskUploadPrepTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.shared = root / "共享" / "仕様書"
        write(self.shared / "a.xlsx", "A")
        write(self.shared / "b.xlsx", "B")
        for obj, name, value in ((storage, "DATABASE_PATH", root / "test.sqlite3"),
                                 (distribute, "_STATUS", {"running": False})):
            patcher = patch.object(obj, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        wf.STATE["busy"] = False
        self.task = work_task.create(str(self.shared), str(root / "handle-excel"))
        distribute.preview_fetch("", "", "", self.task["id"])
        distribute.start()
        self.wait()
        self.work = Path(self.task["work"])
        write(self.work / "a.xlsx", "A fixed")
        write(self.work / "b.xlsx", "B fixed")

    def wait(self):
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def run_upload(self, prepare):
        with patch.object(upload_prep, "prepare", side_effect=prepare), \
                patch("core.excel_ops.start_background_excel", Mock()), patch("core.excel_ops.quit_background_excel", Mock()), \
                patch("core.file_problem.ask", Mock(return_value="skip")):
            distribute.start_task_upload(self.task["id"], ["a.xlsx", "b.xlsx"], {})
            return self.wait()

    def test_every_file_is_prepared_with_its_reference_before_upload(self):
        seen = []

        def prepare(get_excel, local, reference, where):
            seen.append((local.name, reference))
            return {"marks": 1, "tabs": 0, "colors": 2, "color_skipped": 0, "color_problems": [],
                    "reference": True, "normalized": True}

        result = self.run_upload(prepare)
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual([name for name, _ in seen], ["a.xlsx", "b.xlsx"])
        self.assertEqual(seen[0][1], Path(self.task["reference"]) / "a.xlsx")
        self.assertEqual((self.shared / "a.xlsx").read_text(encoding="utf-8"), "A fixed")
        self.assertEqual(result["files"][0]["prep"]["colors"], 2)
        self.assertIn("上传前：清掉标记 1 条", distribute.markdown(result))

    def test_one_failed_preparation_skips_only_that_file(self):
        def prepare(get_excel, local, reference, where):
            if local.name == "b.xlsx":
                where["step"] = "视图标准化"
                raise ValueError("文件正在 Excel 中打开")
            return {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [],
                    "reference": True, "normalized": True}

        result = self.run_upload(prepare)
        # 用户 2026-11-10：没做成的那份不传，其余照传
        self.assertEqual(result["summary"]["已复制"], 1)
        self.assertEqual(result["summary"]["没复制成功"], 1)
        self.assertIn("上传前视图标准化没做成", result["files"][1]["error"])
        self.assertNotEqual((self.shared / "a.xlsx").read_text(encoding="utf-8"), "A")
        self.assertEqual((self.shared / "b.xlsx").read_text(encoding="utf-8"), "B")


if __name__ == "__main__":
    unittest.main()
