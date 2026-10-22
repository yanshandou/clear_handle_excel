import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import excel_ops, storage
from core.common import REVIEW_EMOJI, strip_review_prefix, workbook_rel_name


class FakeInterior:
    def __init__(self, pattern=-4142, color=16777215):
        self.Pattern, self.Color = pattern, color
        self.TintAndShade = 0
        self.PatternColor, self.PatternColorIndex, self.PatternTintAndShade = 0, -4105, 0


class FakeArea:
    def __init__(self, interior, address):
        self.Interior, self.Address = interior, address
        self.MergeArea = self

    def Cells(self, row, col):
        return self


class FakeSheet:
    def __init__(self, cells):
        self.cells = cells

    def Range(self, address):
        key = address.replace("$", "")
        return self.cells.setdefault(key, FakeArea(FakeInterior(), f"${key[0]}${key[1:]}"))


class FakeBook:
    def __init__(self, test, path):
        self.test, self.path = test, path

    def Sheets(self, name):
        return FakeSheet(self.test.cells)

    def Save(self):
        self.test.saves += 1
        Path(self.path).write_bytes(b"saved %d" % self.test.saves)

    def Close(self, SaveChanges):
        pass


class LaterReviewTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "test.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)
        from core import autofix, reporter, workflow
        self.workflow = importlib.reload(workflow)
        reporter_patch = patch.object(reporter, "RECORD_PATH", self.root / "record.md")
        reporter_patch.start()
        self.addCleanup(reporter_patch.stop)
        self.autofix = importlib.reload(autofix)
        self.autofix._worker = Mock(is_alive=Mock(return_value=True))
        p = patch.object(self.workflow, "normalize_workbook_view", Mock())
        p.start()
        self.addCleanup(p.stop)

        self.wb_path = self.root / "設計書.xlsx"
        self.wb_path.write_bytes(b"workbook")
        self.cells = {"B5": FakeArea(FakeInterior(pattern=1, color=0xCCFFCC), "$B$5")}
        self.saves = 0
        self.excel = Mock()
        self.excel.Workbooks.Open.side_effect = lambda path, **kwargs: FakeBook(self, path)

        def item(iid, row, addr):
            return {
                "id": iid, "seq": row, "kind": "hidden", "wb_path": str(self.wb_path), "wb_name": "設計書.xlsx",
                "sheet_name": "S1", "excel_row": row, "cell_addr": addr, "row_height": 15.0, "page_index": 0,
                "page_id": "p1", "text": f"T{row}", "bbox": [0, 0, 1, 1], "occurrence": None,
                "status": "pending", "iterations": 0, "triage_keep": True,
            }
        self.workflow.STATE.update(
            busy=False, folder=str(self.root), job=None,
            items={"i1": item("i1", 5, "B5"), "i2": item("i2", 9, "C9")},
            pages={"p1": {"id": "p1", "wb_path": str(self.wb_path), "wb_name": "設計書.xlsx", "sheet_name": "S1"}},
            checked_files={str(self.wb_path): self.workflow._file_version(self.wb_path)},
            review_queue={},
        )

    def group(self, row):
        self.autofix.sync_groups()
        return next(g for g in self.autofix.AUTOFIX["groups"].values() if g["row"] == row)

    def defer(self, row):
        group = self.group(row)
        self.autofix.defer_group(group["id"], group["item_ids"])

    def marked_path(self):
        return self.wb_path.with_name(REVIEW_EMOJI + self.wb_path.name)

    def test_display_name_and_report_ignore_review_prefix(self):
        marked = self.root / "sub" / f"{REVIEW_EMOJI}a.xlsx"
        self.assertEqual(workbook_rel_name(marked, self.root), "sub/a.xlsx")
        self.assertEqual(workbook_rel_name(marked), "a.xlsx")
        self.assertEqual(strip_review_prefix(f"{REVIEW_EMOJI}a.xlsx"), "a.xlsx")

    def test_deferred_row_leaves_preview_queue_and_is_recorded(self):
        self.defer(5)
        item = self.workflow.STATE["items"]["i1"]
        self.assertTrue(item["later_review"])
        self.assertEqual(item["status"], "manual_review")
        self.assertTrue(self.group(5)["manual"])  # 不再做图
        self.assertFalse(self.group(9)["manual"])
        (entry,) = self.workflow.STATE["review_queue"].values()
        self.assertEqual([(m["sheet"], m["addr"], m["state"]) for m in entry["marks"]], [("S1", "B5", "pending")])
        self.assertEqual(self.autofix._next_task({}), ("mark", str(self.wb_path)))

    def test_no_marking_while_any_task_runs(self):
        self.defer(5)
        self.workflow.STATE["busy"] = True
        self.workflow.STATE["job"] = {"kind": "scan", "status": "running", "completed": [], "options": {}}
        self.assertNotEqual((self.autofix._next_task({}) or ("",))[0], "mark")

    def test_mark_paints_yellow_saves_and_renames_with_state_kept_in_sync(self):
        self.defer(5)
        self.group(9)["status"] = "applying"  # 同文件其他行已点选等待写入，改名后不能丢
        self.assertTrue(self.workflow.mark_review_file(self.excel, str(self.wb_path)))
        area = self.cells["B5"]
        self.assertEqual((area.Interior.Pattern, area.Interior.Color), (1, 0x00FFFF))
        new = self.marked_path()
        self.assertTrue(new.exists())
        self.assertFalse(self.wb_path.exists())
        state = self.workflow.STATE
        self.assertEqual({it["wb_path"] for it in state["items"].values()}, {str(new)})
        self.assertEqual(state["items"]["i1"]["wb_name"], "設計書.xlsx")
        self.assertEqual(state["pages"]["p1"]["wb_path"], str(new))
        self.assertEqual(list(state["review_queue"]), [str(new)])
        self.assertEqual(state["review_queue"][str(new)]["marks"][0]["state"], "marked")
        self.assertEqual(state["review_queue"][str(new)]["marks"][0]["fill"]["Color"], 0xCCFFCC)
        # 只是标黄：检查基线跟着前进，下次扫描不当成改过的文件
        self.assertEqual(state["checked_files"], {str(new): self.workflow._file_version(new)})
        self.assertEqual(self.group(9)["status"], "applying")
        self.assertEqual(self.group(9)["wb_path"], str(new))
        self.assertNotEqual((self.autofix._next_task({}) or ("",))[0], "mark")

    def test_open_file_is_not_marked_until_closed(self):
        self.defer(5)
        lock = self.wb_path.with_name("~$" + self.wb_path.name)
        lock.write_bytes(b"")
        self.assertFalse(self.workflow.mark_review_file(self.excel, str(self.wb_path)))
        self.assertEqual(self.saves, 0)
        self.assertIn("正在 Excel 中打开", self.workflow.STATE["review_queue"][str(self.wb_path)]["error"])

    def test_restore_only_fixed_cells_then_drop_prefix_when_file_is_clean(self):
        self.defer(5)
        self.defer(9)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        new = str(self.marked_path())
        items = self.workflow.STATE["items"]
        items["i1"]["status"] = "fixed"
        self.assertEqual(self.workflow.restore_review_file(self.excel, new), "partial")
        self.assertEqual((self.cells["B5"].Interior.Pattern, self.cells["B5"].Interior.Color), (1, 0xCCFFCC))
        self.assertEqual(self.cells["C9"].Interior.Color, 0x00FFFF)
        self.assertTrue(Path(new).exists())
        self.assertNotIn("later_review", items["i1"])

        items["i2"].update(status="accepted")
        self.assertEqual(self.workflow.restore_review_file(self.excel, new), "clean")
        self.assertEqual(self.cells["C9"].Interior.Pattern, -4142)  # 原来没有底色
        self.assertTrue(self.wb_path.exists())
        self.assertEqual(self.workflow.STATE["review_queue"], {})
        self.assertEqual(items["i2"]["wb_path"], str(self.wb_path))

    def test_cell_recoloured_by_user_is_left_alone(self):
        self.defer(5)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        self.cells["B5"].Interior.Color = 0x0000FF
        self.workflow.STATE["items"]["i1"]["status"] = "fixed"
        self.assertEqual(self.workflow.restore_review_file(self.excel, str(self.marked_path())), "clean")
        self.assertEqual(self.cells["B5"].Interior.Color, 0x0000FF)

    def test_redetected_problem_stays_in_later_review(self):
        self.defer(5)
        old = self.workflow.STATE["items"]["i1"]
        new_item = {**{k: v for k, v in old.items() if k not in {"later_review", "manual_waiting"}},
                    "id": "n1", "status": "pending"}
        merged = self.workflow._merge_redetected({"n1": new_item}, [old], mark_new=True, keep_unmatched=True)
        self.assertTrue(merged["i1"]["later_review"])
        self.assertEqual(merged["i1"]["status"], "manual_review")

    def test_rename_clash_keeps_name_and_stops_retrying(self):
        self.defer(5)
        self.marked_path().write_bytes(b"other")
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        entry = self.workflow.STATE["review_queue"][str(self.wb_path)]
        self.assertIn("同名文件", entry["rename_error"])
        self.assertFalse(self.workflow.review_entry_needs_mark(str(self.wb_path), entry))

    def run_recheck(self, fixed_ids, unsaved=None):
        self.workflow.STATE["job"] = {"kind": "review_recheck", "options": {}, "completed": [], "status": "running"}

        def recheck(ids):
            self.rechecked = list(ids)
            for iid in fixed_ids:
                self.workflow.STATE["items"][iid]["status"] = "fixed"
        with patch.object(self.workflow, "recheck_items_workflow", side_effect=recheck), \
                patch.object(self.workflow, "close_saved_workbooks", return_value=unsaved or {}), \
                patch.object(excel_ops, "win32", Mock(DispatchEx=Mock(return_value=self.excel))):
            self.workflow.review_recheck_workflow()

    def test_batch_recheck_checks_only_queued_problems_and_cleans_fixed_files(self):
        self.defer(5)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        self.run_recheck(["i1"])
        self.assertEqual(self.rechecked, ["i1"])
        self.assertTrue(self.wb_path.exists())
        self.assertEqual(self.workflow.STATE["review_queue"], {})
        self.assertEqual(self.cells["B5"].Interior.Color, 0xCCFFCC)

    def test_batch_recheck_skips_unsaved_file_and_keeps_it_queued(self):
        self.defer(5)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        self.run_recheck([], unsaved={self.marked_path(): "Excel 里还有没保存的修改"})
        self.assertEqual(self.rechecked if hasattr(self, "rechecked") else [], [])
        self.assertTrue(self.marked_path().exists())
        self.assertEqual(self.cells["B5"].Interior.Color, 0x00FFFF)
        self.assertEqual(list(self.workflow.STATE["review_queue"]), [str(self.marked_path())])

    def test_resolved_cells_are_restored_in_background(self):
        self.defer(5)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        new = str(self.marked_path())
        self.assertNotEqual((self.autofix._next_task({}) or ("",))[0], "restore")
        # 不经统一复检就解决了（例如标了无需修改）：后台自己去黄、去 🟡
        self.workflow.STATE["items"]["i1"]["status"] = "accepted"
        self.assertEqual(self.autofix._next_task({}), ("restore", new))
        self.assertEqual(self.workflow.restore_review_file(self.excel, new), "clean")
        self.assertTrue(self.wb_path.exists())
        self.assertEqual(self.cells["B5"].Interior.Color, 0xCCFFCC)

    def test_review_in_other_folder_is_left_alone_after_switching_folders(self):
        self.defer(5)
        self.workflow.mark_review_file(self.excel, str(self.wb_path))
        new = str(self.marked_path())
        other = self.root.parent / (self.root.name + "_other")
        other.mkdir()
        self.addCleanup(other.rmdir)
        # 换文件夹完整扫描：旧问题记录被清掉
        self.workflow.STATE.update(folder=str(other), items={}, pages={}, checked_files={})
        self.assertIsNone(self.autofix._next_task({}))
        self.assertEqual(self.workflow.get_state_snapshot()["review_queue"], {})
        with self.assertRaises(ValueError):
            self.workflow.queue_job("review_recheck")
        # 没检查过，不能当成已解决去掉黄色
        self.assertEqual(self.workflow.restore_review_file(self.excel, new), "partial")
        self.assertEqual(self.cells["B5"].Interior.Color, 0x00FFFF)

        # 换回来重新扫描：同一格子仍有问题，接回稍后人工复检
        self.workflow.STATE.update(folder=str(self.root), checked_files={new: {}})
        self.workflow.STATE["items"]["n1"] = {
            "id": "n1", "seq": 1, "kind": "hidden", "wb_path": new, "wb_name": "設計書.xlsx",
            "sheet_name": "S1", "excel_row": 5, "cell_addr": "B5", "status": "pending",
        }
        self.workflow._relink_review_marks(new)
        self.assertEqual(self.workflow.STATE["review_queue"][new]["marks"][0]["item_ids"], ["n1"])
        self.assertTrue(self.workflow.STATE["items"]["n1"]["later_review"])
        self.assertEqual(self.workflow.restore_review_file(self.excel, new), "partial")


if __name__ == "__main__":
    unittest.main()
