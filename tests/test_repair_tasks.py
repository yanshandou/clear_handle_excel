"""一键修复任务：连续执行、保留跳过原因、全程写入互斥。"""
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import link_repair, mark_finder, workflow as wf


class RepairTasksTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.current, self.old = self.root / "current" / "001", self.root / "backup" / "001"
        self.current.mkdir(parents=True)
        self.old.mkdir(parents=True)
        self.path, self.backup = self.current / "book.xlsx", self.old / "book.xlsx"
        self.path.write_bytes(b"current")
        self.backup.write_bytes(b"backup")
        self.state = {"busy": False, "job": None}
        self.excel = Mock()
        for obj, name, value in (
            (wf, "STATE", self.state), (wf, "log_msg", Mock()),
            (wf, "pythoncom", SimpleNamespace(CoInitialize=Mock(), CoUninitialize=Mock())),
            (link_repair, "_STATUS", {"running": False}), (mark_finder, "_STATUS", {"running": False}),
            (link_repair, "start_background_excel", Mock(return_value=self.excel)),
            (link_repair, "quit_background_excel", Mock()),
            (link_repair.debug_log, "event", Mock()), (link_repair.debug_log, "exception", Mock()),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.hits = [{"sheet": "Main", "addr": addr, "formula": "=#REF!", "fill": "#C0007A", "font": "#FFFFFF"}
                     for addr in ("A1", "A2")]
        self.plan = [{**self.hits[0], "action": "restore", "reason": "可以修复"},
                     {**self.hits[1], "action": "skip", "reason": "内容已变，请人工检查"}]

    def scanned(self, excel, path):
        self.assertTrue(self.state["busy"])
        return {"path": str(path), "hits": self.hits, "fingerprint": link_repair.fingerprint(path)}

    def test_link_scan_compare_apply_run_under_one_reservation_and_keep_skips(self):
        self.state.update(busy=True, maintenance="link_repair")

        def preview(excel, file, old):
            self.assertTrue(self.state["busy"])
            return {**file, "old": str(old), "cells": self.plan}

        def apply(excel, file):
            self.assertTrue(self.state["busy"])
            self.assertEqual([cell["addr"] for cell in file["cells"]], ["A1"])
            return [{**file["cells"][0], "status": "ok", "reason": "已修复"}]

        with patch.object(link_repair, "scan_file", side_effect=self.scanned), \
                patch.object(link_repair, "preview_file", side_effect=preview), \
                patch.object(link_repair, "apply_file", side_effect=apply):
            link_repair._run("repair", self.current, [self.path], self.old.parent, [self.backup])
        data = link_repair.status()["applied"]
        self.assertEqual([cell.get("status") for cell in data["files"][0]["cells"]], ["ok", None])
        self.assertIn("人工检查", data["files"][0]["cells"][1]["reason"])
        self.assertEqual(data["old_folder"], str(self.old.parent))
        self.assertFalse(self.state["busy"])
        self.assertNotIn("maintenance", self.state)

    def test_link_missing_backup_and_apply_failures_keep_problem_locations(self):
        for backups, fail in (([], False), ([self.backup], True)):
            with self.subTest(fail=fail):
                self.state.update(busy=True, maintenance="link_repair")
                with patch.object(link_repair, "scan_file", side_effect=self.scanned), \
                        patch.object(link_repair, "preview_file", return_value={"cells": self.plan}), \
                        patch.object(link_repair, "apply_file", side_effect=RuntimeError("保存失败")) as apply:
                    link_repair._run("repair", self.current, [self.path], self.old, backups)
                file = link_repair.status()["applied"]["files"][0]
                self.assertTrue(file["error"])
                self.assertEqual(len(file["cells"]), 2)
                self.assertFalse(any(cell.get("status") == "ok" for cell in file["cells"]))
                self.assertEqual(apply.call_count, int(fail))

    def run_color(self, **patches):
        self.state.update(busy=True, maintenance="color_repair")
        defaults = {"compare_file_colors": Mock(return_value={"cells":self.plan, "truncated":False}),
                    "apply_file": Mock(return_value=[{**self.plan[0], "status": "ok", "reason": "已恢复颜色"}]),
                    "workbook_locked": Mock(return_value=False)}
        defaults.update(patches)
        with patch.multiple(mark_finder, **defaults), \
                patch("core.excel_ops.start_background_excel", return_value=self.excel), \
                patch("core.excel_ops.quit_background_excel"):
            mark_finder._run_repair(self.current, [self.path], self.old.parent, [self.backup], [("C0007A", "FFFFFF")])
        return mark_finder.status()["applied"]["files"][0], defaults

    def test_color_task_only_applies_verified_cells_and_keeps_skipped_cells(self):
        file, mocks = self.run_color()
        self.assertEqual([cell.get("status") for cell in file["cells"]], ["ok", None])
        self.assertEqual(mocks["apply_file"].call_args.args[3], [self.plan[0]])
        self.assertEqual(file["old"], str(self.backup))
        self.assertFalse(self.state["busy"])
        self.assertNotIn("maintenance", self.state)

    def test_color_task_keeps_tab_and_cell_outcomes_separate_at_the_same_address(self):
        tab = {**mark_finder._tab_hit("Main"), "action":"restore", "reason":"恢复标签"}
        original = self.plan
        self.plan = [tab, *original]
        results = [{**tab, "status":"ok"}, {**original[0], "status":"ok", "reason":"恢复单元格"}]
        file, _ = self.run_color(apply_file=Mock(return_value=results))
        self.assertEqual(len(file["cells"]), 3)
        self.assertEqual([item.get("kind", "cell") for item in file["cells"]], ["tab", "cell", "cell"])
        self.assertEqual(sum(item.get("status") == "ok" for item in file["cells"]), 2)

    def test_color_open_targets_and_changed_files_are_never_written(self):
        file, mocks = self.run_color(workbook_locked=Mock(return_value=True))
        self.assertIn("Excel 中打开", file["error"])
        mocks["apply_file"].assert_not_called()

        def change_during_plan(*args):
            self.path.write_bytes(b"changed externally")
            return {"cells":self.plan}

        file, mocks = self.run_color(compare_file_colors=Mock(side_effect=change_during_plan))
        self.assertIn("发生变化", file["error"])
        mocks["apply_file"].assert_not_called()

    def test_color_no_matches_does_not_start_excel_and_failed_saves_keep_all_cells(self):
        file, mocks = self.run_color(compare_file_colors=Mock(return_value={"cells":[]}))
        self.assertEqual(file["cells"], [])
        mocks["apply_file"].assert_not_called()
        file, mocks = self.run_color(apply_file=Mock(side_effect=RuntimeError("保存失败")))
        self.assertIn("保存失败", file["error"])
        self.assertEqual(len(file["cells"]), 2)
        self.assertFalse(self.state["busy"])

    def test_start_validates_both_folders_and_reserves_before_background_thread(self):
        for module, args in ((link_repair, (str(self.current), str(self.old))),
                             (mark_finder, (str(self.current), str(self.old), [{"fill":"#C0007A", "font":"#FFFFFF"}]))):
            with self.subTest(module=module.__name__):
                with patch.object(module.threading, "Thread") as thread:
                    module.repair(*args)
                    self.assertTrue(self.state["busy"])
                    self.assertTrue(module.status()["running"])
                    thread.return_value.start.assert_called_once()
                    with self.assertRaises(ValueError):
                        module.repair(*args)
                self.state.update(busy=False)
                module._STATUS["running"] = False

    def test_same_folder_unsettled_writes_and_thread_failures_do_not_leave_busy_state(self):
        for module in (link_repair, mark_finder):
            extra = [] if module is link_repair else [[{"fill":"#C0007A", "font":"#FFFFFF"}]]
            with self.subTest(module=module.__name__):
                with self.assertRaises(ValueError):
                    module.repair(str(self.current), str(self.current), *extra)
                self.state["job"] = {"pending_adjustment": {"path": str(self.path)}}
                with self.assertRaises(ValueError):
                    module.repair(str(self.current), str(self.old), *extra)
                self.state["job"] = None
                with patch.object(module.threading, "Thread") as thread:
                    thread.return_value.start.side_effect = RuntimeError("cannot start")
                    with self.assertRaises(RuntimeError):
                        module.repair(str(self.current), str(self.old), *extra)
                self.assertFalse(self.state["busy"])
                self.assertNotIn("maintenance", self.state)
                self.assertFalse(module.status()["running"])


if __name__ == "__main__":
    unittest.main()
