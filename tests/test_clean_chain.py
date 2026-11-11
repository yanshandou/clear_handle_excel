"""精简 Sheet 之后自动接着做的检查：BL 超链接 → 引用错误 →（勾选时）隐藏页 → 统一字体（core/clean_chain.py）。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, call, patch

from core import bl_links, clean_chain, font_fix, hidden_repair, link_repair, storage


class CleanChainTests(unittest.TestCase):
    def setUp(self):
        temp = self.enterContext(tempfile.TemporaryDirectory())
        self.folder = Path(temp) / "current"
        self.reference = Path(temp) / "reference"
        self.folder.mkdir()
        self.reference.mkdir()
        self.a = self.folder / "a.xlsx"
        self.b = self.folder / "b.xlsx"
        for path in (self.a, self.b, self.reference / self.a.name, self.reference / self.b.name):
            path.write_bytes(path.name.encode())

        self.first = Mock(ReadOnly=False)
        self.second = Mock(ReadOnly=False)
        self.excel, self.reference_excel = Mock(), Mock()
        self.open = self.mock(clean_chain, "open_workbook_checked",
                              side_effect=lambda excel, path, **_: self.first if Path(path) == self.a else self.second)
        self.start = self.mock(clean_chain, "start_background_excel", side_effect=[self.excel, self.reference_excel])
        self.mock(clean_chain.file_problem, "quit_quietly")
        self.guard = Mock()
        self.guard.watchdog.skip_requested = self.guard.watchdog.stop_requested = False
        self.guard.file.return_value = False
        self.guard.watch.side_effect = lambda excel: excel
        self.guard.problem.side_effect = lambda error, *_, **kw: (False, str(error))
        self.mock(clean_chain.file_problem, "Guard", return_value=self.guard)
        self.mock(clean_chain.wf, "pythoncom")
        self.mock(clean_chain.wf, "log_msg")
        self.enterContext(patch.object(clean_chain.wf, "STATE", {"busy": False, "job": None}))
        self.enterContext(patch.object(storage, "DATABASE_PATH", Path(temp) / "db.sqlite3"))
        for module in (clean_chain, bl_links, link_repair, hidden_repair, font_fix):
            self.enterContext(patch.object(module, "_STATUS", {}))
        self.enterContext(patch.object(clean_chain, "_LOADED", False))

        self.bl_hit = {"sheet": "S", "addr": "B2", "anchor": "B2"}
        self.link_hit = {"sheet": "S", "addr": "C3", "anchor": "C3", "kind": "cell", "action": "restore"}
        self.link_skip = {"sheet": "S", "addr": "D4", "anchor": "D4", "old_sheet": "S_old",
                          "kind": "cell", "action": "skip", "reason": "备份工作表无法唯一配对，请人工检查"}
        self.hidden_hit = {"sheet": "S", "addr": "E:E", "anchor": "E1", "kind": "columns", "action": "restore"}
        self.bl_scan = self.mock(bl_links, "scan_workbook", side_effect=lambda book: [self.bl_hit] if book is self.first else [])
        self.bl_clear = self.mock(bl_links, "clear_workbook", return_value=[{**self.bl_hit, "status": "ok"}])
        self.link_scan = self.mock(link_repair, "scan_workbook", side_effect=lambda book: {
            "hits": [self.link_hit] if book is self.first else [], "sheets": [], "sheet_names": []})
        self.link_preview = self.mock(link_repair, "preview_file",
                                      side_effect=lambda _, entry, path: {**entry, "old": str(path),
                                          "old_fingerprint": clean_chain.sheet_history.fingerprint(path),
                                          "cells": [dict(self.link_hit), dict(self.link_skip)]})
        self.link_apply = self.mock(link_repair, "apply_workbook", return_value=[{**self.link_hit, "status": "ok"}])
        self.hidden_plan = self.mock(hidden_repair, "plan_file", side_effect=lambda _, path, old, options, book: {
            "old": str(old), "old_fingerprint": clean_chain.sheet_history.fingerprint(old),
            "checks": {"S": "fingerprint"}, "cells": [dict(self.hidden_hit)] if book is self.first else []})
        self.hidden_apply = self.mock(hidden_repair, "apply_workbook", return_value=[{**self.hidden_hit, "status": "ok"}])
        self.mock(font_fix, "yu_cells", return_value={})
        self.mock(font_fix, "refresh_changed_cells", return_value={})
        self.font = self.mock(font_fix, "fix_open_workbook", side_effect=lambda book, *args: {
            "shapes": 3 if book is self.first else 0, "cells": 0, "sheets": [{"sheet": "S"}] if book is self.first else []})
        self.chain = {"old_folder": str(self.reference), "hidden": True}

    def mock(self, module, name, **options):
        return self.enterContext(patch.object(module, name, **options))

    def steps(self):
        return {step["key"]: step for step in clean_chain.status()["steps"]}

    def test_all_steps_share_one_open_and_one_save_per_workbook(self):
        events = Mock()
        for name, method in (("bl", self.bl_clear), ("links", self.link_apply), ("hidden", self.hidden_apply), ("font", self.font)):
            events.attach_mock(method, name)
        clean_chain.run(str(self.folder), self.chain)
        self.assertEqual([event[0] for event in events.mock_calls], ["bl", "links", "hidden", "font", "font"])
        self.assertEqual([event.args[0] for event in events.mock_calls], [self.first] * 4 + [self.second])
        self.assertEqual(self.open.call_args_list, [
            call(self.excel, str(self.a), UpdateLinks=0, ReadOnly=False),
            call(self.excel, str(self.b), UpdateLinks=0, ReadOnly=False),
        ])
        self.first.Save.assert_called_once_with()
        self.first.Close.assert_called_once_with(SaveChanges=False)
        self.second.Save.assert_not_called()
        self.second.Close.assert_called_once_with(SaveChanges=False)
        self.assertIs(self.link_preview.call_args.args[0], self.reference_excel)
        self.assertIs(self.hidden_plan.call_args.args[0], self.reference_excel)
        self.assertEqual({k: s["fixed"] for k, s in self.steps().items()}, {"bl": 1, "links": 1, "hidden": 1, "font": 3})
        problem = self.steps()["links"]["problems"][0]
        self.assertEqual((problem["old"], problem["old_sheet"], problem["cell"]),
                         (str(self.reference / self.a.name), "S_old", "D4"))
        self.assertEqual(link_repair.status()["applied"]["files"][0]["cells"][0]["status"], "ok")
        self.assertEqual(hidden_repair.status()["applied"]["files"][0]["cells"][0]["status"], "ok")

        with patch.object(clean_chain, "_STATUS", {}), patch.object(clean_chain, "_LOADED", False):
            self.assertEqual(self.steps()["hidden"]["fixed"], 1)
            self.assertFalse(clean_chain.status()["running"])

    def test_without_reference_links_and_hidden_are_skipped(self):
        clean_chain.run(str(self.folder), {"hidden": True})
        steps = self.steps()
        self.assertEqual(steps["links"]["state"], "skipped")
        self.assertEqual(steps["hidden"]["state"], "skipped")
        self.assertEqual(steps["font"]["fixed"], 3)
        self.link_scan.assert_not_called()
        self.hidden_plan.assert_not_called()
        self.start.assert_called_once()

    def test_unticked_steps_are_not_run(self):
        clean_chain.run(str(self.folder), {**self.chain, "bl": False, "hidden": False, "font": False})
        self.assertEqual(self.steps()["bl"]["state"], "off")
        self.assertEqual(self.steps()["links"]["fixed"], 1)
        self.bl_scan.assert_not_called()
        self.hidden_plan.assert_not_called()
        self.font.assert_not_called()

    def test_plan_error_allows_later_steps_to_save(self):
        self.link_preview.side_effect = ValueError("备份发生变化，请重新扫描")
        clean_chain.run(str(self.folder), self.chain)
        steps = self.steps()
        self.assertEqual(steps["links"]["state"], "failed")
        self.assertEqual(steps["hidden"]["fixed"], 1)
        self.assertEqual(steps["font"]["fixed"], 3)
        self.first.Save.assert_called_once_with()

    def test_write_failure_discards_the_current_file_and_continues(self):
        self.link_apply.side_effect = RuntimeError("写入失败")
        clean_chain.run(str(self.folder), self.chain)
        self.first.Save.assert_not_called()
        self.first.Close.assert_called_once_with(SaveChanges=False)
        self.second.Close.assert_called_once_with(SaveChanges=False)
        self.assertEqual(self.steps()["bl"]["fixed"], 0)
        self.assertEqual(bl_links.status()["applied"]["files"][0]["cells"][0]["status"], "failed")
        self.assertEqual(self.steps()["links"]["problems"][0]["reason"], "写入失败")
        self.assertEqual(self.font.call_args_list, [call(self.second, {}, unittest.mock.ANY)])

    def test_failed_save_does_not_report_font_changes_as_saved(self):
        self.first.Save.side_effect = RuntimeError("无法保存")
        clean_chain.run(str(self.folder), self.chain)
        self.assertEqual(self.steps()["font"]["fixed"], 0)
        self.assertEqual(font_fix.status()["result"]["files"][0]["cells"], 0)
        self.assertEqual(font_fix.status()["result"]["files"][0]["error"], "无法保存")

    def test_busy_backend_is_reported_without_releasing_its_reservation(self):
        clean_chain.wf.STATE["busy"] = True
        clean_chain.run(str(self.folder), self.chain)
        self.assertEqual(self.steps()["bl"]["state"], "failed")
        self.open.assert_not_called()
        self.assertTrue(clean_chain.wf.STATE["busy"])

    def test_options_and_prepare_guard(self):
        self.assertIsNone(clean_chain.options_from(None))
        self.assertEqual(clean_chain.options_from({"hidden": True, "old_folder": ' "D:\\参考" '}),
                         {"bl": True, "links": True, "hidden": True, "font": True, "old_folder": "D:\\参考"})
        self.assertEqual(clean_chain.options_from({"bl": False, "font": False})["bl"], False)
        for bad in ("x", {"hidden": "yes"}, {"old_folder": 1}, {"bl": 0}):
            with self.assertRaises(ValueError):
                clean_chain.options_from(bad)
        clean_chain.prepare("/x", {"hidden": False, "old_folder": ""})
        clean_chain.drop("/x")
        self.assertEqual(clean_chain.status(), {})
        clean_chain.prepare("/x", {"hidden": False, "old_folder": ""})
        clean_chain._STATUS["stage"] = "chain"
        with self.assertRaises(ValueError):
            clean_chain.prepare("/y", {"hidden": False, "old_folder": ""})

    def test_interrupted_chain_is_marked_after_restart(self):
        clean_chain.prepare("/x", {"hidden": False, "old_folder": "/ref"})
        with patch.object(clean_chain, "_STATUS", {}), patch.object(clean_chain, "_LOADED", False):
            saved = clean_chain.status()
            self.assertFalse(saved["running"])
            self.assertTrue(saved["interrupted"])
            self.assertEqual({s["key"]: s["state"] for s in saved["steps"]}, {"bl": "failed", "links": "failed", "hidden": "off", "font": "failed"})


if __name__ == "__main__":
    unittest.main()


class ChainBlocksOtherJobsTests(unittest.TestCase):
    """精简后自动检查在跑时（主页任务已放开后台标志）不能开始别的任务：两边一起改 Excel，
    导出 PDF 刚记好的哈希又被自动检查存掉，打开调整页面就全部重导（用户 2026-11-10）"""

    def test_queue_job_refuses_while_chain_runs(self):
        from core import workflow as wf
        running = {"running": True, "stage": "chain",
                   "steps": [{"key": "font", "label": "统一字体", "state": "running"}]}
        with patch.object(clean_chain, "status", return_value=running), \
                patch.dict(wf.STATE, {"busy": False}):
            with self.assertRaises(ValueError) as caught:
                wf.queue_job("export_pdf", {"folder": "/x"})
        self.assertIn("统一字体", str(caught.exception))

    def test_clean_stage_does_not_block(self):
        from core import workflow as wf
        with patch.object(clean_chain, "status", return_value={"running": True, "stage": "clean"}), \
                patch.dict(wf.STATE, {"busy": True}):
            self.assertFalse(wf.queue_job("export_pdf", {"folder": "/x"}))
