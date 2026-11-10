"""精简 Sheet 之后自动接着做的检查：BL 超链接 → 引用错误 →（勾选时）隐藏页 → 统一字体（core/clean_chain.py）。"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import bl_links, clean_chain, font_fix, hidden_repair, link_repair, storage


class FakeJob:
    """假装后台任务：start 时马上把结果放好，running 一直是 False"""

    def __init__(self, module, key, result, error=None):
        self.module, self.key, self.result, self.error = module, key, result, error
        self.calls = []

    def __call__(self, *args):
        self.calls.append(args)
        if self.error:
            raise self.error
        self.status[self.key] = self.result

    status: dict


class CleanChainTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        for obj, name, value in ((storage, "DATABASE_PATH", Path(temp.name) / "db.sqlite3"),
                                 (clean_chain, "_STATUS", {}), (clean_chain, "_LOADED", False),
                                 (clean_chain, "POLL_SECONDS", 0)):
            patcher = patch.object(obj, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.bl_status, self.link_status, self.hidden_status, self.font_status = {}, {}, {}, {}
        for module, status in ((bl_links, self.bl_status), (link_repair, self.link_status), (hidden_repair, self.hidden_status),
                               (font_fix, self.font_status)):
            patcher = patch.object(module, "status", lambda status=status: dict(status))
            patcher.start()
            self.addCleanup(patcher.stop)
        bl_file = {"name": "a.xlsx", "path": "/x/a.xlsx", "hits": [{"sheet": "S", "addr": "B2"}], "error": None}
        self.bl_start = self.job(bl_links, "start", self.bl_status, "result",
                                 {"folder": "/x", "files": [bl_file, {"name": "b.xlsx", "path": "/x/b.xlsx", "hits": [], "error": "打不开"}]})
        self.bl_clear = self.job(bl_links, "clear", self.bl_status, "applied",
                                 {"folder": "/x", "files": [{**bl_file, "cells": [{"sheet": "S", "anchor": "B2", "status": "ok"}]}]})
        self.links = self.job(link_repair, "repair", self.link_status, "applied", {"folder": "/x", "files": [
            {"name": "a.xlsx", "path": "/x/a.xlsx", "error": None, "cells": [
                {"sheet": "S", "addr": "C3", "anchor": "C3", "status": "ok"},
                {"sheet": "S", "addr": "D4", "anchor": "D4", "status": "skip", "reason": "备份工作表无法唯一配对，请人工检查"}]}]})
        self.hidden = self.job(hidden_repair, "repair", self.hidden_status, "applied", {"folder": "/x", "files": [
            {"name": "a.xlsx", "path": "/x/a.xlsx", "error": None, "cells": [
                {"sheet": "S", "addr": "E:E", "anchor": "E1", "kind": "columns", "status": "ok"}]}]})

        self.font = self.job(font_fix, "start", self.font_status, "result", {"folder": "/x", "files": [
            {"name": "a.xlsx", "path": "/x/a.xlsx", "error": None, "shapes": 3, "cells": 0, "sheets": []},
            {"name": "c.xlsx", "path": "/x/c.xlsx", "error": None, "shapes": 0, "cells": 0, "sheets": []}]})

    def job(self, module, name, status, key, result):
        fake = FakeJob(module, key, result)
        fake.status = status
        patcher = patch.object(module, name, fake)
        patcher.start()
        self.addCleanup(patcher.stop)
        return fake

    def steps(self):
        return {step["key"]: step for step in clean_chain.status()["steps"]}

    def test_running_step_shows_its_own_file_progress(self):
        clean_chain.prepare("/x", {"hidden": False, "old_folder": "/ref"})
        seen = []
        reads = iter([{"running": True, "total": 4, "done": 2, "current": "c.xlsx"}])

        def status():
            current = next(reads, None)
            if current is None:
                seen.append(self.steps()["links"].get("progress"))
                return {"running": False, "applied": {"files": []}}
            return current

        with patch.object(link_repair, "status", status):
            clean_chain._update("links", state="running")
            clean_chain._wait(link_repair.status, "links", "扫描并修复")
        self.assertEqual(seen[0], {"phase": "扫描并修复", "done": 2, "total": 4, "current": "c.xlsx"})
        self.assertNotIn("progress", self.steps()["links"])  # 做完就收起

    def test_runs_bl_then_links_and_lists_what_could_not_be_fixed(self):
        clean_chain.prepare("/x", {"hidden": False, "old_folder": "/ref"})
        self.assertEqual(self.steps()["bl"]["state"], "waiting")
        clean_chain.run("/x", {"hidden": False, "old_folder": "/ref"})
        steps = self.steps()
        self.assertEqual(self.bl_start.calls, [("/x",)])
        self.assertEqual(len(self.bl_clear.calls), 1)
        self.assertEqual(steps["bl"]["fixed"], 1)
        self.assertEqual(steps["bl"]["state"], "failed")  # b.xlsx 打不开：记下来
        self.assertEqual(steps["bl"]["problems"][0]["file"], "b.xlsx")
        self.assertEqual(self.links.calls, [("/x", "/ref")])
        self.assertEqual((steps["links"]["state"], steps["links"]["fixed"], steps["links"]["problem_total"]), ("warn", 1, 1))
        self.assertEqual(steps["links"]["problems"][0]["cell"], "D4")
        self.assertIn("改不了 1 处", steps["links"]["text"])
        self.assertEqual(steps["hidden"]["state"], "off")
        self.assertEqual(self.hidden.calls, [])
        self.assertFalse(clean_chain.status()["running"])

    def test_problems_remember_the_paired_reference_file(self):
        self.links.result["files"][0].update(old="/ref/a.xlsx")
        self.links.result["files"][0]["cells"][1]["old_sheet"] = "S_old"
        clean_chain.run("/x", {"hidden": False, "old_folder": "/ref"})
        problem = self.steps()["links"]["problems"][0]
        self.assertEqual((problem["old"], problem["old_sheet"], problem["cell"]), ("/ref/a.xlsx", "S_old", "D4"))
        self.assertNotIn("old", self.steps()["bl"]["problems"][0])

    def test_hidden_runs_when_ticked_and_result_survives_restart(self):
        clean_chain.run("/x", {"hidden": True, "old_folder": "/ref"})
        self.assertEqual(self.hidden.calls, [("/x", "/ref")])
        self.assertEqual(self.steps()["hidden"]["state"], "done")
        with patch.object(clean_chain, "_STATUS", {}), patch.object(clean_chain, "_LOADED", False):
            saved = clean_chain.status()
            self.assertEqual(saved["folder"], "/x")
            self.assertEqual({s["key"]: s["fixed"] for s in saved["steps"]}, {"bl": 1, "links": 1, "hidden": 1, "font": 3})

    def test_font_runs_without_reference_folder(self):
        clean_chain.run("/x", {"hidden": False, "old_folder": ""})
        self.assertEqual(self.font.calls, [("/x",)])
        font = self.steps()["font"]
        self.assertEqual((font["state"], font["fixed"]), ("done", 3))
        self.assertIn("1 份 Excel 换成标准字体：3 个图形", font["text"])

    def test_without_reference_links_and_hidden_are_skipped(self):
        clean_chain.run("/x", {"hidden": True, "old_folder": ""})
        steps = self.steps()
        self.assertEqual(steps["links"]["state"], "skipped")
        self.assertEqual(steps["hidden"]["state"], "skipped")
        self.assertEqual(self.links.calls, [])

    def test_no_bl_hits_skips_clear(self):
        self.bl_start.result = {"folder": "/x", "files": [{"name": "a.xlsx", "path": "/x/a.xlsx", "hits": [], "error": None}]}
        clean_chain.run("/x", {"hidden": False, "old_folder": "/ref"})
        self.assertEqual(self.bl_clear.calls, [])
        self.assertEqual(self.steps()["bl"]["state"], "done")

    def test_a_failing_step_does_not_stop_the_next_one(self):
        self.bl_start.error = RuntimeError("Excel 崩了")
        clean_chain.run("/x", {"hidden": False, "old_folder": "/ref"})
        steps = self.steps()
        self.assertEqual(steps["bl"]["state"], "failed")
        self.assertIn("Excel 崩了", steps["bl"]["text"])
        self.assertEqual(steps["links"]["fixed"], 1)

    def test_busy_backend_is_waited_for_then_reported(self):
        self.bl_start.error = ValueError("后台任务正在执行，请等它结束")
        with patch.object(clean_chain, "BUSY_WAIT_SECONDS", 0):
            clean_chain.run("/x", {"hidden": False, "old_folder": "/ref"})
        self.assertEqual(self.steps()["bl"]["state"], "failed")
        self.assertIn("没能开始", self.steps()["bl"]["text"])

    def test_options_and_prepare_guard(self):
        self.assertIsNone(clean_chain.options_from(None))
        self.assertEqual(clean_chain.options_from({"hidden": True, "old_folder": ' "D:\\参考" '}),
                         {"hidden": True, "old_folder": "D:\\参考"})
        for bad in ("x", {"hidden": "yes"}, {"old_folder": 1}):
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
