import copy
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import fix_images, reporter, storage


class ReportAccumulateTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.folder = self.root / "excel"
        self.folder.mkdir()
        for name, value in [(storage, ("DATABASE_PATH", self.root / "t.sqlite3")),
                            (reporter, ("RECORD_PATH", self.root / "record.md")),
                            (fix_images, ("IMAGES_DIR", self.root / "_report_images"))]:
            p = patch.object(name, *value)
            p.start()
            self.addCleanup(p.stop)
        saved = copy.deepcopy(reporter.EXECUTION_STATE)
        self.addCleanup(lambda: (reporter.EXECUTION_STATE.clear(), reporter.EXECUTION_STATE.update(saved)))
        reporter.EXECUTION_STATE.update(folder="", files={}, steps=[], batch=[], started_at="")

    def test_same_folder_keeps_earlier_steps_and_other_folder_starts_over(self):
        a, b = self.folder / "a.xlsx", self.folder / "b.xlsx"
        reporter.init_execution_batch("精简 Sheet（删除备份表）", self.folder, [a, b])
        reporter.update_file_progress("a.xlsx", status="done", all_sheets=["S1", "S1 (2)"], deleted_backups=["S1 (2)"])
        reporter.update_file_progress("b.xlsx", status="done")
        reporter.finish_execution_batch()

        reporter.init_execution_batch("批量导出 PDF", self.folder, [a], pdf_folder=self.folder)
        reporter.update_file_progress("a.xlsx", status="done", exported_pdfs=1)
        reporter.finish_execution_batch()
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(state["files"]["a.xlsx"]["deleted_backups"], ["S1 (2)"])
        self.assertEqual(state["files"]["a.xlsx"]["exported_pdfs"], 1)
        self.assertIn("b.xlsx", state["files"])
        self.assertEqual((state["completed_files"], state["total_files"]), (1, 1))  # 进度只算这一步的文件
        self.assertIn("- ~~S1 (2)~~", (self.folder / "_执行记录报告.md").read_text(encoding="utf-8"))

        other = self.root / "other"
        other.mkdir()
        reporter.init_execution_batch("批量导出 PDF", other, [other / "c.xlsx"])
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(list(state["files"]), ["c.xlsx"])
        self.assertEqual(state["steps"], ["批量导出 PDF"])

    def test_progress_writes_are_batched_but_flushed_on_finish(self):
        a = self.folder / "a.xlsx"
        reporter.init_execution_batch("视图标准化（A1 + 100% 缩放）", self.folder, [a])
        with patch.object(reporter, "save_snapshot") as save:
            for _ in range(20):
                reporter.update_file_progress("a.xlsx", current_action="处理中")
            self.assertLessEqual(save.call_count, 1)
            reporter.finish_execution_batch()
            self.assertGreaterEqual(save.call_count, 1)
            self.assertIsNone(reporter._write_timer)

    def test_scan_after_clean_keeps_deleted_sheets(self):
        a = self.folder / "sub" / "a.xlsx"
        reporter.init_execution_batch("精简 Sheet（删除备份表）", self.folder, [a])
        reporter.update_file_progress("sub/a.xlsx", status="done", all_sheets=["履歴", "設定", "設定20220518", "情報"],
                                      deleted_backups=["設定20220518"])
        reporter.finish_execution_batch()
        # 切字扫描读到的工作表里已经没有删掉的表，也不知道前面删过什么
        reporter.init_execution_batch("Web 切字扫描（增量）", self.folder, [a])
        reporter.update_file_progress("sub/a.xlsx", all_sheets=["履歴", "設定", "情報"], deleted_backups=[])
        entry = reporter.get_execution_report_snapshot()["files"]["sub/a.xlsx"]
        self.assertEqual(entry["deleted_backups"], ["設定20220518"])
        self.assertEqual(entry["all_sheets"], ["履歴", "設定", "設定20220518", "情報"])

    def test_review_marked_name_updates_original_entry(self):
        a = self.folder / "sub" / "画面 .xlsx"
        reporter.init_execution_batch("Web 切字扫描（增量）", self.folder, [a])
        reporter.update_file_progress("sub/画面 .xlsx", status="needs_fix", clip_detected=4, clip_unfixed=4)
        reporter.finish_execution_batch()
        marked = self.folder / "sub" / "🟡画面 .xlsx"
        reporter.init_execution_batch("人工调整后导出复检（不自动加高）", self.folder, [marked], preserve_existing_files=True)
        reporter.update_file_progress("🟡画面 .xlsx", status="done", clip_detected=4, clip_fixed=4, clip_unfixed=0)
        reporter.finish_execution_batch()
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(list(state["files"]), ["sub/画面 .xlsx"])
        self.assertEqual(state["files"]["sub/画面 .xlsx"]["status"], "done")
        self.assertEqual((state["completed_files"], state["total_files"]), (1, 1))
        self.assertIn("✅ 已全部处理完成", state["markdown"])

    def test_restore_merges_old_review_marked_duplicates(self):
        reporter.EXECUTION_STATE.update(folder=str(self.folder), batch=["sub/a.xlsx", "🟡a.xlsx"], files={
            "sub/a.xlsx": {"index": 1, "name": "sub/a.xlsx", "path": str(self.folder / "sub" / "a.xlsx"), "status": "pending"},
            "🟡a.xlsx": {"index": 2, "name": "🟡a.xlsx", "path": "🟡a.xlsx", "status": "needs_fix"},
            "🟡b.xlsx": {"index": 3, "name": "🟡b.xlsx", "path": "🟡b.xlsx", "status": "done", "all_sheets": ["S"]},
            "🟡c.xlsx": {"index": 4, "name": "🟡c.xlsx", "path": "🟡c.xlsx", "status": "needs_fix", "all_sheets": []},
            "sub/Ｗeb会議.xlsx": {"index": 5, "name": "sub/Ｗeb会議.xlsx", "status": "pending",
                                 "current_action": "排队等待处理..."},
            "🟡Web会議.xlsx": {"index": 6, "name": "🟡Web会議.xlsx", "status": "done", "current_action": "已微调保存",
                              "normalized_a1": True, "fixes": {"i1": {"sheet": "S", "cell": "A1", "row": 1, "reason": "r"}}},
        })
        self.assertTrue(reporter._merge_review_duplicates_unlocked())
        files = reporter.EXECUTION_STATE["files"]
        # 有内容的改回原名；没内容的空壳去掉；全角半角不同也能并回原文件，且不丢后面步骤的结果
        self.assertEqual(sorted(files), ["b.xlsx", "sub/a.xlsx", "sub/Ｗeb会議.xlsx"])
        self.assertEqual(reporter.EXECUTION_STATE["batch"], ["sub/a.xlsx"])
        web = files["sub/Ｗeb会議.xlsx"]
        self.assertEqual((web["status"], web["current_action"], web["normalized_a1"]), ("done", "已微调保存", True))
        self.assertIn("i1", web["fixes"])

    def test_folder_change_clears_rows_even_when_preserving(self):
        other = self.root / "other"
        other.mkdir()
        reporter.init_execution_batch("批量导出 PDF", other, [other / "y.xlsx"])
        reporter.init_execution_batch("Web 闭环行高微调", self.folder, [self.folder / "a.xlsx"], preserve_existing_files=True)
        self.assertEqual(list(reporter.get_execution_report_snapshot()["files"]), ["a.xlsx"])

    def test_done_with_unfixed_clips_stays_needs_fix(self):
        a = self.folder / "a.xlsx"
        reporter.init_execution_batch("Web 切字扫描（增量）", self.folder, [a])
        reporter.update_file_progress("a.xlsx", status="needs_fix", clip_detected=2, clip_unfixed=2)
        reporter.init_execution_batch("视图标准化（A1 + 100% 缩放）", self.folder, [a])
        reporter.update_file_progress("a.xlsx", status="done", normalized_a1=True)
        self.assertEqual(reporter.get_execution_report_snapshot()["files"]["a.xlsx"]["status"], "needs_fix")

    def test_fixes_listed_with_place_and_reason_only(self):
        a = self.folder / "a.xlsx"
        reporter.init_execution_batch("Web 切字扫描（增量）", self.folder, [a])
        reporter.update_file_progress("a.xlsx", status="done", all_sheets=["基本情報"], clip_detected=3,
                                      clip_fixed=2, clip_accepted=1, clip_unfixed=0, fixes={
            "i2": {"sheet": "基本情報", "cell": "C12", "row": 12, "reason": "文字压线"},
            "i1": {"sheet": "基本情報", "cell": "B5", "row": 5, "reason": "列宽不够，数字显示成 ###"},
            "i3": None,
        })
        # 以后的扫描不再提到 i1（旧问题项不在工作台了），修复记录仍保留；i2 改回未修复则去掉
        reporter.update_file_progress("a.xlsx", fixes={"i2": None})
        reporter.update_file_progress("a.xlsx", fixes={"i2": {"sheet": "基本情報", "cell": "C12", "row": 12, "reason": "文字压线"}})
        reporter.finish_execution_batch()
        md = reporter.render_execution_markdown()
        self.assertIn("**② 修复切字：2 处**", md)
        self.assertLess(md.index("| 1 | 基本情報 | B5（第 5 行） | 列宽不够，数字显示成 ### |"),
                        md.index("| 2 | 基本情報 | C12（第 12 行） | 文字压线 |"))
        for word in ("人工", "自动", "跳过", "复检", "确认"):
            self.assertNotIn(word, md)

    def test_status_reflects_unfixed_problems(self):
        a, b = self.folder / "a.xlsx", self.folder / "b.xlsx"
        reporter.init_execution_batch("Web 切字扫描（增量）", self.folder, [a, b])
        reporter.update_file_progress("a.xlsx", status="needs_fix", clip_detected=2, clip_unfixed=1)
        reporter.update_file_progress("b.xlsx", status="done")
        reporter.finish_execution_batch()
        md = reporter.render_execution_markdown()
        self.assertIn("⚠️ 还有 1 个文件没处理完", md)
        self.assertIn("## 1. a.xlsx（还没处理完）", md)


if __name__ == "__main__":
    unittest.main()
