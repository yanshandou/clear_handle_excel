import copy
import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import reporter, storage
from core.hidden_text import REASON_HASH


class ReportFixesTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.folder = self.root / "excel"
        (self.folder / "sub").mkdir(parents=True)
        for target, name, value in [(storage, "DATABASE_PATH", self.root / "t.sqlite3"),
                                    (reporter, "RECORD_PATH", self.root / "record.md")]:
            p = patch.object(target, name, value)
            p.start()
            self.addCleanup(p.stop)
        saved = copy.deepcopy(reporter.EXECUTION_STATE)
        self.addCleanup(lambda: (reporter.EXECUTION_STATE.clear(), reporter.EXECUTION_STATE.update(saved)))
        reporter.EXECUTION_STATE.update(folder="", files={}, steps=[], batch=[], started_at="", status="idle")
        from core import workflow
        self.workflow = importlib.reload(workflow)

    def item(self, iid, path, status, **extra):
        return {"id": iid, "seq": 1, "kind": "clip", "wb_path": str(path), "wb_name": "sub/画面.xlsx",
                "sheet_name": "基本情報", "excel_row": 12, "cell_addr": "C12", "status": status, **extra}

    def test_resync_fills_fixes_and_final_state_for_review_marked_file(self):
        original = self.folder / "sub" / "画面.xlsx"
        reporter.init_execution_batch("人工调整后导出复检（不自动加高）", self.folder, [original])
        reporter.finish_execution_batch()  # 旧版本留下的记录：停在「待处理」
        marked = self.folder / "sub" / "🟡画面.xlsx"
        self.workflow.STATE.update(folder=str(self.folder), items={
            "a": self.item("a", marked, "fixed"),
            "b": self.item("b", marked, "accepted"),
            "c": self.item("c", marked, "fixed", kind="hidden", hidden_reason=REASON_HASH, cell_addr="D3", excel_row=3),
        })
        self.workflow.resync_report_from_state()
        state = reporter.get_execution_report_snapshot()
        self.assertEqual(list(state["files"]), ["sub/画面.xlsx"])
        entry = state["files"]["sub/画面.xlsx"]
        self.assertEqual(entry["status"], "done")
        self.assertEqual(set(entry["fixes"]), {"a", "c"})
        md = state["markdown"]
        self.assertIn("✅ 已全部处理完成", md)
        self.assertIn("| 1 | 基本情報 | D3（第 3 行） | 列宽不够，数字显示成 ### |", md)
        self.assertIn("| 2 | 基本情報 | C12（第 12 行） | 文字压线，被边框切到（行高不够） |", md)

    def test_resync_settles_stale_queue_notes(self):
        a, b = self.folder / "sub" / "画面.xlsx", self.folder / "sub" / "無.xlsx"
        reporter.init_execution_batch("人工调整后导出复检（不自动加高）", self.folder, [a, b])
        reporter.finish_execution_batch()  # 两个文件都停在「排队等待处理...」
        self.workflow.STATE.update(folder=str(self.folder), checked_files={str(b): {}}, items={
            "a": self.item("a", a, "fixed"), "b": self.item("b", a, "accepted"),
        })
        self.workflow.resync_report_from_state()
        files = reporter.get_execution_report_snapshot()["files"]
        self.assertEqual((files["sub/画面.xlsx"]["status"], files["sub/画面.xlsx"]["current_action"]),
                         ("done", "已处理完：修复 1 处，1 处无需修改"))
        self.assertEqual((files["sub/無.xlsx"]["status"], files["sub/無.xlsx"]["current_action"]),
                         ("done", "检查完毕，无切字"))

    def test_workbench_does_not_write_into_other_folder_report(self):
        other = self.root / "other"
        other.mkdir()
        reporter.init_execution_batch("批量导出 PDF", other, [other / "画面.xlsx"])
        reporter.finish_execution_batch()
        path = self.folder / "sub" / "画面.xlsx"
        self.workflow.STATE.update(folder=str(self.folder), items={"a": self.item("a", path, "fixed")})
        self.workflow._sync_problem_counts(str(path), "x")
        entry = reporter.get_execution_report_snapshot()["files"]["画面.xlsx"]
        self.assertEqual(entry.get("clip_detected", 0), 0)


if __name__ == "__main__":
    unittest.main()
