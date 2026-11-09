"""调整页面里的切字（用户 2026-11-09 把查看切字并进调整页面）"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import adjust_cells, batch, cell_toolbox, pdf_store, workflow as wf


class BoxesTests(unittest.TestCase):
    def test_box_is_fraction_of_page(self):
        pages = {"p1": {"page_width": 200.0, "page_height": 400.0, "total_pages": 3}}
        items = [{"bbox": [20, 40, 120, 80], "page_id": "p1", "pdf_path": "/x/a.pdf", "page_index": 1},
                 {"bbox": [20, 40, 120, 80], "page_id": "p1", "pdf_path": "/x/a.pdf", "page_index": 1},  # 同一块只画一次
                 {"bbox": [5, 5, 5, 9], "page_id": "p1", "pdf_path": "/x/a.pdf", "page_index": 1},       # 没有宽度
                 {"bbox": None, "page_id": "p1", "pdf_path": "/x/a.pdf"}]
        with patch.dict(wf.STATE, {"pages": pages}):
            boxes = adjust_cells._boxes(items, True)
            self.assertEqual(adjust_cells._boxes(items, False), [])
        self.assertEqual(len(boxes), 1)
        self.assertEqual(boxes[0]["page"], 1)
        self.assertEqual(boxes[0]["pages"], 3)
        self.assertEqual(boxes[0]["box"], [0.1, 0.1, 0.6, 0.2])


class ValidateTests(unittest.TestCase):
    def test_plan_and_skips(self):
        plan, skips = adjust_cells.validate(
            [{"card_id": "a", "tool": "grow", "ratio": 1.5}, {"card_id": "b", "tool": "shrink"},
             {"card_id": "c", "tool": "break"}],
            ["c", "c", "d"])
        self.assertEqual(plan, {"a": {"tool": "grow", "ratio": 1.5}, "b": {"tool": "shrink"}})
        self.assertEqual(skips, ["c", "d"])

    def test_bad_input(self):
        with self.assertRaises(ValueError):
            adjust_cells.validate([{"card_id": "a", "tool": "grow", "ratio": 3}], [])
        with self.assertRaises(ValueError):
            adjust_cells.validate([{"card_id": "a", "tool": "delete"}], [])
        with self.assertRaises(ValueError):
            adjust_cells.validate([], [1])


class RunStepsTests(unittest.TestCase):
    def test_order_skip_then_rescan_then_fix(self):
        calls = []
        cards = {"a": {"id": "a"}, "b": {"id": "b"}, "s": {"id": "s", "dismissed": True}}

        def run_job(kind, options):
            calls.append(kind)
            if kind == "batch_fix":
                return {"result": {"/r/x.xlsx": {"name": "x.xlsx", "done": 1, "passed": 1}}}
            return {"status": "done"}

        logs = []
        with patch.object(wf, "activate_folder"), \
                patch.dict(wf.STATE, {"checked_files": {"/r/x.xlsx": "v1"}}), \
                patch.object(adjust_cells, "_live", side_effect=lambda ids: [i for i in ids if i != "b"]), \
                patch.object(adjust_cells, "needs_check", return_value=["/r/x.xlsx", "/r/new.xlsx"]), \
                patch.object(adjust_cells, "run_job", side_effect=run_job), \
                patch.object(batch, "mark_many", side_effect=lambda ids, color: calls.append(("mark", color))), \
                patch.object(batch, "_data", return_value={"cards": cards}), \
                patch.object(cell_toolbox, "allowed_tools", return_value=["grow", "shrink"]), \
                patch.object(cell_toolbox, "job_options", side_effect=lambda todo: {"cells": todo}) as options:
            out = adjust_cells.run_steps("/r", {"a": {"tool": "grow", "ratio": 1.5}, "b": {"tool": "shrink"}}, ["s"],
                                         None, lambda text: None, lambda rows, ok, kind: logs.append(rows[0]["name"]),
                                         lambda stage: calls.append(("stage", stage)))
        self.assertEqual(calls, [("stage", "skip"), ("mark", "blue"), "batch_commit",
                                 ("stage", "scan"), "batch_scan", ("stage", "cells"), "batch_fix"])
        options.assert_called_once_with({"a": {"tool": "grow", "ratio": 1.5}})  # b 重查以后没了，不改
        self.assertEqual(out["skipped"], 1)
        self.assertEqual(out["rechecked"], 1)  # 没查过的新文件不算「改过要重查」
        self.assertIn("1 格没改", logs)


class FollowSaveTests(unittest.TestCase):
    def test_record_follows_program_save(self):
        root = Path(tempfile.mkdtemp()).resolve()
        wb = root / "spec.xlsx"
        wb.write_bytes(b"x")
        pdf = root / "PDFS" / "spec_A.pdf"
        pdf.parent.mkdir(parents=True)
        pdf.write_bytes(b"%PDF")
        pdf_store.record(root, wb, "h1", [(pdf, "A")])
        self.assertTrue(pdf_store.has_record(root, wb))
        self.assertFalse(pdf_store.follow_save(root, wb, "other", "h2"))  # 改之前就旧了：不跟
        self.assertIsNone(pdf_store.reusable(root, wb, "h2"))
        self.assertTrue(pdf_store.follow_save(root, wb, "h1", "h2"))
        self.assertEqual(pdf_store.reusable(root, wb, "h2"), [(pdf, "A")])
        self.assertFalse(pdf_store.has_record(root, root / "none.xlsx"))


if __name__ == "__main__":
    unittest.main()
