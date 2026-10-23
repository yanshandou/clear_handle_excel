import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class FakeCell:
    def __init__(self, addr, pattern=-4142, color=16777215, font=0):
        self.Address = addr
        self.Interior = Mock(Pattern=pattern, Color=color, TintAndShade=0)
        self.Font = Mock(Color=font)


class FakeSheet:
    def __init__(self):
        self.cells = {}
        self.merges = {}

    def cell(self, addr):
        key = "$" + addr[0] + "$" + addr[1:] if not addr.startswith("$") else addr
        return self.cells.setdefault(key, FakeCell(key))

    def Range(self, addr):
        cell = self.cell(addr)
        members = [self.cell(a) for a in self.merges.get(cell.Address, [cell.Address])]
        area = Mock(Cells=members)
        interior = Mock()

        def set_interior(name, value):
            for m in members:
                setattr(m.Interior, name, value)
        type(interior).Pattern = property(lambda _: members[0].Interior.Pattern,
                                          lambda _, v: set_interior("Pattern", v))
        type(interior).Color = property(lambda _: members[0].Interior.Color,
                                        lambda _, v: set_interior("Color", v))
        area.Interior = interior
        cell.MergeArea = area
        return cell


class ScanOnlyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, name, value in (
            (storage, "DATABASE_PATH", self.root / "state.sqlite3"),
            (reporter, "RECORD_PATH", self.root / "record.md"),
            (fix_images, "IMAGES_DIR", self.root / "images"),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.path = str(self.root / "book.xlsx")
        Path(self.path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []},
                             job={"kind": "batch_scan", "options": {}, "completed": [], "pending_adjustment": None})
        self.sheet = FakeSheet()
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.sheet
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        patch.object(batch, "normalize_workbook_view").start()
        patch.object(self.wf, "_sync_problem_counts").start()
        self.addCleanup(patch.stopall)

    def add_item(self, address, row=5, **extra):
        iid = f"i{address}"
        self.wf.STATE["items"][iid] = {
            "id": iid, "seq": len(self.wf.STATE["items"]) + 1, "kind": "hidden", "text": address,
            "wb_path": self.path, "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": address,
            "excel_row": row, "status": "pending", "image": "", **extra}
        return iid

    def scan(self, style=None):
        batch._data().setdefault("highlight_files", {})[self.path] = batch.highlight_style(style)
        batch._hand_over(self.path)
        batch._sync_highlights(lambda: self.excel, self.path)

    def card(self, address):
        return batch._data()["cards"][batch._key(self.path, "S1", address)]

    def test_all_problems_go_to_manual_and_cells_are_colored(self):
        self.add_item("A5")
        self.add_item("B7", row=7)
        self.scan()
        for addr in ("A5", "B7"):
            card = self.card(addr)
            self.assertEqual((card["color"], card["manual"], card["finalized"], card["waiting"]),
                             ("yellow", True, True, False))
            cell = self.sheet.cell(addr)
            self.assertEqual(cell.Interior.Pattern, batch.XL_PATTERN_SOLID)
            self.assertEqual(cell.Interior.Color, batch._excel_color("#C0007A"))
            self.assertEqual(cell.Font.Color, batch._excel_color("#FFFFFF"))
        self.assertEqual(self.wf.STATE["items"]["iA5"]["status"], "manual_review")
        self.book.Save.assert_called_once()
        # 行高没动：只扫描不写行高
        self.assertEqual(batch._data()["rows"], {})

    def test_custom_colors_and_validation(self):
        self.add_item("A5")
        self.scan({"fill": "#1d4ed8", "font": "#ffff00"})
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#1D4ED8"))
        self.assertEqual(self.sheet.cell("A5").Font.Color, 0x00FFFF)
        with self.assertRaises(ValueError):
            batch.highlight_style({"fill": "#000000", "font": "#000000"})
        self.assertEqual(batch.highlight_style({"fill": "red"}), batch.DEFAULT_HIGHLIGHT)

    def test_recheck_pass_restores_original_colors_but_keeps_user_changes(self):
        self.sheet.cells["$A$5"] = FakeCell("$A$5", pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        # 用户人工处理时自己把 B5 改成别的底色
        self.sheet.cell("B5").Interior.Color = 0x123456
        for addr in ("A5", "B5"):
            self.card(addr).update(color="blue", verified=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        a5, b5 = self.sheet.cell("A5"), self.sheet.cell("B5")
        self.assertEqual((a5.Interior.Pattern, a5.Interior.Color, a5.Font.Color), (1, 0xCCFFCC, 0x0000FF))
        self.assertEqual(b5.Interior.Color, 0x123456)
        self.assertEqual(b5.Font.Color, 0)  # 字色还是标记色时照常还原
        self.assertEqual(batch._marks(), {})

    def test_merged_area_and_mixed_font_color(self):
        self.sheet.merges["$A$5"] = ["$A$5", "$B$5"]
        self.sheet.cells["$A$5"] = FakeCell("$A$5", font=None)
        self.add_item("A5")
        self.scan()
        self.assertEqual(self.sheet.cell("B5").Interior.Color, batch._excel_color("#C0007A"))
        self.assertIsNone(self.sheet.cell("A5").Font.Color)  # 混合字色不动
        self.card("A5").update(dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.cell("B5").Interior.Pattern, batch.XL_PATTERN_NONE)

    def test_unconfirmed_location_is_not_colored(self):
        self.add_item("A5", location_confirmed=False)
        self.scan()
        self.assertIn("未在 Excel 中标色", self.card("A5")["message"])
        self.excel.Workbooks.Open.assert_not_called()

    def test_clear_restores_everything(self):
        self.add_item("A5")
        self.scan()
        batch._sync_highlights(lambda: self.excel, self.path, clear=True)
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, batch.XL_PATTERN_NONE)
        self.assertEqual(batch._marks(), {})

    def test_user_marked_cards_keep_their_color(self):
        self.add_item("A5")
        batch._card_for([self.wf.STATE["items"]["iA5"]]).update(color="green", marked=True)
        self.scan()
        self.assertEqual(self.card("A5")["color"], "green")
        self.excel.Workbooks.Open.assert_not_called()

    def test_run_scan_only_does_not_repair(self):
        self.add_item("A5")
        self.wf.STATE["job"]["options"] = {"folder": str(self.root), "scan_only": True,
                                           "highlight": {"fill": "#6D28D9", "font": "#FFFFFF"}}
        session = Mock()
        session.app.return_value = self.excel
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock(timed_out=False)), \
                patch.object(self.wf, "list_workbooks", return_value=(self.root, [Path(self.path)])), \
                patch.object(self.wf.pythoncom, "CoInitialize", create=True), \
                patch.object(self.wf.pythoncom, "CoUninitialize", create=True), \
                patch.object(batch, "init_execution_batch"), patch.object(batch, "finish_execution_batch"), \
                patch.object(batch, "_refresh"), patch.object(batch, "_repair_file") as repair:
            self.assertEqual(batch.run(), [self.path])
        repair.assert_not_called()
        self.assertEqual(self.card("A5")["color"], "yellow")
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#6D28D9"))
        self.assertIn(self.path, self.wf.STATE["job"]["completed"])


if __name__ == "__main__":
    unittest.main()
