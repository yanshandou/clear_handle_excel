import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, page_breaks, reporter, storage, workflow
from core.pdf_detector import fitz

MANUAL, NONE, AUTO = page_breaks.XL_PAGEBREAK_MANUAL, page_breaks.XL_PAGEBREAK_NONE, -4105


class FakeArea:
    def __init__(self, sheet, top, left, bottom, right):
        self.sheet, self.Row, self.Column = sheet, top, left
        self.Rows = Mock(Count=bottom - top + 1)
        self.Columns = Mock(Count=right - left + 1)

    def Cells(self, r, c):
        return self.sheet.Cells(self.Row + r - 1, self.Column + c - 1)


class FakeCell:
    def __init__(self, sheet, row, col):
        self.sheet, self.row, self.col = sheet, row, col

    @property
    def Text(self):
        return self.sheet.texts.get((self.row, self.col), "")

    def _area(self):
        for top, left, bottom, right in self.sheet.merges:
            if top <= self.row <= bottom and left <= self.col <= right:
                return top, left, bottom, right
        return None

    @property
    def MergeCells(self):
        return self._area() is not None

    @property
    def MergeArea(self):
        area = self._area() or (self.row, self.col, self.row, self.col)
        return FakeArea(self.sheet, *area)


class FakeRow:
    def __init__(self, sheet, row):
        self.sheet, self.row = sheet, row

    @property
    def PageBreak(self):
        return MANUAL if self.row in self.sheet.manual else AUTO if self.row in self.sheet.auto_rows() else NONE

    @PageBreak.setter
    def PageBreak(self, value):
        if value == MANUAL:
            self.sheet.manual.add(self.row)
        else:
            self.sheet.manual.discard(self.row)


class FakeSheet:
    """每行 10pt、每页 page 行高；手动分页另起一页。merges: [(top, left, bottom, right)]"""

    def __init__(self, rows=60, page=100.0, merges=(), texts=None):
        self.rows, self.page = rows, page
        self.merges = list(merges)
        self.texts = dict(texts or {})
        self.manual: set[int] = set()
        self.PageSetup = Mock(PrintArea="", Zoom=100, FitToPagesTall=False, Order=1)
        self.UsedRange = Mock(Row=1, Column=1, Columns=Mock(Count=8))
        self.Parent = Mock()

    def breaks(self):
        result, used = [], 0.0
        for row in range(1, self.rows + 1):
            if row in self.manual and row > 1:
                result.append((row, True))
                used = 0.0
            elif used + 10 > self.page:
                result.append((row, False))
                used = 0.0
            used += 10
        return result

    def auto_rows(self):
        return {row for row, manual in self.breaks() if not manual}

    @property
    def HPageBreaks(self):
        found = self.breaks()
        pb = Mock(side_effect=lambda i: Mock(Location=Mock(Row=found[i - 1][0]),
                                              Type=MANUAL if found[i - 1][1] else AUTO))
        pb.Count = len(found)
        return pb

    @property
    def VPageBreaks(self):
        return Mock(Count=0)

    def Activate(self):
        pass

    def Cells(self, row, col):
        return FakeCell(self, row, col)

    def Rows(self, row):
        return FakeRow(self, row)

    def Range(self, a, b=None):
        cells = [(r, c) for r in range(a.row, b.row + 1) for c in range(a.col, b.col + 1)]
        merged = [FakeCell(self, r, c).MergeCells for r, c in cells]
        return Mock(MergeCells=all(merged) if all(merged) else (None if any(merged) else False))


class DetectTests(unittest.TestCase):
    def test_finds_text_block_cut_by_automatic_break_only(self):
        # 每页 10 行：自动分页在第 11 行上方；B9:F12 有文字被切开，H9:H12 空的不算
        sheet = FakeSheet(merges=[(9, 2, 12, 6), (9, 8, 12, 8), (3, 1, 4, 1)], texts={(9, 2): "見出し"})
        breaks = page_breaks.read_breaks(sheet)
        self.assertEqual(breaks["h"][0], (11, False))
        blocks = page_breaks.find_cut_blocks(sheet, breaks)
        self.assertEqual([b["addr"] for b in blocks], ["B9:F12"])
        self.assertEqual((blocks[0]["top"], blocks[0]["bottom"], blocks[0]["break_row"]), (9, 12, 11))

    def test_manual_break_inside_block_is_authors_choice(self):
        sheet = FakeSheet(merges=[(5, 1, 7, 3)], texts={(5, 1): "表題"})
        sheet.manual.add(6)
        self.assertEqual(page_breaks.find_cut_blocks(sheet, page_breaks.read_breaks(sheet)), [])

    def test_hidden_block_is_not_reported(self):
        sheet = FakeSheet(merges=[(9, 2, 12, 6)], texts={(9, 2): "見出し"})
        breaks = page_breaks.read_breaks(sheet)
        self.assertEqual(page_breaks.find_cut_blocks(sheet, breaks, lambda r, c: False), [])

    def test_unreadable_breaks_skip_detection(self):
        sheet = Mock()
        sheet.HPageBreaks.Count = Mock(side_effect=RuntimeError)  # int(Mock) 失败
        self.assertIsNone(page_breaks.read_breaks(sheet))

    def test_page_index_follows_print_order_and_refuses_mismatched_page_counts(self):
        sheet = FakeSheet()
        breaks = {"h": [(11, False), (21, False)], "v": [5]}
        self.assertEqual(page_breaks.page_index(sheet, breaks, 15, 6, 6), 4)  # 先列后行：第 2 列第 2 段
        sheet.PageSetup.Order = page_breaks.XL_OVER_THEN_DOWN
        self.assertEqual(page_breaks.page_index(sheet, breaks, 15, 6, 6), 3)
        self.assertIsNone(page_breaks.page_index(sheet, breaks, 15, 6, 5))  # 空白页被跳过时不猜

    def test_render_break_joins_two_pages_and_finds_text(self):
        doc = fitz.open()
        for text in ("上一页", "見出しテキスト"):
            page = doc.new_page(width=300, height=400)
            page.insert_text((40, 360 if text == "上一页" else 40), text, fontname="japan", fontsize=12)
        image, found_page, bbox = page_breaks.render_break(doc, 0, 1, "見出しテキスト")
        self.assertTrue(image.startswith("data:image/png;base64,"))
        self.assertEqual(found_page, 1)
        self.assertEqual(len(bbox), 4)


class PrepareTests(unittest.TestCase):
    def session_for(self, sheet):
        session = Mock()
        session.copy_of.return_value.Sheets.return_value = sheet
        session.chosen_pdf.return_value = Path(tempfile.mkdtemp()) / "chosen.pdf"
        return session

    def item(self, top, bottom, left=2):
        return {"sheet_name": "S1", "excel_row": top, "text": "見出し",
                "block": {"top": top, "bottom": bottom, "left": left, "right": 6, "break_row": 11}}

    def test_inserts_break_checks_and_restores_copy(self):
        sheet = FakeSheet(merges=[(9, 2, 12, 6)], texts={(9, 2): "見出し"})
        exported = []
        with patch.object(batch.wf, "_export_sheet_pdf", side_effect=lambda s, pdf: exported.append(set(s.manual))), \
                patch.object(batch, "fitz") as fake_fitz, patch.object(page_breaks, "shot_around", return_value=("img", None, [])):
            fake_fitz.open.return_value.__enter__.return_value = Mock()
            chosen = batch._prepare_break(self.session_for(sheet), "x.xlsx", self.item(9, 12), {})
        self.assertEqual(exported, [{9}])  # 导出时分页已插在第 9 行上方
        self.assertEqual(chosen["after"], "img")
        self.assertEqual(sheet.manual, set())  # 副本改回去

    def test_block_taller_than_a_page_goes_manual(self):
        sheet = FakeSheet(merges=[(5, 2, 20, 6)], texts={(5, 2): "長い"})
        with self.assertRaisesRegex(ValueError, "比一页还高"):
            batch._prepare_break(self.session_for(sheet), "x.xlsx", self.item(5, 20), {})
        self.assertEqual(sheet.manual, set())

    def test_fit_to_pages_tall_is_not_touched(self):
        sheet = FakeSheet()
        sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesTall = False, 2
        with self.assertRaisesRegex(ValueError, "页高"):
            batch._prepare_break(self.session_for(sheet), "x.xlsx", self.item(9, 12), {})
        self.assertEqual(sheet.manual, set())

    def test_finds_blank_row_above_and_places_break_between_blank_and_text_rows(self):
        # 内容在 6..12 行被切开，第 4 行空白、第 5 行开始有字 → 分页插在第 4 行和第 5 行之间（第 5 行上方）
        sheet = FakeSheet(merges=[(6, 2, 12, 6)], texts={(2, 2): "前段", (5, 2): "第5行标题", (6, 2): "第6行内容"})
        self.assertEqual(page_breaks.find_break_row(sheet, {"top": 6, "bottom": 12}), 5)
        exported = []
        with patch.object(batch.wf, "_export_sheet_pdf", side_effect=lambda s, pdf: exported.append(set(s.manual))), \
                patch.object(batch, "fitz") as fake_fitz, patch.object(page_breaks, "shot_around", return_value=("img", None, [])):
            fake_fitz.open.return_value.__enter__.return_value = Mock()
            chosen = batch._prepare_break(self.session_for(sheet), "x.xlsx", self.item(6, 12), {})
        self.assertEqual(exported, [{5}])
        self.assertEqual(chosen["row"], 5)
        self.assertEqual(sheet.manual, set())

    def test_searches_upward_across_multiple_text_rows_or_vertical_merges(self):
        # 第 3 行空白，第 4、5 行都有字（或纵向合并跨过 4..6 行），内容在 6..11 行 → 分页插在第 4 行上方
        sheet = FakeSheet(merges=[(6, 2, 11, 6)], texts={(2, 2): "前段", (4, 2): "大標題", (5, 2): "小標題", (6, 2): "本文"})
        self.assertEqual(page_breaks.find_break_row(sheet, {"top": 6, "bottom": 11}), 4)
        # A4:A6 为纵向合并单元格，B4 虽无字但 A4:A6 跨过 4-5 分界线，不能在 5 上方切开 A4:A6，应再往上到 4 上方
        merged_side = FakeSheet(merges=[(4, 1, 6, 1), (6, 2, 11, 6)],
                                texts={(2, 2): "前段", (4, 1): "分類", (5, 2): "項目", (6, 2): "本文"})
        self.assertEqual(page_breaks.find_break_row(merged_side, {"top": 6, "bottom": 11}), 4)

    def test_falls_back_to_block_top_when_no_blank_row_or_upward_break_still_cuts_block(self):
        # 当前页 1..8 行全有字（无空白行）→ 退回在合并块顶部第 9 行上方插分页
        no_blank = FakeSheet(merges=[(9, 2, 12, 6)], texts={(r, 2): f"行{r}" for r in range(1, 10)})
        self.assertEqual(page_breaks.find_break_row(no_blank, {"top": 9, "bottom": 12}), 9)
        # 第 3 行空白、第 4..8 行有字，但从第 4 行到第 14 行共 11 行（110pt > 单页 100pt），
        # 插在第 4 行上方会导致 9..14 仍被下一页自动分页切开 → 自动退回插在第 9 行上方（9..14 共 60pt 可放下）
        too_tall_from_blank = FakeSheet(
            merges=[(9, 2, 14, 6)],
            texts={(2, 2): "前段", **{((r, 2)): f"行{r}" for r in range(4, 10)}},
        )
        self.assertEqual(page_breaks.find_break_row(too_tall_from_blank, {"top": 9, "bottom": 14}), 4)
        exported = []
        with patch.object(batch.wf, "_export_sheet_pdf", side_effect=lambda s, pdf: exported.append(set(s.manual))), \
                patch.object(batch, "fitz") as fake_fitz, patch.object(page_breaks, "shot_around", return_value=("img", None, [])):
            fake_fitz.open.return_value.__enter__.return_value = Mock()
            chosen = batch._prepare_break(self.session_for(too_tall_from_blank), "x.xlsx", self.item(9, 14), {})
        self.assertEqual(exported, [{9}])
        self.assertEqual(chosen["row"], 9)
        self.assertEqual(too_tall_from_blank.manual, set())


class BatchBreakTests(unittest.TestCase):
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
        self.original = FakeSheet(merges=[(9, 2, 12, 6)], texts={(9, 2): "見出し"})
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.original
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        patch.object(batch, "normalize_workbook_view").start()
        self.refresh = patch.object(batch, "_refresh").start()
        self.addCleanup(patch.stopall)
        self.item = {"id": "b1", "seq": 1, "kind": "pagebreak", "text": "見出し", "wb_path": self.path,
                     "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": "B9:F12", "excel_row": 9,
                     "block": {"top": 9, "bottom": 12, "left": 2, "right": 6, "break_row": 11},
                     "location_confirmed": True, "status": "pending", "image": "", "page_index": 1,
                     "bbox": [10, 10, 100, 25]}
        self.wf.STATE["items"]["b1"] = self.item

    def repair(self):
        session = Mock()
        with patch.object(batch, "_prepare_break", return_value={"after": "data:image/png;base64,YWZ0ZXI=", "pdf": None}):
            batch._repair_file(self.excel, session, self.path)
        return session

    def test_break_is_written_reported_and_reverted_by_green(self):
        session = self.repair()
        self.assertEqual(self.original.manual, {9})
        self.book.Save.assert_called_once()
        session.written_with.assert_called_once()
        card = batch._data()["cards"][batch._key(self.path, "S1", "B9:F12")]
        self.assertEqual((card["color"], card["verified"]), ("blue", True))
        self.assertEqual(self.item["status"], "fixed")
        self.assertNotIn("row_height", self.item)  # 插分页不是改行高
        self.assertEqual(self.wf._fix_record(self.item)["reason"], page_breaks.FIX_REASON)
        snap = batch.snapshot()["cards"][0]
        self.assertEqual(snap["page_break"], {"row": 9, "saved": True, "applied": True})
        self.assertNotIn("row_height", snap)
        self.assertEqual(snap["problems"][0]["label"], "分页切开")

        batch.mark(card["id"], "green")
        self.wf.STATE["job"].update(kind="batch_commit", options={"card_ids": [card["id"]]}, completed=[])
        batch._commit_review(self.excel)
        self.assertEqual(self.original.manual, set())
        self.assertTrue(card["dismissed"])

    def test_resume_does_not_insert_twice_and_user_break_is_not_overwritten(self):
        rid = batch._key(self.path, "S1", 9, "break")
        journal = {"path": self.path, "version": "old", "rows": [
            {"id": rid, "kind": "break", "sheet": "S1", "row": 9, "expected": 0.0, "target": 1.0}]}
        self.original.manual.add(9)
        batch._write_rows(self.excel, journal)  # 已到位：文件变过也不再保存
        self.book.Save.assert_not_called()
        restore = {**journal, "version": self.wf._file_version(Path(self.path)), "rows": [
            {"id": rid, "kind": "break", "sheet": "S1", "row": 9, "expected": 1.0, "target": 0.0}]}
        self.original.manual.clear()
        self.original.manual.add(9)
        batch._write_rows(self.excel, restore)
        self.assertEqual(self.original.manual, set())

    def test_break_and_row_height_on_same_row_are_separate_changes(self):
        hidden = {**self.item, "id": "h1", "seq": 2, "kind": "hidden", "cell_addr": "B9", "block": None}
        self.wf.STATE["items"]["h1"] = hidden
        order = []
        def prepare_row(_s, _p, items, **_):
            order.append(items[0]["kind"])
            return 10.0, None, {}
        with patch.object(batch, "_prepare_break", side_effect=lambda *a, **k: order.append("pagebreak") or
                          {"after": "", "pdf": None}), patch.object(batch, "_prepare_row", side_effect=prepare_row):
            batch._repair_file(self.excel, Mock(), self.path)
        self.assertEqual(order, ["pagebreak", "hidden"])  # 同一行先插分页
        self.assertNotEqual(batch._rid(self.item), batch._rid(hidden))

    def test_failure_goes_yellow_manual_without_touching_original(self):
        with patch.object(batch, "_prepare_break", side_effect=ValueError(page_breaks.TALL_REASON)):
            batch._repair_file(self.excel, Mock(), self.path)
        card = batch._data()["cards"][batch._key(self.path, "S1", "B9:F12")]
        self.assertEqual((card["color"], card["manual"]), ("yellow", True))
        self.assertIn("比一页还高", card["message"])
        self.assertEqual(self.original.manual, set())
        self.book.Save.assert_not_called()

    def test_upward_break_row_is_written_renamed_and_reverted_correctly(self):
        # 内容在第 6 行（B6:F12），向上找到第 4 行空白、第 5 行有字，分页插在第 5 行上方
        self.item.update(cell_addr="B6:F12", excel_row=6,
                         block={"top": 6, "bottom": 12, "left": 2, "right": 6, "break_row": 11})
        session = Mock()
        with patch.object(batch, "_prepare_break", return_value={
            "after": "data:image/png;base64,YWZ0ZXI=", "pdf": None, "row": 5,
        }):
            batch._repair_file(self.excel, session, self.path)
        self.assertEqual(self.original.manual, {5})
        card = batch._data()["cards"][batch._key(self.path, "S1", "B6:F12")]
        self.assertIn("第 5 行上方插入分页", card["message"])
        snap = batch.snapshot()["cards"][0]
        self.assertEqual(snap["page_break"], {"row": 5, "saved": True, "applied": True})
        # 工作表改名后仍能对上卡片的 excel_row=6，且标绿提交能把第 5 行上方的分页还原去掉
        self.item["sheet_name"] = "S2"
        batch.rename_sheet(self.path, "S1", "S2")
        self.assertEqual(batch._row_record(card)["row"], 5)
        batch.mark(card["id"], "green")
        self.wf.STATE["job"].update(kind="batch_commit", options={"card_ids": [card["id"]]}, completed=[])
        batch._commit_review(self.excel)
        self.assertEqual(self.original.manual, set())
        self.assertTrue(card["dismissed"])


if __name__ == "__main__":
    unittest.main()
