"""打印时内容特别少的页：量 PDF、判断、只修有把握的（造成空页的手动分页、最后几页全是空格子），别的标黄"""
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import page_breaks, page_fill, pdf_store
from core.pdf_detector import fitz
from tests.test_page_breaks import FakeSheet

ROW = 10.0      # 每行 10pt
MARGIN = 10.0   # 上下边距
PER_PAGE = 10   # 每页 10 行


class PrintSheet(FakeSheet):
    """在 test_page_breaks.FakeSheet 上加：打印区域（只认 A1:Hn 这种）、边距、带边框的空行、导出 PDF"""

    def __init__(self, rows, texts, bordered=(), merges=()):
        super().__init__(rows=rows, page=ROW * PER_PAGE, merges=merges, texts=texts)
        self.Name = "S1"
        self.bordered = set(bordered)
        self.PageSetup = Mock(PrintArea="", Zoom=100, FitToPagesTall=False, Order=1,
                              TopMargin=MARGIN, BottomMargin=MARGIN, PrintTitleRows="")
        self.UsedRange = Mock(Row=1, Column=1, Rows=Mock(Count=rows), Columns=Mock(Count=8))
        self.Shapes = Mock(Count=0)
        self.Application = Mock(ReferenceStyle=1)

    def last_row(self):
        area = str(self.PageSetup.PrintArea or "")
        return int(area.rsplit("$", 1)[-1]) if area else self.rows

    def breaks(self):
        keep, self.rows = self.rows, self.last_row()
        try:
            return super().breaks()
        finally:
            self.rows = keep

    def Range(self, a, b=None):
        if isinstance(a, str):
            bottom = int(a.rsplit("$", 1)[-1])
            return Mock(Row=1, Column=1, Rows=Mock(Count=bottom), Columns=Mock(Count=8))
        rng = super().Range(a, b)
        rng.Address = lambda *args: f"$A$1:$H${b.row}"
        rng.Find = Mock(side_effect=RuntimeError("no Find"))  # 走逐行找字
        return rng

    def export(self, _sheet, pdf_path):
        """每行 10pt 画到 PDF 上：有字的行写字，带边框的空行画一条线"""
        rows = list(range(1, self.last_row() + 1))
        starts = [1] + [row for row, _ in self.breaks()] + [self.last_row() + 1]
        doc = fitz.open()
        for i in range(len(starts) - 1):
            page = doc.new_page(width=200, height=ROW * PER_PAGE + 2 * MARGIN)
            for k, row in enumerate(range(starts[i], starts[i + 1])):
                y = MARGIN + k * ROW
                if row in rows and self.texts.get((row, 1)):
                    page.insert_text((5, y + 8), self.texts[(row, 1)], fontsize=7)
                elif row in self.bordered:
                    page.draw_line((5, y + ROW), (150, y + ROW), color=(0, 0, 0))
        doc.save(str(pdf_path))
        doc.close()


def texts(rows):
    return {(r, 1): f"row{r}" for r in rows}


class PageFillTests(unittest.TestCase):
    def setUp(self):
        self.pdf = Path(tempfile.mkdtemp()) / "S1.pdf"
        patcher = patch.object(page_fill.debug_log, "event")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_measure_uses_body_between_margins(self):
        sheet = PrintSheet(rows=25, texts=texts(range(1, 26)))
        sheet.export(sheet, self.pdf)
        pages = page_fill.measure(self.pdf, MARGIN, MARGIN)
        self.assertEqual(len(pages), 3)
        self.assertGreater(pages[0]["fill"], .9)
        self.assertAlmostEqual(pages[2]["fill"], .48, delta=.05)  # 5 行
        self.assertTrue(all(p["text"] for p in pages))

    def test_nothing_to_report_on_normal_sheet(self):
        sheet = PrintSheet(rows=25, texts=texts(range(1, 26)))
        sheet.export(sheet, self.pdf)
        self.assertIsNone(page_fill.inspect(sheet, self.pdf))

    def test_manual_break_making_one_row_page_is_removed(self):
        # 第 21 行一行单独一页（第 22 行上方有手动分页），前后页都满
        sheet = PrintSheet(rows=35, texts=texts(range(1, 36)))
        sheet.manual.add(22)
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([(it["kind"], it["page"], it.get("row")) for it in found["items"]], [("break", 2, 22)])

        notes, changed = page_fill.fix_sheet(sheet, self.pdf, sheet.export, found)
        self.assertTrue(changed)
        self.assertNotIn(22, sheet.manual)
        self.assertEqual(len(notes), 1)
        self.assertTrue(notes[0]["fixed"])
        self.assertIn("清掉了全部 1 处手动分页（第 22 行上方），改成自动分页，5 页 → 4 页", notes[0]["text"])
        with fitz.open(str(self.pdf)) as doc:
            self.assertEqual(len(doc), 4)

    def test_manual_break_kept_when_removing_it_cuts_merged_block(self):
        # 第 11 行单独一页；删掉第 12 行上方的分页后自动分页落在第 21 行，会切开 A19:C21
        sheet = PrintSheet(rows=35, texts=texts(range(1, 36)), merges=[(19, 1, 21, 3)])
        sheet.manual.add(12)
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([it["kind"] for it in found["items"]], ["break"])

        notes, changed = page_fill.fix_sheet(sheet, self.pdf, sheet.export, found)
        self.assertFalse(changed)
        self.assertIn(12, sheet.manual)  # 改回去了
        self.assertEqual(len(notes), 1)
        self.assertFalse(notes[0]["fixed"])
        self.assertEqual(notes[0]["page"], 1)
        self.assertIn("清掉全部手动分页会切开合并格 A19:C21", notes[0]["text"])
        with fitz.open(str(self.pdf)) as doc:
            self.assertEqual(len(doc), 5)  # PDF 也恢复成原来的样子

    def test_all_manual_breaks_are_cleared_at_once(self):
        # 用户 2026-11-08 的例子：一处空页是手动分页造成的，就连别的手动分页一起清掉，一次排好
        # 第 21 行单独一页（第 22 行上方手动分页），第 28 行上方还有一处手动分页让第 4 页只有 6 行
        sheet = PrintSheet(rows=45, texts=texts(range(1, 46)))
        sheet.manual.update({22, 28})
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([(it["kind"], it.get("row")) for it in found["items"]], [("break", 22)])

        notes, changed = page_fill.fix_sheet(sheet, self.pdf, sheet.export, found)
        self.assertTrue(changed)
        self.assertEqual(sheet.manual, set())
        self.assertEqual(len(notes), 1)
        self.assertIn("清掉了全部 2 处手动分页（第 22、28 行上方），改成自动分页，6 页 → 5 页", notes[0]["text"])
        self.assertTrue(notes[0]["cleared"])
        with fitz.open(str(self.pdf)) as doc:
            self.assertEqual(len(doc), 5)

    def test_all_manual_breaks_restored_when_clearing_cuts_merged_block(self):
        sheet = PrintSheet(rows=45, texts=texts(range(1, 46)), merges=[(29, 1, 31, 3)])
        sheet.manual.update({12, 28})
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([it["kind"] for it in found["items"]], ["break"])

        notes, changed = page_fill.fix_sheet(sheet, self.pdf, sheet.export, found)
        self.assertFalse(changed)
        self.assertEqual(sheet.manual, {12, 28})  # 全部改回去
        self.assertFalse(notes[0]["fixed"])
        self.assertIn("清掉全部手动分页会切开合并格 A29:C31", notes[0]["text"])

    def test_toolbox_counts_and_clears_manual_breaks(self):
        # 调整页面「清手动分页」：只数页高自动的表；清完不核对，交给人看新的 PDF（用户 2026-11-08 定）
        sheet = PrintSheet(rows=45, texts=texts(range(1, 46)))
        sheet.manual.update({22, 28})
        sheet.ResetAllPageBreaks = lambda: sheet.manual.clear()
        self.assertEqual(page_fill.sheet_layout(sheet)["manual"], 2)
        found = page_fill.clear_manual_breaks(sheet)
        self.assertEqual((found["ok"], found["count"]), (True, 2))
        self.assertEqual(sheet.manual, set())
        self.assertEqual(page_fill.sheet_layout(sheet)["manual"], 0)
        again = page_fill.clear_manual_breaks(sheet)
        self.assertFalse(again["ok"])
        self.assertIn("已经没有手动分页", again["reason"])

    def test_toolbox_falls_back_to_rows_when_reset_fails(self):
        sheet = PrintSheet(rows=45, texts=texts(range(1, 46)))
        sheet.manual.update({22, 28})
        sheet.ResetAllPageBreaks = Mock(side_effect=RuntimeError("不让改"))
        found = page_fill.clear_manual_breaks(sheet)
        self.assertTrue(found["ok"])
        self.assertEqual(sheet.manual, set())

    def test_fixed_tall_sheet_reports_no_manual_breaks(self):
        sheet = PrintSheet(rows=45, texts=texts(range(1, 46)))
        sheet.manual.add(22)
        sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesTall = False, 4
        self.assertEqual(page_fill.sheet_layout(sheet)["manual"], 0)

    def test_trailing_page_of_empty_bordered_rows_is_dropped_from_print_area(self):
        sheet = PrintSheet(rows=25, texts=texts(range(1, 21)), bordered=range(21, 26))
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([(it["kind"], it["page"], it["start"]) for it in found["items"]], [("tail", 2, 21)])

        notes, changed = page_fill.fix_sheet(sheet, self.pdf, sheet.export, found)
        self.assertTrue(changed)
        self.assertEqual(sheet.PageSetup.PrintArea, "$A$1:$H$20")
        self.assertIn("打印区域收到 A1:H20，3 页 → 2 页", notes[0]["text"])

    def test_trailing_page_with_shapes_is_only_flagged(self):
        sheet = PrintSheet(rows=25, texts=texts(range(1, 21)), bordered=range(21, 26))
        shape = Mock(BottomRightCell=Mock(Row=23))
        sheet.Shapes = Mock(Count=1, Item=lambda i: shape)
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        self.assertEqual([it["kind"] for it in found["items"]], ["flag"])
        self.assertIn("图片或图形", page_fill.flags(found)[0]["text"])

    def test_short_last_page_with_text_is_flagged(self):
        sheet = PrintSheet(rows=21, texts=texts(range(1, 22)))
        sheet.export(sheet, self.pdf)
        found = page_fill.inspect(sheet, self.pdf)
        notes = page_fill.flags(found)
        self.assertEqual([(n["page"], n["fixed"]) for n in notes], [(2, False)])
        self.assertIn("最后一页只印了约", notes[0]["text"])

    def test_fit_to_pages_tall_is_skipped(self):
        sheet = PrintSheet(rows=35, texts=texts(range(1, 36)))
        sheet.manual.add(22)
        sheet.PageSetup.Zoom, sheet.PageSetup.FitToPagesTall = False, 1
        sheet.export(sheet, self.pdf)
        self.assertIsNone(page_fill.inspect(sheet, self.pdf))

    def test_unreadable_pdf_is_skipped(self):
        self.pdf.write_bytes(b"not a pdf")
        self.assertIsNone(page_fill.inspect(PrintSheet(rows=5, texts={}), self.pdf))


class FixSparsePagesTests(unittest.TestCase):
    def test_locked_workbook_turns_fixes_into_yellow_notes(self):
        from core import workflow as wf
        found = {"pages": 5, "layout": {}, "items": [
            {"kind": "break", "page": 2, "row": 22, "fill": .05, "rows": "第 21～21 行"}]}
        with patch.object(wf, "workbook_locked", return_value=True), patch.object(wf, "log_msg") as log, \
                patch.object(wf.debug_log, "event"):
            notes = wf._fix_sparse_pages(Mock(), Path("a.xlsx"), {"x.pdf": ("S1", found)})
        self.assertEqual(notes["x.pdf"][0]["page"], 2)
        self.assertFalse(notes["x.pdf"][0]["fixed"])
        self.assertIn("没能改 Excel", notes["x.pdf"][0]["text"])
        self.assertTrue(any("正在 Excel 里打开" in str(c) for c in log.call_args_list))


class NotesStoreTests(unittest.TestCase):
    def test_notes_are_kept_with_export_record_until_pdf_changes(self):
        root = Path(tempfile.mkdtemp()).resolve()
        wb = root / "a.xlsx"
        wb.write_bytes(b"x")
        pdf = root / "PDFS" / "a_S1.pdf"
        pdf.parent.mkdir()
        pdf.write_bytes(b"%PDF")
        note = [{"page": 2, "fixed": False, "text": "第 3 页只印了约 5%"}]
        pdf_store.record(root, wb, "h", [(pdf, "S1")], {str(pdf): note})
        self.assertEqual(pdf_store.notes_map(root), {str(pdf): note})
        pdf.write_bytes(b"%PDF changed")
        self.assertEqual(pdf_store.notes_map(root), {})


if __name__ == "__main__":
    unittest.main()
