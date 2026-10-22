import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

from core import excel_ops, sheet_preview
from core.sheet_preview import find_similar_sheets, preview_image_path, read_sheet_content


def fake_sheet(values, row=1, column=1, rows=None, columns=None):
    if isinstance(values, tuple) and values and isinstance(values[0], tuple):
        rows = rows or len(values)
        columns = columns or len(values[0])
    used = SimpleNamespace(
        Row=row, Column=column, Value=values,
        Rows=SimpleNamespace(Count=rows or 1), Columns=SimpleNamespace(Count=columns or 1),
    )
    return SimpleNamespace(UsedRange=used)


class ReadSheetContentTests(unittest.TestCase):
    def test_reads_cells_with_absolute_positions_and_normalized_numbers(self):
        content = read_sheet_content(fake_sheet((("名称", 1.0), (None, " x ")), row=3, column=2))
        self.assertEqual(content["cells"], {"3,2": "名称", "3,3": "1", "4,3": "x"})
        self.assertEqual(content["rows"], ["名称\t1", "\tx"])

    def test_single_cell_and_empty_sheet(self):
        self.assertEqual(read_sheet_content(fake_sheet("abc"))["rows"], ["abc"])
        self.assertEqual(read_sheet_content(fake_sheet(None)), {"cells": {}, "rows": []})

    def test_chart_sheet_or_huge_range_is_skipped(self):
        self.assertIsNone(read_sheet_content(SimpleNamespace()))
        self.assertIsNone(read_sheet_content(fake_sheet("a", rows=10_000, columns=100)))


class FindSimilarSheetsTests(unittest.TestCase):
    def content(self, rows, offset=1):
        cells = {}
        for r, row in enumerate(rows):
            for c, value in enumerate(row.split("\t")):
                if value:
                    cells[f"{r + offset},{c + 1}"] = value
        return {"cells": cells, "rows": rows}

    def test_identical_sheets_point_to_each_other(self):
        rows = ["a\t1", "b\t2", "c\t3"]
        result = find_similar_sheets({"一覧": self.content(rows), "一覧(2)": self.content(rows)})
        self.assertEqual(result["一覧"], {"match": "一覧(2)", "ratio": 1.0, "identical": True})
        self.assertEqual(result["一覧(2)"]["match"], "一覧")

    def test_inserted_row_is_similar_but_not_identical(self):
        base = [f"row{i}" for i in range(20)]
        result = find_similar_sheets({
            "A": self.content(base),
            "B": self.content(base[:5] + ["new"] + base[5:]),
        })
        self.assertFalse(result["A"]["identical"])
        self.assertGreaterEqual(result["A"]["ratio"], 0.9)
        self.assertLess(result["A"]["ratio"], 1.0)

    def test_edited_cells_in_every_row_still_found_as_old_version(self):
        # 设计书改版：每行都改了一个单元格，按行比几乎没有相同行，按单元格比仍然大部分一致
        old = [f"項目{i}\t説明{i}\t旧" for i in range(10)]
        new = [f"項目{i}\t説明{i}\t新" for i in range(10)]
        result = find_similar_sheets({"設計": self.content(new), "設計_20240105": self.content(old)})
        self.assertEqual(result["設計_20240105"]["match"], "設計")
        self.assertFalse(result["設計_20240105"]["identical"])
        self.assertEqual(result["設計_20240105"]["ratio"], 0.67)  # 30 个单元格里 20 个一致

    def test_same_rows_at_other_position_are_not_identical(self):
        rows = ["a", "b"]
        result = find_similar_sheets({"A": self.content(rows), "B": self.content(rows, offset=3)})
        self.assertFalse(result["A"]["identical"])
        self.assertEqual(result["A"]["ratio"], 0.99)

    def test_dissimilar_empty_and_unreadable_sheets_are_not_reported(self):
        result = find_similar_sheets({
            "A": self.content(["a", "b", "c", "d"]),
            "B": self.content(["a", "x", "y", "z"]),
            "Empty": self.content([]),
            "Chart": None,
        })
        self.assertEqual(result, {})

    def test_prefers_identical_match(self):
        rows = [f"r{i}" for i in range(10)]
        result = find_similar_sheets({
            "A": self.content(rows),
            "B": self.content(rows[:9] + ["x"]),
            "C": self.content(rows),
        })
        self.assertEqual(result["A"]["match"], "C")


class PreviewImagePathTests(unittest.TestCase):
    def test_only_files_inside_cache_with_expected_names(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(sheet_preview, "PREVIEW_CACHE_DIR", Path(tmp)):
            key = "0123456789abcdef"
            (Path(tmp) / key).mkdir()
            (Path(tmp) / key / "0.png").write_bytes(b"png")
            self.assertEqual(preview_image_path(key, "0.png"), Path(tmp) / key / "0.png")
            self.assertIsNone(preview_image_path(key, "1.png"))
            self.assertIsNone(preview_image_path("..", "0.png"))
            self.assertIsNone(preview_image_path(key, "../manifest.json"))
            self.assertIsNone(preview_image_path(key, "manifest.json"))



class RenderSheet:
    def __init__(self, name, values, visible=-1, printable=True):
        self.Name = name
        self.Visible = visible
        self.printable = printable
        self.UsedRange = fake_sheet(values).UsedRange

    def ExportAsFixedFormat(self, **kwargs):
        if not self.printable:
            raise RuntimeError("nothing to print")
        doc = sheet_preview.fitz.open()
        for _ in range(2):
            doc.new_page().insert_text((72, 72), self.Name)
        doc.save(kwargs["Filename"])
        doc.close()


class RenderWorkbookPreviewsTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.folder = self.root / "excel"
        self.folder.mkdir()
        (self.folder / "spec.xlsx").write_bytes(b"x")
        rows = (("a", 1), ("b", 2))
        self.book = Mock()
        self.book.Sheets = [
            RenderSheet("Main", rows),
            RenderSheet("Main(2)", rows, visible=0),
            RenderSheet("Blank", None, printable=False),
        ]
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        for name, value in [
            ("PREVIEW_CACHE_DIR", self.root / "cache"),
            ("pythoncom", Mock()),
        ]:
            patcher = patch.object(sheet_preview, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        patcher = patch.object(excel_ops, "win32", Mock(DispatchEx=Mock(return_value=self.excel)))
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_renders_thumbnails_marks_similarity_and_reuses_cache(self):
        result = sheet_preview.render_workbook_previews(str(self.folder), "spec.xlsx")
        sheets = {s["name"]: s for s in result["sheets"]}
        self.assertFalse(result["cached"])
        self.assertEqual(sheets["Main"]["pages"], 2)
        self.assertTrue(preview_image_path(result["key"], sheets["Main"]["image"]))
        # 导出的 PDF 一并保留，可在浏览器里查看完整 PDF
        self.assertTrue(preview_image_path(result["key"], sheets["Main"]["pdf"]))
        # 隐藏表临时取消隐藏后照样出图；只读打开、关闭时不保存
        self.assertTrue(sheets["Main(2)"]["hidden"])
        self.assertTrue(sheets["Main(2)"]["image"])
        self.book.Close.assert_called_once_with(SaveChanges=False)
        self.assertEqual(sheets["Main"]["similar"], {"match": "Main(2)", "ratio": 1.0, "identical": True})
        self.assertEqual(sheets["Blank"]["image"], "")
        self.assertIn("空白", sheets["Blank"]["note"])

        again = sheet_preview.render_workbook_previews(str(self.folder), "spec.xlsx")
        self.assertTrue(again["cached"])
        self.assertEqual(again["sheets"], result["sheets"])
        self.excel.Workbooks.Open.assert_called_once()

    def test_rejects_files_outside_scan_results(self):
        with self.assertRaises(ValueError):
            sheet_preview.render_workbook_previews(str(self.folder), "../secret.xlsx")


if __name__ == "__main__":
    unittest.main()
