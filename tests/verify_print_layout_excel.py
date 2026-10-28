"""真实 Excel 冒烟验证，仅创建临时样表：python tests/verify_print_layout_excel.py。"""
import hashlib
import json
from pathlib import Path
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import print_layout as layout
from core.excel_ops import start_background_excel, quit_background_excel, pythoncom
from core.pdf_detector import fitz


def export(sheet, pdf):
    sheet.ExportAsFixedFormat(Type=0, Filename=str(pdf), Quality=0, IncludeDocProperties=True,
                              IgnorePrintAreas=False, OpenAfterPublish=False)


def main():
    records = {}
    layout.configure(lambda: records)
    pythoncom.CoInitialize()
    excel = book = None
    temporary = tempfile.TemporaryDirectory(prefix="orientation_excel_")
    try:
        excel = start_background_excel()
        directory = temporary.name
        root = Path(directory)
        path = root / "samples.xlsx"
        book = excel.Workbooks.Add()
        while book.Worksheets.Count < 3:
            book.Worksheets.Add(After=book.Worksheets(book.Worksheets.Count))
        names = ["Wide", "Tall", "FitWidth"]
        for index, name in enumerate(names, 1):
            sheet = book.Worksheets(index)
            sheet.Name = name
            columns, rows = (2, 90) if name == "Tall" else (12, 75)
            data = tuple(tuple(f"r{row:02}c{col:02}" for col in range(1, columns + 1))
                         for row in range(1, rows + 1))
            area = sheet.Range(sheet.Cells(1, 1), sheet.Cells(rows, columns))
            area.Value = data
            area.Font.Name = "Arial"
            area.Font.Size = 11
            area.ColumnWidth = 9
            area.RowHeight = 18
            sheet.PageSetup.Orientation = 1
            sheet.PageSetup.PaperSize = 9
            sheet.PageSetup.LeftMargin = 36
            sheet.PageSetup.RightMargin = 36
            sheet.PageSetup.TopMargin = 36
            sheet.PageSetup.BottomMargin = 36
            sheet.PageSetup.PrintArea = area.Address
            sheet.PageSetup.PrintTitleRows = "$1:$1"
            sheet.PageSetup.RightFooter = "Page &P of &N"
            if name == "FitWidth":
                sheet.PageSetup.Zoom = False
                sheet.PageSetup.FitToPagesWide = 1
                sheet.PageSetup.FitToPagesTall = False
            else:
                sheet.PageSetup.Zoom = 100
        book.SaveAs(str(path), FileFormat=51)
        before = hashlib.sha256(path.read_bytes()).hexdigest()
        result = []
        for name in names:
            sheet = book.Worksheets(name)
            output = root / f"{name}.pdf"
            record, made = layout.prepare(sheet, str(path), records, export, output)
            if record is None:
                diagnostics = {}
                for prop in ("Orientation",) + layout.SETTINGS:
                    try:
                        value = getattr(sheet.PageSetup, prop)
                        diagnostics[prop] = [type(value).__name__, repr(value)]
                    except Exception as error:
                        diagnostics[prop] = str(error)
                print(json.dumps({"identity": layout.identity(sheet), "settings": diagnostics}, ensure_ascii=True), flush=True)
            assert made and not record.get("fallback"), record
            assert sheet.PageSetup.Orientation == 1
            expected = 1 if name == "Tall" else 2
            assert record["selected"] == expected, record
            layout.bind(book, str(path))
            with layout.applied(sheet):
                assert sheet.PageSetup.Orientation == expected
                export(sheet, root / f"{name}_again.pdf")
            assert sheet.PageSetup.Orientation == 1
            with fitz.open(output) as doc:
                assert all((p.rect.width > p.rect.height) == (expected == 2) for p in doc)
            result.append({"sheet": name, "selected": layout.label(record["selected"]),
                           "reason": record["reason"], "metrics": record["metrics"]})
        assert hashlib.sha256(path.read_bytes()).hexdigest() == before
        book.Close(SaveChanges=False)
        book = None
        reopened = excel.Workbooks.Open(str(path), ReadOnly=True, UpdateLinks=0)
        try:
            assert all(reopened.Worksheets(name).PageSetup.Orientation == 1 for name in names)
        finally:
            reopened.Close(SaveChanges=False)
        print(json.dumps({"source_unchanged": True, "results": result}, ensure_ascii=True))
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            quit_background_excel(excel)
        pythoncom.CoUninitialize()
        temporary.cleanup()


if __name__ == "__main__":
    main()
