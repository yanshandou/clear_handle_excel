"""真实 Excel 验证，仅在临时目录创建样表：python tests/verify_formula_links_excel.py。"""
from pathlib import Path
import json
import sys
import tempfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import formula_links as links, link_repair
from core.excel_ops import pythoncom, start_background_excel, quit_background_excel


def make_book(excel, path, arrays=True):
    book = excel.Workbooks.Add()
    main = book.Worksheets(1)
    main.Name = "Main"
    while book.Worksheets.Count > 1:
        book.Worksheets(book.Worksheets.Count).Delete()
    old = book.Worksheets.Add(After=main)
    old.Name = "旧表"
    old.Range("A1:A3").Value2 = ((10,), (20,), (30,))
    old.Range("D1:E3").Value2 = (("Key", "Amount"), ("x", 4), ("y", 5))
    table = old.ListObjects.Add(1, old.Range("D1:E3"), None, 1)
    table.Name = "TableData"
    middle = book.Worksheets.Add(After=old)
    middle.Name = "Middle"
    middle.Range("A1").Value2 = 40
    last = book.Worksheets.Add(After=middle)
    last.Name = "Last"
    last.Range("A1").Value2 = 50
    book.Names.Add(Name="OriginalValue", RefersTo="='旧表'!$A$1")
    book.Names.Add(Name="AliasValue", RefersTo="=OriginalValue")
    formulas = {
        "A1": "='旧表'!A1", "A2": "='旧表'!A1*2", "A3": "=IFERROR('旧表'!A2,0)",
        "A4": "=AliasValue", "A5": '=INDIRECT(B5&"!A1")', "A6": "=A2+1",
        "A7": '=IF(\'旧表\'!A1=10,"=literal","")', "A8": "=SUM('旧表:Last'!A1)",
        "A9": "=SUM(TableData[Amount])", "A21": "=#REF!",
    }
    main.Range("B5").Value2 = "旧表"
    for addr, value in formulas.items():
        main.Range(addr).Formula = value
    main.Range("A2").NumberFormat = "0.00"
    main.Range("A2").Interior.Color = 123456
    main.Range("F2:G2").Merge()
    main.Range("F2").Formula = "='旧表'!A2"
    if arrays:
        main.Range("C1:C2").FormulaArray = "='旧表'!A1:A2*2"
        try:
            main.Range("D5").Formula2 = "=SEQUENCE(2,1,'旧表'!A1)"
        except Exception:
            pass  # 本机旧版 Excel 不支持动态数组，其他断链场景仍需验证。
    excel.CalculateFull()
    book.SaveAs(str(path), FileFormat=51)
    return book


def main():
    pythoncom.CoInitialize()
    excel = book = None
    with tempfile.TemporaryDirectory(prefix="formula_links_excel_") as directory:
        root = Path(directory)
        try:
            excel = start_background_excel()
            book = make_book(excel, root / "guard.xlsx")
            main = book.Worksheets("Main")
            before = {addr: main.Range(addr).Value2 for addr in ("A1", "A2", "A3", "A4", "A5", "A7", "A8", "A9", "C1:C2", "F2")}
            if main.Range("D5").HasFormula:
                before["D5:D6"] = main.Range("D5:D6").Value2
            kept = links.delete_sheets_preserving_values(book, ["旧表"])
            for addr, value in before.items():
                assert main.Range(addr).Value2 == value, (addr, value, main.Range(addr).Value2)
                assert not main.Range(addr).HasFormula, addr
            assert main.Range("A6").HasFormula and main.Range("A6").Value2 == 21
            assert main.Range("A2").NumberFormat == "0.00" and main.Range("A2").Interior.Color == 123456
            assert main.Range("F2").MergeCells
            assert links.error_positions(book) == {("Main", "A21")}
            book.Save()
            book.Close(SaveChanges=False)
            book = None

            book = make_book(excel, root / "blocked.xlsx", False)
            book.Worksheets("Main").Protect(Password="test")
            try:
                links.delete_sheets_preserving_values(book, ["旧表"])
                raise AssertionError("受保护的工作表应中止删表")
            except Exception as error:
                assert "旧表" in [sheet.Name for sheet in book.Sheets], str(error)
            book.Close(SaveChanges=False)
            book = None

            original = root / "backup" / "001"
            current = root / "current" / "001"
            original.mkdir(parents=True)
            current.mkdir(parents=True)
            old_path, path = original / "book.xlsx", current / "book.xlsx"
            book = make_book(excel, old_path, True)
            book.SaveCopyAs(str(path))
            book.Close(SaveChanges=False)
            book = excel.Workbooks.Open(str(path), UpdateLinks=0)
            book.Sheets("旧表").Delete()  # 重现旧版本的损坏。
            book.Worksheets("Main").Name = "新Main"
            book.Save()
            book.Close(SaveChanges=False)
            book = None
            old_stamp, current_stamp = link_repair.fingerprint(old_path), link_repair.fingerprint(path)
            scanned = link_repair.scan_file(excel, path)
            scanned.update(name="book.xlsx", error=None)
            assert len(scanned["hits"]) >= 8, scanned
            plan = link_repair.preview_file(excel, scanned, old_path)
            assert link_repair.fingerprint(old_path) == old_stamp
            assert link_repair.fingerprint(path) == current_stamp
            accepted = [cell for cell in plan["cells"] if cell["action"] == "restore"]
            assert any(cell["addr"] == "A2" and cell["value"] == 20 for cell in accepted), plan
            assert any(cell["addr"] == "A3" and cell["value"] == 20 for cell in accepted), plan
            assert any(cell["addr"] == "A21" and cell["action"] == "skip" for cell in plan["cells"]), plan
            plan["cells"] = accepted
            result = link_repair.apply_file(excel, plan)
            book = excel.Workbooks.Open(str(path), ReadOnly=True, UpdateLinks=0)
            main = book.Worksheets("新Main")
            for addr, value in {"A1": 10, "A2": 20, "A3": 20, "A4": 10, "A5": 10, "A7": "=literal", "A9": 9}.items():
                assert main.Range(addr).Value2 == value, (addr, main.Range(addr).Value2)
            assert main.Range("A6").HasFormula and main.Range("A6").Value2 == 21
            assert main.Range("C1:C2").Value2 == ((20.0,), (40.0,))
            assert not main.Range("C1:C2").HasFormula
            assert links.error_positions(book) == {("新Main", "A21")}
            assert main.Range("A2").NumberFormat == "0.00" and main.Range("A2").Interior.Color == 123456
            book.Close(SaveChanges=False)
            book = None
            assert link_repair.fingerprint(old_path) == old_stamp
            try:
                link_repair.apply_file(excel, plan)
                raise AssertionError("过期预览应拒绝再次覆盖")
            except ValueError as error:
                assert "变化" in str(error)
            print(json.dumps({"guarded_regions": len(kept), "repaired_regions": len(result),
                              "backup_unchanged": True, "formats_preserved": True,
                              "protected_sheet_blocked": True, "stale_preview_blocked": True}))
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            quit_background_excel(excel)
            pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
