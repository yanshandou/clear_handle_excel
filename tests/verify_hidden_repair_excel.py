"""真实 Excel 临时样本验证：python tests/verify_hidden_repair_excel.py。"""
import json
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import hidden_repair as repair, workflow as wf
from core.excel_ops import pythoncom, quit_background_excel, start_background_excel


def make_sample(excel, current, old, name, file_format):
    workbook = excel.Workbooks.Add()
    try:
        main = workbook.Worksheets(1)
        main.Name = "Main"
        while workbook.Worksheets.Count > 1:
            workbook.Worksheets(workbook.Worksheets.Count).Delete()
        main.Range("A1:F2").Value2 = (("title", "private B", "private C", "normal", "very narrow", "formula"),
                                     (1, 2, 3, 4, 5, 6))
        main.Range("F2").Formula = "=1+2"
        main.Range("G1:H1").Merge()
        main.Range("G1").Value2 = "merged"
        main.Range("A:A").ColumnWidth = 10
        main.Range("B:C").EntireColumn.Hidden = True
        main.Range("E:E").ColumnWidth = 0.25
        main.Range("Z:Z").EntireColumn.Hidden = True
        main.Range("F2").NumberFormat = "0.00"
        for name_, state in (("Aux", 0), ("Deep", 2), ("Changed", -1)):
            sample = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
            sample.Name = name_
            sample.Range("A1").Value2 = name_
            sample.Range("B:B").EntireColumn.Hidden = True
            sample.Visible = state
        workbook.SaveAs(str(old / name), FileFormat=file_format)
        workbook.SaveCopyAs(str(current / name))
    finally:
        workbook.Close(SaveChanges=False)
    workbook = excel.Workbooks.Open(str(current / name), UpdateLinks=0)
    try:
        for sample in workbook.Worksheets:
            sample.Visible = -1
            sample.Range("A:Z").Columns.AutoFit()
            sample.Columns(2).Hidden = False
        assert workbook.Worksheets("Main").Range("B:C").EntireColumn.Hidden is False, "样本中 AutoFit 应已显示隐藏列"
        main = workbook.Worksheets("Main")
        main.Range("A:A").ColumnWidth = 23
        main.Range("Q:Q").EntireColumn.Hidden = True  # 当前后来隐藏的列不能被取消隐藏。
        main.Columns(26).Hidden = False  # UsedRange 外的列也要对照原始备份。
        workbook.Worksheets("Changed").Range("A1").Value2 = "edited or moved"
        workbook.Save()
    finally:
        workbook.Close(SaveChanges=False)


def main():
    with tempfile.TemporaryDirectory(prefix="hidden_repair_") as directory:
        root = Path(directory)
        current, old = root / "current" / "001", root / "backup" / "001"
        current.mkdir(parents=True)
        old.mkdir(parents=True)
        formats = (("sample.xlsx", 51), ("sample.xlsm", 52), ("sample.xls", 56), ("sample.xlsb", 50))
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = start_background_excel()
            for name, file_format in formats:
                make_sample(excel, current, old, name, file_format)
        finally:
            quit_background_excel(excel)
            pythoncom.CoUninitialize()
        stamps = {path: repair.fingerprint(path) for path in old.iterdir() if not path.name.startswith("~$")}
        with patch.object(wf, "STATE", {"busy": False, "job": None}), patch.object(wf, "log_msg"):
            repair.repair(str(current), str(old.parent))
            deadline = time.monotonic() + 120
            while repair.status()["running"]:
                if time.monotonic() > deadline:
                    raise TimeoutError("真实 Excel 隐藏恢复超过 120 秒")
                time.sleep(0.2)
            result = repair.status()["applied"]
            assert not any(file["error"] for file in result["files"]), result
            assert not wf.STATE["busy"] and "maintenance" not in wf.STATE
        pythoncom.CoInitialize()
        excel = None
        try:
            excel = start_background_excel()
            for name, _ in formats:
                workbook = excel.Workbooks.Open(str(current / name), UpdateLinks=0, ReadOnly=True)
                try:
                    main_sheet = workbook.Worksheets("Main")
                    assert main_sheet.Range("B:C").EntireColumn.Hidden is True
                    assert main_sheet.Range("Z:Z").EntireColumn.Hidden is True
                    assert main_sheet.Range("Q:Q").EntireColumn.Hidden is True
                    assert main_sheet.Range("E:E").EntireColumn.Hidden is False
                    assert abs(main_sheet.Range("E:E").ColumnWidth - 0.25) < 0.05
                    assert abs(main_sheet.Range("A:A").ColumnWidth - 23) < 0.05
                    assert main_sheet.Range("F2").Formula == "=1+2"
                    assert main_sheet.Range("F2").NumberFormat == "0.00"
                    assert main_sheet.Range("G1:H1").MergeCells is True
                    assert workbook.Worksheets("Aux").Visible == 0
                    assert workbook.Worksheets("Deep").Visible == 2
                    assert workbook.Worksheets("Changed").Range("B:B").EntireColumn.Hidden is False
                finally:
                    workbook.Close(SaveChanges=False)
        finally:
            quit_background_excel(excel)
            pythoncom.CoUninitialize()
        assert all(repair.fingerprint(path) == stamp for path, stamp in stamps.items())
        assert all(any(item["action"] == "skip" and item["sheet"] == "Changed" for item in file["cells"])
                   for file in result["files"])
        print(json.dumps({"formats": [name for name, _ in formats], "hidden_sheets_and_columns": True,
                          "very_hidden_preserved": True, "narrow_width_restored_without_hiding": True,
                          "normal_widths_formulas_and_merges_preserved": True, "changed_sheet_skipped": True,
                          "backup_unchanged": True, "background_lock_released": True}))


if __name__ == "__main__":
    main()
