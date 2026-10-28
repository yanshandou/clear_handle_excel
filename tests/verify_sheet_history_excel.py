"""真实 Excel 验证删表重命名后的隐藏状态、列宽、颜色与内容核对。"""
import copy
import json
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import color_compare, hidden_repair, mark_finder, sheet_history, workflow as wf
from core.excel_ops import pythoncom, start_background_excel, quit_background_excel


def verify(excel, root, suffix, file_format):
    old_path, path = root / f"backup{suffix}", root / f"current{suffix}"
    book = excel.Workbooks.Add()
    try:
        while book.Sheets.Count > 1:
            book.Sheets(book.Sheets.Count).Delete()
        wrong = book.Sheets(1)
        wrong.Name = "Layout"
        wrong.Range("A1").Value2 = "old version"
        wrong.Range("A2").Value2 = 42
        current = book.Worksheets.Add(After=wrong)
        current.Name = "新Layout"
        keep = book.Worksheets.Add(After=current)
        keep.Name = "Keep"
        keep.Range("A1").Value2 = "stable"
        keep.Range("A2").Formula = "='新Layout'!A1"
        current.Range("A1").Value2 = "current version"
        current.Range("A2").Formula = "=Layout!A2"
        current.Range("A3").Formula = "=Keep!A1"
        current.Range("B:C").EntireColumn.Hidden = True
        current.Range("D:D").ColumnWidth = 0.25
        current.Range("A1:A3").Interior.Color = 0x332211
        book.SaveAs(str(old_path), FileFormat=file_format)
        book.SaveCopyAs(str(path))
    finally:
        book.Close(SaveChanges=False)
    before_hash = sheet_history.fingerprint(old_path)
    wf._clean_backup_sheets(lambda: excel, path, root, {"delete": ["Layout"], "uncertain": []})
    # 再改一次名，验证不是只保留最后一条对应关系。
    wf._clean_backup_sheets(lambda: excel, path, root,
                            {"delete": [], "uncertain": [], "rename": {"Layout": "Final"}})
    book = excel.Workbooks.Open(str(path), UpdateLinks=0)
    try:
        current = book.Sheets("Final")
        current.Range("B:C").EntireColumn.Hidden = False
        current.Range("D:D").ColumnWidth = 12
        current.Range("A1:A3").Interior.Color = 0x0000FF
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    plan = hidden_repair.plan_file(excel, path, old_path, hidden_repair.DEFAULT_OPTIONS)
    ready = [i for i in plan["cells"] if i["action"] == "restore"]
    assert {(i["kind"], i["addr"], i["old_sheet"]) for i in ready} == {
        ("columns", "B:C", "新Layout"), ("narrow", "D:D", "新Layout")}, plan
    hidden_repair.apply_file(excel, plan)
    stamps = (sheet_history.fingerprint(path), sheet_history.fingerprint(old_path))
    color_plan = color_compare.plan_file(excel, path, old_path)["cells"]
    assert len(color_plan) == 3 and all(i["action"] == "restore" and i["old_sheet"] == "新Layout" for i in color_plan), color_plan
    done = mark_finder.apply_file(excel, str(path), str(old_path), color_plan, versions=stamps)
    assert all(i["status"] == "ok" for i in done), done
    assert sheet_history.fingerprint(old_path) == before_hash, "备份被改写"
    book = excel.Workbooks.Open(str(path), ReadOnly=True, UpdateLinks=0)
    try:
        current = book.Sheets("Final")
        assert current.Range("B:C").EntireColumn.Hidden is True
        assert abs(current.Range("D:D").ColumnWidth - 0.25) < 0.05
        assert current.Range("A1:A3").Interior.Color == 0x332211
        assert current.Range("A2").Value2 == 42 and not current.Range("A2").HasFormula
        assert "Final" in book.Sheets("Keep").Range("A2").Formula
    finally:
        book.Close(SaveChanges=False)
    return {"format": suffix, "history": len(sheet_history.get(path)), "restored_columns": len(ready), "restored_colors": len(done)}


def verify_legacy_inference(excel, root):
    old_path, path = root / "legacy_backup.xlsx", root / "legacy_current.xlsx"
    book = excel.Workbooks.Add()
    try:
        while book.Sheets.Count > 1:
            book.Sheets(book.Sheets.Count).Delete()
        old = book.Sheets(1)
        old.Name = "Layout"
        old.Range("A1").Value2 = "outdated"
        new = book.Worksheets.Add(After=old)
        new.Name = "新Layout"
        new.Range("A1").Value2 = "current version"
        new.Range("A2").Formula = "='新Layout'!A1"
        new.Range("B:B").EntireColumn.Hidden = True
        keep = book.Worksheets.Add(After=new)
        keep.Name = "Keep"
        keep.Range("A1").Value2 = "keep"
        book.SaveAs(str(old_path), FileFormat=51)
        book.SaveCopyAs(str(path))
    finally:
        book.Close(SaveChanges=False)
    wf._clean_backup_sheets(lambda: excel, path, root, {"delete": ["Layout"], "uncertain": []})
    wf.STATE["sheet_history"].clear()  # 模拟旧版本清理后没有保存操作历史。
    book = excel.Workbooks.Open(str(path), UpdateLinks=0)
    try:
        book.Worksheets("Layout").Columns(2).Hidden = False
        book.Save()
    finally:
        book.Close(SaveChanges=False)
    plan = hidden_repair.plan_file(excel, path, old_path, hidden_repair.DEFAULT_OPTIONS)
    found = [item for item in plan["cells"] if item["kind"] == "columns"]
    assert len(found) == 1 and found[0]["old_sheet"] == "新Layout" and found[0]["match"] == "rule" and found[0]["action"] == "restore", plan
    hidden_repair.apply_file(excel, plan)
    book = excel.Workbooks.Open(str(path), UpdateLinks=0, ReadOnly=True)
    try:
        assert book.Worksheets("Layout").Columns(2).Hidden is True
    finally:
        book.Close(SaveChanges=False)
    return {"without_history": True, "inferred_old_sheet": found[0]["old_sheet"], "restored_columns": len(found)}


def main():
    pythoncom.CoInitialize()
    excel = None
    try:
        with tempfile.TemporaryDirectory(prefix="sheet_history_") as directory:
            state = copy.deepcopy(wf.STATE)
            state.update(sheet_history={}, folder=directory, pages={}, items={}, checked_files={}, review_queue={})
            with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"), \
                    patch.object(wf, "update_file_progress"), patch.object(wf, "log_msg"), \
                    patch.object(wf, "get_execution_report_snapshot", return_value={"files": {}}):
                excel = start_background_excel()
                results = [verify(excel, Path(directory), suffix, fmt) for suffix, fmt in ((".xlsx", 51), (".xls", 56))]
                results.append(verify_legacy_inference(excel, Path(directory)))
                print(json.dumps({"ok": True, "results": results}, ensure_ascii=False))
    finally:
        quit_background_excel(excel)
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
