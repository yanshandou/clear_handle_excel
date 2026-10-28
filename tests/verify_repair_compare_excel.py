"""用两份同名临时工作簿验证新旧对比，不接触用户 Excel：python tests/verify_repair_compare_excel.py。"""
import json
import sys
import tempfile
from pathlib import Path
from unittest.mock import patch

import openpyxl
import win32process

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import excel_ops as ops, link_repair, repair_compare as compare


def main():
    with tempfile.TemporaryDirectory(prefix="repair_compare_") as directory:
        root = Path(directory)
        current, backup = root / "current" / "same_name.xlsx", root / "backup" / "same_name.xlsx"
        for path, sheet_name, value in ((current, "NewMain", 20), (backup, "Main", 10)):
            path.parent.mkdir()
            book = openpyxl.Workbook()
            book.active.title = sheet_name
            book.active["B7"] = value
            book.active["B8"] = value + 1
            book.save(path)
            book.close()
        stamps = {path:link_repair.fingerprint(path) for path in (current, backup)}
        ops.pythoncom.CoInitialize()
        excel = None
        viewer = compare.ComparisonWindows()
        try:
            excel = ops.win32.DispatchEx("Excel.Application")
            excel.DisplayAlerts = False
            excel.AutomationSecurity = ops.MSO_AUTOMATION_SECURITY_FORCE_DISABLE
            target = {"current":current, "backup":backup, "sheet":"NewMain", "old_sheet":"Main", "cell":"B7"}
            with patch.object(ops, "_visible_excel", return_value=excel), patch.object(ops, "_bring_to_front"):
                first = viewer.open(target)
                assert first == {"opened":["current", "backup"], "note":""}, first
                apps = [viewer.sessions[(current, False)], viewer.sessions[(backup, True)]]
                pids = [win32process.GetWindowThreadProcessId(int(app.Hwnd))[1] for app in apps]
                assert pids[0] != pids[1], pids
                current_book, backup_book = [app.Workbooks(1) for app in apps]
                assert Path(current_book.FullName) == current and Path(backup_book.FullName) == backup
                assert not current_book.ReadOnly and backup_book.ReadOnly
                assert [app.ActiveSheet.Name for app in apps] == ["NewMain", "Main"]
                assert [app.Selection.Address for app in apps] == ["$B$7", "$B$7"]
                assert all(app.Workbooks.Count == 1 for app in apps)
                assert current_book.Saved and backup_book.Saved
                # 后续点击复用两边，并保留当前文件尚未保存的手动修改。
                current_book.Worksheets("NewMain").Range("C1").Value2 = "unsaved change"
                second = viewer.open({**target, "cell":"B8"})
                assert not second["note"], second
                assert [app.Selection.Address for app in apps] == ["$B$8", "$B$8"]
                assert [win32process.GetWindowThreadProcessId(int(app.Hwnd))[1] for app in apps] == pids
                assert all(app.Workbooks.Count == 1 for app in apps)
                assert current_book.Worksheets("NewMain").Range("C1").Value2 == "unsaved change" and not current_book.Saved
                assert all(link_repair.fingerprint(path) == stamp for path, stamp in stamps.items())
                import win32gui
                rectangles = [win32gui.GetWindowRect(int(app.Workbooks(1).Windows(1).Hwnd)) for app in apps]
                assert rectangles[0][0] < rectangles[1][0] and abs(rectangles[0][2] - rectangles[1][0]) < 20, rectangles
                print(json.dumps({"same_name_separate_processes":True, "backup_readonly":True,
                    "renamed_sheets_located":True, "side_by_side":True, "windows_reused":True,
                    "unsaved_current_edits_preserved":True, "files_unchanged":True}))
        finally:
            # 这些实例全部由本测试创建，只关闭临时样本，绝不枚举或退出用户的 Excel。
            unique = {int(app.Hwnd):app for app in viewer.sessions.values()}
            if excel is not None:
                unique[int(excel.Hwnd)] = excel
            for app in unique.values():
                app.DisplayAlerts = False
                for book in list(app.Workbooks):
                    book.Close(SaveChanges=False)
                app.Quit()
            ops.pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
