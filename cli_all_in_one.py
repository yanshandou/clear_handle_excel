import shutil
import sys
from collections.abc import Iterable, Iterator
from datetime import datetime
from pathlib import Path
from typing import TypeVar

from core.common import (
    EXCEL_EXTENSIONS,
    EXCLUDE_PDF_SHEETS,
    RECORD_PATH,
    XL_SHEET_VISIBLE,
    XL_TYPE_PDF,
    append_record,
    is_backup_sheet,
    sanitize_filename,
    unique_pdf_path,
)
from core.excel_ops import (
    adjust_row_heights_against_clipping,
    normalize_workbook_view,
    pythoncom,
    scan_workbook_sheets,
    win32,
)
from core.reporter import (
    finish_execution_batch,
    init_execution_batch,
    update_file_progress,
)

__all__ = [
    "EXCEL_EXTENSIONS",
    "EXCLUDE_PDF_SHEETS",
    "RECORD_PATH",
    "adjust_row_heights_against_clipping",
    "append_record",
    "is_backup_sheet",
    "main",
    "normalize_sheets_view",
    "process_workbooks",
    "progress_iter",
    "progress_write",
    "sanitize_filename",
    "scan_workbook_sheets",
]

_T = TypeVar("_T")

try:
    from tqdm import tqdm as _tqdm
except ImportError:
    _tqdm = None


def progress_iter(iterable: Iterable[_T], desc: str = "", unit: str = "") -> Iterator[_T]:
    if _tqdm is not None:
        yield from _tqdm(iterable, desc=desc, unit=unit, dynamic_ncols=True)
        return

    items = list(iterable)
    total = len(items)
    if total == 0:
        return
    bar_width = 24
    for completed, item in enumerate(items, start=1):
        filled = int(bar_width * (completed - 1) / total)
        bar = "#" * filled + "-" * (bar_width - filled)
        sys.stderr.write(f"\r{desc}: [{bar}] {completed - 1}/{total} {unit}")
        sys.stderr.flush()
        yield item
    filled = bar_width
    bar = "#" * filled
    sys.stderr.write(f"\r{desc}: [{bar}] {total}/{total} {unit}\n")
    sys.stderr.flush()


def progress_write(message: str) -> None:
    if _tqdm is not None:
        _tqdm.write(message)
        return
    sys.stderr.write(f"\n{message}\n")
    sys.stderr.flush()


def normalize_sheets_view(excel, wb) -> None:
    """将所有可见工作表定位到左上角 A1、缩放设为 100%，并停留在第一个可见 Sheet"""
    normalize_workbook_view(excel, wb, on_error=progress_write)


def process_workbooks(
    workbooks: list[Path],
    do_delete_backups: bool,
    do_fix_clipping: bool,
    do_normalize_view: bool,
    do_export_pdf: bool,
    backup_folder: Path,
    pdf_folder: Path,
) -> None:
    pythoncom.CoInitialize()
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False

    backup_created = False
    used_pdf_paths: set[Path] = set()
    if do_export_pdf:
        pdf_folder.mkdir(parents=True, exist_ok=True)

    init_execution_batch(
        task_name="CLI 命令行批处理任务",
        folder=backup_folder.parent,
        workbooks=workbooks,
        pdf_folder=pdf_folder if do_export_pdf else None,
        backup_folder=backup_folder,
    )

    try:
        for workbook_path in progress_iter(workbooks, desc="处理工作簿", unit="文件"):
            wb = None
            file_backed_up = False
            need_save = False
            exported_count = 0
            deleted_list: list[str] = []

            update_file_progress(
                workbook_path.name,
                status="processing",
                current_action="正在打开并处理工作簿...",
            )

            try:
                wb = excel.Workbooks.Open(str(workbook_path.resolve()))
                original_sheets = [sh.Name for sh in wb.Sheets]
                backup_sheets = [name for name in original_sheets if is_backup_sheet(name)]

                def ensure_backup():
                    nonlocal backup_created, file_backed_up
                    if not backup_created:
                        backup_folder.mkdir(parents=True, exist_ok=True)
                        backup_created = True
                    if not file_backed_up:
                        shutil.copy2(workbook_path, backup_folder / workbook_path.name)
                        file_backed_up = True

                # 1. 删除备份 Sheet
                if do_delete_backups and backup_sheets:
                    if len(backup_sheets) >= len(original_sheets):
                        progress_write(f"跳过删除（删除后将无剩余工作表）: {workbook_path.name}")
                    else:
                        ensure_backup()
                        for sheet_name in backup_sheets:
                            wb.Sheets(sheet_name).Delete()
                        deleted_list = list(backup_sheets)
                        need_save = True
                        progress_write(f"已清理: {workbook_path.name} (删除 {len(backup_sheets)} 个备份表)")

                # 2. 智能增加多行文本的行高（防 PDF 切字）
                if do_fix_clipping:
                    ensure_backup()
                    adjust_row_heights_against_clipping(wb, extra_points=6.0)
                    need_save = True

                # 3. A1 定位 + 100% 缩放
                if do_normalize_view:
                    ensure_backup()
                    normalize_sheets_view(excel, wb)
                    need_save = True

                if need_save:
                    wb.Save()

                # 4. 导出 PDF
                if do_export_pdf:
                    file_stem = sanitize_filename(workbook_path.stem)
                    for sheet in wb.Sheets:
                        sheet_name = sheet.Name
                        if sheet_name in EXCLUDE_PDF_SHEETS or is_backup_sheet(sheet_name):
                            continue
                        if sheet.Visible != XL_SHEET_VISIBLE:
                            continue

                        safe_sheet_name = sanitize_filename(sheet_name)
                        pdf_path = unique_pdf_path(pdf_folder, file_stem, safe_sheet_name, used_pdf_paths)

                        try:
                            sheet.ExportAsFixedFormat(
                                Type=XL_TYPE_PDF,
                                Filename=str(pdf_path.resolve()),
                                Quality=0,
                                IncludeDocProperties=True,
                                IgnorePrintAreas=False,
                                OpenAfterPublish=False,
                            )
                            exported_count += 1
                        except Exception as pdf_err:
                            progress_write(f"  导出 PDF 失败 [{workbook_path.name} -> {sheet_name}]: {pdf_err}")

                    progress_write(f"已导出 PDF: {workbook_path.name} (共 {exported_count} 个 Sheet)")

                update_file_progress(
                    workbook_path.name,
                    status="done",
                    all_sheets=original_sheets,
                    deleted_backups=deleted_list,
                    exported_pdfs=exported_count,
                    normalized_a1=do_normalize_view,
                    current_action=f"已完成 (清理 {len(deleted_list)} 表, 导出 {exported_count} 份 PDF)",
                )

            except Exception as error:
                progress_write(f"处理失败 {workbook_path.name}: {error}")
                update_file_progress(
                    workbook_path.name,
                    status="error",
                    error=str(error),
                    current_action=f"处理失败: {error}",
                )
            finally:
                if wb is not None:
                    wb.Close(SaveChanges=False)
    finally:
        excel.Quit()
        pythoncom.CoUninitialize()
        finish_execution_batch("CLI 批处理任务执行完毕")

    if backup_created:
        print(f"\n原始文件备份位置: {backup_folder}")
    if do_export_pdf:
        print(f"PDF 导出文件夹位置: {pdf_folder}")
    print(f"Markdown 执行记录报告已保存至: {RECORD_PATH}")


def main() -> None:
    folder_text = input("Excel 文件夹路径: ").strip().strip('"').strip("'")
    if not folder_text:
        return
    folder = Path(folder_text).expanduser().resolve()
    if not folder.is_dir():
        print(f"文件夹不存在: {folder}")
        return

    workbooks = sorted(
        path
        for path in folder.iterdir()
        if path.is_file()
        and path.suffix.lower() in EXCEL_EXTENSIONS
        and not path.name.startswith("~$")
    )
    if not workbooks:
        print("该文件夹中没有找到 Excel 文件。")
        return

    pending_delete: list[tuple[Path, list[str]]] = []
    for workbook_path in progress_iter(workbooks, desc="扫描文件", unit="文件"):
        try:
            all_sheets, backup_sheets = scan_workbook_sheets(workbook_path)
            if backup_sheets:
                if len(backup_sheets) >= len(all_sheets):
                    progress_write(f"警告: {workbook_path.name} 所有工作表均被识别为备份表，将跳过删除。")
                else:
                    pending_delete.append((workbook_path, backup_sheets))
        except Exception as error:
            progress_write(f"无法读取 {workbook_path.name}: {error}")

    do_delete = False
    if pending_delete:
        print("\n检测到以下备份工作表（名称含日期、bak/backup、old/旧、バックアップ/备份、コピー/copy/副本，或 Excel 副本「 (2)」）:")
        for workbook_path, sheet_names in pending_delete:
            print(f"\n{workbook_path.name}")
            for sheet_name in sheet_names:
                print(f"  - {sheet_name}")
        if input("\n[1/4] 是否删除这些备份工作表？(输入 y 确认): ").strip().lower() == "y":
            do_delete = True
    else:
        print("\n[1/4] 没有找到需要删除的备份工作表。")

    do_fix_clipping = (
        input("\n[2/4] 是否智能加高多行文本行的高度（防导出 PDF 切字）？(输入 y 确认): ")
        .strip()
        .lower()
        == "y"
    )

    do_normalize = (
        input("\n[3/4] 是否将所有工作表定位到 A1 单元格并以 100% 缩放保存？(输入 y 确认): ")
        .strip()
        .lower()
        == "y"
    )

    do_pdf = (
        input("\n[4/4] 是否将工作表导出为 PDF（排除「変更履歴」及备份表）？(输入 y 确认): ")
        .strip()
        .lower()
        == "y"
    )

    if not (do_delete or do_fix_clipping or do_normalize or do_pdf):
        print("已取消，未执行任何操作。")
        return

    timestamp = f"{datetime.now():%Y%m%d_%H%M%S}"
    backup_folder = folder / f"_sheet_cleanup_backup_{timestamp}"
    pdf_folder = folder / f"_exported_pdfs_{timestamp}"

    target_workbooks = (
        workbooks
        if (do_fix_clipping or do_normalize or do_pdf)
        else [item[0] for item in pending_delete]
    )

    process_workbooks(
        workbooks=target_workbooks,
        do_delete_backups=do_delete,
        do_fix_clipping=do_fix_clipping,
        do_normalize_view=do_normalize,
        do_export_pdf=do_pdf,
        backup_folder=backup_folder,
        pdf_folder=pdf_folder,
    )


if __name__ == "__main__":
    main()
