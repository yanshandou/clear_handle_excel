import ctypes
import hashlib
import json
import os
import shutil
import sys
import uuid
from pathlib import Path


SOURCE_ROOT = r"EDIT_SOURCE_ROOT"

SCRIPT_DIR = Path(__file__).resolve().parent
LOCAL_ROOT = SCRIPT_DIR.parent
MANIFEST_PATH = SCRIPT_DIR / "allocation.json"


def configure_stdio():
    for stream in (sys.stdout, sys.stderr):
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                try:
                    stream.reconfigure(errors="replace")
                except Exception:
                    pass


def get_desktop_path():
    if os.name != "nt":
        raise ValueError("仅 Windows 支持使用桌面文件夹名称，请填写绝对路径。")
    desktop_id = ctypes.create_string_buffer(
        uuid.UUID("B4BFCC3A-DB2C-424C-B029-7FE99A87C641").bytes_le, 16
    )
    shell = ctypes.WinDLL("shell32", use_last_error=True)
    memory = ctypes.WinDLL("ole32")
    shell.SHGetKnownFolderPath.argtypes = [
        ctypes.c_void_p, ctypes.c_uint32, ctypes.c_void_p, ctypes.POINTER(ctypes.c_void_p)
    ]
    shell.SHGetKnownFolderPath.restype = ctypes.c_long
    memory.CoTaskMemFree.argtypes = [ctypes.c_void_p]
    memory.CoTaskMemFree.restype = None
    desktop = ctypes.c_void_p()
    result = shell.SHGetKnownFolderPath(ctypes.byref(desktop_id), 0, None, ctypes.byref(desktop))
    if result != 0:
        raise OSError(f"无法获取 Windows 桌面路径，错误代码：0x{result & 0xFFFFFFFF:08X}")
    try:
        return Path(ctypes.wstring_at(desktop.value))
    finally:
        memory.CoTaskMemFree(desktop)


def resolve_configured_path(value):
    raw_value = os.fspath(value)
    path = Path(raw_value).expanduser()
    if path.is_absolute():
        return path
    if path.name != raw_value or raw_value in ("", ".", ".."):
        raise ValueError(f"请填写完整绝对路径，或仅填写桌面上的文件夹名称：{raw_value}")
    return get_desktop_path() / path


def load_allocations():
    with MANIFEST_PATH.open(encoding="utf-8-sig") as manifest_file:
        allocations = json.load(manifest_file)
    if not isinstance(allocations, list) or not allocations:
        raise ValueError("分配清单为空或格式不正确。")
    for allocation in allocations:
        if not isinstance(allocation, dict):
            raise ValueError("分配清单格式不正确。")
        for field in ("source", "destination"):
            parts = allocation.get(field)
            if not isinstance(parts, list) or not parts:
                raise ValueError("分配清单中的路径格式不正确。")
            for part in parts:
                if (not isinstance(part, str) or not part.strip()
                        or part in (".", "..") or part.endswith((" ", "."))
                        or any(char in '<>:"/\\|?*' or ord(char) < 32 for char in part)):
                    raise ValueError("分配清单中包含无效路径。")
    return allocations


def path_key(path):
    return os.path.normcase(os.path.abspath(path))


def paths_overlap(first, second):
    first_key = path_key(first)
    second_key = path_key(second)
    try:
        common = os.path.commonpath((first_key, second_key))
    except ValueError:
        return False
    return common == first_key or common == second_key


def is_linked_folder(path):
    return path.is_symlink() or bool(getattr(path.lstat(), "st_file_attributes", 0) & 0x400)


def file_hash(path):
    digest = hashlib.sha256()
    with path.open("rb") as file:
        for chunk in iter(lambda: file.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.digest()


def check_existing_file(source, destination, allow_overwrite=False):
    if destination.is_dir():
        raise ValueError(f"目标位置已有同名文件夹：{destination}")
    if not allow_overwrite and destination.exists() and (source.stat().st_size != destination.stat().st_size
                                                         or file_hash(source) != file_hash(destination)):
        raise ValueError(f"目标位置已有不同的文件，不会覆盖：{destination}")


def copy_allocations(source_root, destination_root, allocations, source_field, overwrite_existing=None):
    source_root = resolve_configured_path(source_root)
    destination_root = resolve_configured_path(destination_root)
    if not source_root.is_absolute() or not source_root.is_dir():
        raise ValueError(f"源目录必须是已存在的绝对路径：{source_root}")
    if not destination_root.is_absolute() or not destination_root.is_dir():
        raise ValueError(f"目标目录必须是已存在的绝对路径：{destination_root}")
    source_root = source_root.resolve()
    destination_root = destination_root.resolve()
    if paths_overlap(source_root, destination_root):
        raise ValueError("源目录和目标目录不能相同或互相包含。")

    folders = {}
    files = {}
    source_folders = []
    for allocation in allocations:
        source = source_root.joinpath(*allocation[source_field])
        destination = destination_root.joinpath(*allocation["destination"])
        if not source.is_dir():
            raise ValueError(f"找不到源文件夹：{source}")
        if is_linked_folder(source):
            raise ValueError(f"不支持链接文件夹：{source}")
        if any(paths_overlap(source, existing) for existing in source_folders):
            raise ValueError(f"分配清单中的源文件夹互相重叠：{source}")
        source_folders.append(source)

        def raise_walk_error(error):
            raise error

        for current, directory_names, file_names in os.walk(source, onerror=raise_walk_error):
            current_path = Path(current)
            target_folder = destination / current_path.relative_to(source)
            folders[path_key(target_folder)] = target_folder
            for directory_name in directory_names:
                directory = current_path / directory_name
                if is_linked_folder(directory):
                    raise ValueError(f"不支持链接文件夹：{directory}")
            for file_name in file_names:
                source_file = current_path / file_name
                target_file = target_folder / file_name
                key = path_key(target_file)
                if key in files:
                    raise ValueError(f"多个源文件对应同一目标：{target_file}")
                files[key] = (source_file, target_file)

    for folder in folders.values():
        current = folder
        while current != destination_root:
            if current.is_file() or path_key(current) in files:
                raise ValueError(f"目标文件夹的位置已有文件：{current}")
            current = current.parent
    for source_file, target_file in files.values():
        check_existing_file(source_file, target_file, allow_overwrite=overwrite_existing is not None)

    for folder in folders.values():
        folder.mkdir(parents=True, exist_ok=True)
    copied = 0
    for source_file, target_file in files.values():
        check_existing_file(source_file, target_file, allow_overwrite=overwrite_existing is not None)
        if target_file.exists() and overwrite_existing is not None:
            overwrite_existing(source_file, target_file)
            copied += 1
        elif not target_file.exists():
            with source_file.open("rb") as source_stream, target_file.open("xb") as target_stream:
                shutil.copyfileobj(source_stream, target_stream)
            copied += 1
    return copied, len(files)


def main():
    configure_stdio()
    try:
        copied, total = copy_allocations(SOURCE_ROOT, LOCAL_ROOT, load_allocations(), "source")
        print(f"第一步完成：共 {total} 个文件，新复制 {copied} 个文件。")
    except (OSError, ValueError, KeyError) as error:
        print(f"第一步失败：{error}")
        print("复制可能未完成；请解决问题后重新运行。")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())