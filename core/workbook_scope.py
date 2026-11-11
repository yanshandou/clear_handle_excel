"""Keep a selection of original Excel files and folders in its own workbench."""

import hashlib
import json
import os
import tempfile
import threading
from dataclasses import dataclass
from functools import cached_property, lru_cache
from pathlib import Path
from typing import Callable

from core import data_dir
from core.common import review_name_variants, strip_review_prefix

_BATCH_DIR = "批量路径"
_MANIFEST = "路径列表.json"
_LOCK = threading.Lock()


def parse_path(value: str) -> Path:
    value = value.strip()
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        value = value[1:-1]
    if not value.strip():
        raise ValueError("路径不能为空")
    return Path(value).expanduser().resolve()


def _file_key(path: Path) -> Path:
    return path.with_name(strip_review_prefix(path.name))


def file_variants(path: Path) -> list[Path]:
    return [path.with_name(name) for name in dict.fromkeys(
        [path.name, strip_review_prefix(path.name), *review_name_variants(path.name)]
    )]


@dataclass(frozen=True)
class WorkbookScope:
    folders: tuple[Path, ...]
    files: tuple[Path, ...]

    @cached_property
    def volume_roots(self) -> dict[str, Path]:
        groups: dict[str, list[Path]] = {}
        for folder in (*self.folders, *(p.parent for p in self.files)):
            groups.setdefault(os.path.normcase(folder.anchor), []).append(folder)
        return {anchor: Path(os.path.commonpath(groups[anchor])) for anchor in sorted(groups)}

    @property
    def common_root(self) -> Path | None:
        return next(iter(self.volume_roots.values())) if len(self.volume_roots) == 1 else None

    def relative(self, path: Path) -> Path:
        anchor = os.path.normcase(path.anchor)
        if anchor not in self.volume_roots:
            raise ValueError(f"路径不在批量任务中：{path}")
        relative = path.relative_to(self.volume_roots[anchor])
        if self.common_root is not None:
            return relative
        return Path(str(list(self.volume_roots).index(anchor) + 1)) / relative

    def source_path(self, relative: Path) -> Path | None:
        if relative.anchor or '..' in relative.parts:
            return None
        base = self.common_root
        if base is not None:
            return base / relative
        if not relative.parts:
            return None
        first, *parts = relative.parts
        return next((root.joinpath(*parts) for i, root in enumerate(self.volume_roots.values(), 1) if str(i) == first), None)

    def contains(self, path: Path) -> bool:
        return bool(path.name and _file_key(path) in self.files) or any(
            path.is_relative_to(folder) for folder in self.folders
        )

    def find(self, relative: Path) -> Path | None:
        path = relative if relative.is_absolute() else self.source_path(relative)
        if path is None or not path.name:
            return None
        for candidate in file_variants(path):
            if candidate.is_file():
                resolved = candidate.resolve()
                if self.contains(resolved):
                    return resolved
        return None


@lru_cache(maxsize=128)
def _load(root: Path) -> WorkbookScope | None:
    try:
        data = json.loads((root / _MANIFEST).read_text(encoding="utf-8"))
    except FileNotFoundError:
        return None
    return WorkbookScope(tuple(map(Path, data["folders"])), tuple(map(Path, data["files"])))


def get(root: Path | str) -> WorkbookScope | None:
    path = Path(root)
    if path.parent.name != _BATCH_DIR or path.parent != data_dir.DATA_DIR.resolve() / _BATCH_DIR:
        return None
    return _load(path)


def input_folders(value: str) -> list[str]:
    """Show the selected directories in the reference form without scanning their contents."""
    folders = set()
    for line in value.splitlines():
        if not line.strip():
            continue
        path = parse_path(line)
        scope = get(path)
        if scope:
            folders.update(scope.folders)
        elif path.is_dir():
            folders.add(path)
    return sorted(map(str, folders))


def resolve_input(value: str, validate: Callable[[str], str]) -> str:
    """A single folder keeps its existing identity. File selections have a stable directory for their artifacts."""
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    if not lines:
        raise ValueError("请输入 Excel 文件或文件夹路径")
    if len(lines) == 1:
        single = lines[0]
        validate(single)
        path = parse_path(single)
        if path.is_dir() or not path.name or not any(p.is_file() for p in file_variants(path)):
            return str(path)

    folders, files = set(), set()
    for line in lines:
        validate(line)
        path = parse_path(line)
        validate(str(path))
        if path.is_dir():
            if get(path):
                raise ValueError("请填写批量任务中的原始文件或文件夹路径")
            folders.add(path)
        elif path.name and any(candidate.is_file() for candidate in file_variants(path)):
            files.add(_file_key(path))
        else:
            raise ValueError(f"路径不存在或无法访问：{path}")

    # 保留显式选择的子文件夹，以便它可以单独指定参考目录。
    files = {file for file in files if not any(parent in folders for parent in file.parents)}
    if len(folders) == 1 and not files:
        return str(next(iter(folders)))

    data = {"folders": sorted(map(str, folders)), "files": sorted(map(str, files))}
    identity = {key: sorted(map(os.path.normcase, paths)) for key, paths in data.items()}
    # 路径留短一些，给 Excel 导出 PDF 的文件名留空间。
    digest = hashlib.sha256(json.dumps(identity).encode()).hexdigest()[:24]
    root = data_dir.DATA_DIR.resolve() / _BATCH_DIR / digest
    with _LOCK:
        manifest = root / _MANIFEST
        if not manifest.exists():
            root.mkdir(parents=True, exist_ok=True)
            temp = tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=root, delete=False)
            temp_path = Path(temp.name)
            try:
                with temp:
                    json.dump(data, temp, ensure_ascii=False, indent=2)
                os.replace(temp_path, manifest)
            finally:
                temp_path.unlink(missing_ok=True)
            _load.cache_clear()
    return str(root)


def contains(root: Path, path: Path) -> bool:
    scope = get(root)
    return scope.contains(path) if scope else path.is_relative_to(root)


def overlaps_reference(root: Path, reference: Path) -> bool:
    scope = get(root)
    return bool(scope and (
        any(folder.is_relative_to(reference) or reference.is_relative_to(folder) for folder in scope.folders)
        or any(file.is_relative_to(reference) for file in scope.files)
    ))


def list_files(root: Path, extensions: set[str]) -> list[Path]:
    if not root.is_dir():
        raise ValueError(f"文件夹不存在: {root}")

    def eligible(path: Path) -> bool:
        return path.suffix.lower() in extensions and not path.name.startswith("~$") and path.is_file()

    scope = get(root)
    if scope is None:
        return [path for path in root.rglob("*") if eligible(path)]

    found = set()
    for folder in scope.folders:
        if not folder.is_dir():
            raise ValueError(f"批量路径中的文件夹不存在：{folder}")
        if any(parent in scope.folders for parent in folder.parents):
            continue
        for path in folder.rglob("*"):
            if eligible(path):
                resolved = path.resolve()
                if scope.contains(resolved):
                    found.add(resolved)
    for file in scope.files:
        candidates = [p for p in file_variants(file) if p.is_file()]
        if not candidates:
            raise ValueError(f"批量路径中的文件不存在：{file}")
        if len(candidates) > 1:
            raise ValueError(f"同一目录中有多个带复检标记的同名文件，请先区分文件名：{file}")
        actual = candidates[0]
        if not eligible(actual):
            raise ValueError(f"此功能无法处理该文件：{actual}")
        resolved = actual.resolve()
        if not scope.contains(resolved):
            raise ValueError(f"批量路径中的文件已指向其他位置：{actual}")
        found.add(resolved)
    return list(found)
