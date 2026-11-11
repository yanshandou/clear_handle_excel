"""Choose a reference directory for each workbook and look up its exact filename."""

from dataclasses import dataclass
import os
from pathlib import Path
from typing import Callable

from core import workbook_scope, work_task
from core.common import strip_review_prefix


def _directory(value: str) -> Path:
    if len(value.strip().splitlines()) != 1:
        raise ValueError("每个参考目录请填写一个文件夹路径")
    work_task.require_folder(value, "reference")
    path = workbook_scope.parse_path(value)
    work_task.require_folder(str(path), "reference")
    if not path.is_dir() or workbook_scope.get(path):
        raise ValueError(f"参考文件夹不存在或无效：{path}")
    return path


@dataclass(frozen=True)
class ReferenceFolders:
    default: Path | None
    overrides: dict[Path, Path]

    @classmethod
    def from_payload(cls, payload: dict) -> "ReferenceFolders":
        default = payload.get("old_folder", "")
        overrides = payload.get("reference_folders", {})
        if not isinstance(default, str) or not isinstance(overrides, dict):
            raise ValueError("参考文件夹设置格式不正确")
        by_folder = {}
        for folder, reference in overrides.items():
            if not isinstance(folder, str) or not isinstance(reference, str) or not reference.strip():
                raise ValueError("请为处理文件夹填写有效的参考文件夹路径")
            work = workbook_scope.parse_path(folder)
            target = _directory(reference)
            if work in by_folder and by_folder[work] != target:
                raise ValueError(f"同一个处理文件夹指定了不同的参考目录：{work}")
            by_folder[work] = target
        return cls(_directory(default) if default.strip() else None, by_folder)

    def __bool__(self) -> bool:
        return self.default is not None or bool(self.overrides)

    def payload(self) -> dict:
        return {"old_folder": str(self.default or ""),
                "reference_folders": {str(folder): str(reference) for folder, reference in self.overrides.items()}}

    def directories(self, root: Path) -> set[Path]:
        scope = workbook_scope.get(root)
        selected = set(scope.folders) if scope else {root}
        for folder in self.overrides:
            if folder not in selected:
                raise ValueError(f"已指定参考目录的处理文件夹不在当前批次中：{folder}")
        directories = set(self.overrides.values())
        if self.default is not None:
            directories.add(self.default)
        for directory in directories:
            overlap = (workbook_scope.overlaps_reference(root, directory) if scope else
                       root.is_relative_to(directory) or directory.is_relative_to(root))
            if overlap:
                raise ValueError(f"处理路径与参考文件夹不能相同或互相包含：{directory}")
        return directories


def _filename(name: str) -> str:
    return os.path.normcase(strip_review_prefix(Path(name).name))


class ReferenceIndex:
    def __init__(self, root: Path, options: ReferenceFolders, extensions: set[str]):
        self.options = options
        self.root = root
        self.scope = workbook_scope.get(root)
        self.directories = options.directories(root)
        self.extensions = extensions
        self.by_name: dict[Path, dict[str, set[Path]]] = {}

    def _index(self, directory: Path) -> dict[str, set[Path]]:
        if directory in self.by_name:
            return self.by_name[directory]
        index: dict[str, set[Path]] = {}
        for path in workbook_scope.list_files(directory, self.extensions):
            actual = path.resolve()
            selected = self.scope.contains(actual) if self.scope else actual.is_relative_to(self.root)
            if actual.is_relative_to(directory) and not selected:
                index.setdefault(_filename(path.name), set()).add(actual)
        self.by_name[directory] = index
        return index

    @property
    def files(self) -> set[Path]:
        return {path for directory in self.directories for matches in self._index(directory).values() for path in matches}

    def __str__(self) -> str:
        return str(self.options.default or "")

    def match(self, name: str) -> tuple[Path | None, str]:
        relative = Path(name)
        current = self.scope.source_path(relative) if self.scope else self.root / relative
        directory = self.options.default
        if current is not None:
            directory = next((self.options.overrides[parent] for parent in current.parents
                              if parent in self.options.overrides), directory)
        if directory is None:
            return None, "这份 Excel 未指定参考文件夹"
        candidates = self._index(directory).get(_filename(name), set())
        if len(candidates) == 1:
            return next(iter(candidates)), ""
        if candidates:
            return None, f"参考文件夹中有 {len(candidates)} 份同名文件，请指定更具体的参考文件夹：{directory}"
        return None, f"参考文件夹中没有同名 Excel：{directory}"


def prepare(root: Path, options: str | ReferenceFolders,
            legacy_list: Callable[[str], tuple[Path, list[Path]]], extensions: set[str], *,
            disjoint: bool = False) -> tuple[Path | ReferenceIndex, list[Path]]:
    if isinstance(options, str) and workbook_scope.get(root):
        options = ReferenceFolders.from_payload({"old_folder": options})
    if isinstance(options, ReferenceFolders):
        index = ReferenceIndex(root, options, extensions)
        return index, list(index.files)
    old_root, paths = legacy_list(options)
    if root == old_root or workbook_scope.overlaps_reference(root, old_root) or (
        disjoint and (root in old_root.parents or old_root in root.parents)
    ):
        raise ValueError("当前文件夹与备份文件夹不能相同或互相包含，请分别选择两份文件所在的目录")
    return old_root, paths


def unmatched(index: Path | ReferenceIndex | None, name: str, fallback: str) -> str:
    return index.match(name)[1] if isinstance(index, ReferenceIndex) else fallback
