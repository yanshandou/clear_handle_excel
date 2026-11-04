import json
import os
import shutil
import tempfile
import uuid
from datetime import datetime, timezone
from pathlib import Path

from step1_copy_files import (
    LOCAL_ROOT,
    SCRIPT_DIR,
    configure_stdio,
    copy_allocations,
    file_hash,
    load_allocations,
)


ONEDRIVE_ROOT = r"EDIT_ONEDRIVE_SYNC_FOLDER"
LOG_PATH = SCRIPT_DIR / "onedrive-overwrite-log.jsonl"


def write_log(record):
    record["time"] = datetime.now(timezone.utc).isoformat()
    with LOG_PATH.open("a", encoding="utf-8") as log_file:
        log_file.write(json.dumps(record, ensure_ascii=False) + "\n")
        log_file.flush()
        os.fsync(log_file.fileno())


def overwrite_with_log(source, destination):
    temporary_path = None
    operation_id = uuid.uuid4().hex
    try:
        with source.open("rb") as source_file:
            with tempfile.NamedTemporaryFile(dir=destination.parent, delete=False) as temporary_file:
                temporary_path = temporary_file.name
                shutil.copyfileobj(source_file, temporary_file)
        write_log({
            "id": operation_id,
            "status": "pending",
            "source": str(source),
            "destination": str(destination),
            "old_sha256": file_hash(destination).hex(),
            "new_sha256": file_hash(Path(temporary_path)).hex(),
        })
        try:
            os.replace(temporary_path, destination)
        except OSError as error:
            write_log({"id": operation_id, "status": "failed", "error": str(error)})
            raise
        temporary_path = None
        write_log({"id": operation_id, "status": "completed"})
    finally:
        if temporary_path is not None:
            os.unlink(temporary_path)


def main():
    configure_stdio()
    try:
        copied, total = copy_allocations(
            LOCAL_ROOT, ONEDRIVE_ROOT, load_allocations(), "destination", overwrite_with_log
        )
        print(f"第二步完成：共 {total} 个文件，已复制或覆盖 {copied} 个文件到 OneDrive 同步文件夹。")
        print(f"覆盖记录（发生覆盖时生成）：{LOG_PATH}")
        print("请检查 OneDrive 客户端，确认云端同步已完成。")
    except (OSError, ValueError, KeyError) as error:
        print(f"第二步失败：{error}")
        print("复制可能未完成；请解决问题后重新运行。")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())