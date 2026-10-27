"""Isolated synthetic snapshot benchmark: python tests/benchmark_workbench.py."""
import copy
import json
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from core import batch, workflow as wf


def measure(count=1000, files=80, samples=7):
    saved = copy.deepcopy(wf.STATE)
    try:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            paths = [root / f"book-{i}.xlsx" for i in range(files)]
            for path in paths:
                path.write_bytes(b"synthetic workbook")
            cards, items = {}, {}
            for i in range(count):
                path = str(paths[i % files])
                item = {"id": f"item-{i}", "seq": i, "kind": "clip", "wb_path": path,
                        "wb_name": Path(path).name, "sheet_name": "Sheet1", "cell_addr": f"A{i + 1}",
                        "excel_row": i + 1, "text": "Synthetic review text " * 5,
                        "status": "manual_review", "location_confirmed": True,
                        "page_index": 0, "bbox": [0, 0, 100, 20]}
                items[item["id"]] = item
                cards[str(i)] = {**item, "id": str(i), "color": "yellow", "finalized": True,
                                 "waiting": False, "verified": False, "fix_images": {},
                                 "signatures": [{"text": item["text"], "bbox": [0, 0, 100, 20]}] * 5,
                                 "item_ids": [item["id"]]}
            wf.STATE.update(folder=directory, items=items, pages={}, busy=False, job=None,
                            scan_progress={}, logs=[], batch_review={"cards": cards, "rows": {}, "ignored": []})
            results = {}
            for name, handed in (("idle_yellow", False), ("handed_over", True)):
                for card in cards.values():
                    card["handed_over"] = handed
                batch.snapshot()
                timings = []
                for _ in range(samples):
                    started = time.perf_counter()
                    snapshot = batch.snapshot()
                    timings.append((time.perf_counter() - started) * 1000)
                with patch.object(batch, "_items_at", wraps=batch._items_at) as scans, \
                        patch.object(wf, "_file_version", wraps=wf._file_version) as versions, \
                        patch.object(wf, "in_current_folder", wraps=wf.in_current_folder) as folders:
                    batch.snapshot()
                results[name] = {"cards": count, "files": files,
                                 "median_ms": round(statistics.median(timings), 2),
                                 "p95_ms": round(sorted(timings)[-1], 2),
                                 "body_bytes": len(json.dumps(snapshot, ensure_ascii=False).encode()),
                                 "items_at_calls": scans.call_count, "file_version_calls": versions.call_count,
                                 "folder_checks": folders.call_count}
            return results
    finally:
        wf.STATE.clear()
        wf.STATE.update(saved)


if __name__ == "__main__":
    if "--baseline" in sys.argv:
        source = subprocess.check_output(["git", "show", "940e58a:core/batch.py"],
                                         cwd=Path(__file__).resolve().parents[1], encoding="utf-8")
        # Run the original implementation only in this isolated benchmark process.
        exec(compile(source, "baseline/core/batch.py", "exec"), batch.__dict__)
    print(json.dumps(measure(), indent=2))
