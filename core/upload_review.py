"""Run workbook comparisons while approved files are being uploaded."""
from __future__ import annotations

import json
import tempfile
import threading
import uuid
from pathlib import Path
from typing import Any, Callable, Iterator

from core import debug_log, excel_diff


Candidate = tuple[dict[str, Any], Path, Path, str | None]


class UploadReview:
    def __init__(self, *, pdf_only: bool = False) -> None:
        self._condition = threading.Condition()
        self._pending: dict[str, Candidate] = {}
        self._approved: list[Candidate] = []
        self._done = False
        self._closed = False
        self._current = ""
        self._pdf_only = pdf_only
        self._worker: threading.Thread | None = None
        self._results = tempfile.TemporaryDirectory(prefix="excel-upload-diff-", ignore_cleanup_errors=True)

    def _path(self, token: str) -> Path:
        return Path(self._results.name) / f"{token}.json"

    def status(self) -> tuple[list[dict[str, Any]], str]:
        with self._condition:
            pending = [{"token": token, "rel": item["rel"], "reason": item["review_reason"],
                        "missing": expected is None}
                       for token, (item, _, _, expected) in self._pending.items()]
            return pending, self._current

    def result(self, token: str) -> dict[str, Any]:
        with self._condition:
            if token not in self._pending:
                raise ValueError("这份上传确认已结束，请回到上传页")
            result = self._path(token).read_text(encoding="utf-8")
        return json.loads(result)

    def decide(self, token: str, upload: bool) -> None:
        with self._condition:
            candidate = self._pending.pop(token, None)
            if candidate is None:
                raise ValueError("这份上传确认已结束，请回到上传页")
            item = candidate[0]
            if upload:
                item["review_confirmed"] = True
                self._approved.append(candidate)
            else:
                item.update(status="skipped", error="你选择了跳过上传")
            self._condition.notify_all()

    def batches(self, candidates: list[Candidate], verify: Callable[[Candidate], None], *,
                require_review: bool = True) -> Iterator[list[Candidate]]:
        if not require_review:
            if candidates:
                yield candidates
            return
        direct = [entry for entry in candidates if not entry[0].get("review_reason")]
        review = [entry for entry in candidates if entry[0].get("review_reason")]
        if review:
            self._worker = threading.Thread(target=self._compare, args=(review, verify), daemon=True,
                                            name="上传前 Excel 对比")
            self._worker.start()
        if direct:
            yield direct
        if not review:
            return
        while True:
            with self._condition:
                self._condition.wait_for(lambda: self._approved or (self._done and not self._pending))
                if not self._approved:
                    return
                approved, self._approved = self._approved, []
            yield approved

    def _compare(self, candidates: list[Candidate], verify: Callable[[Candidate], None]) -> None:
        try:
            for candidate in candidates:
                item, local, remote, expected = candidate
                with self._condition:
                    if self._closed:
                        return
                    self._current = item["rel"]
                token = uuid.uuid4().hex
                try:
                    verify(candidate)
                    diff = excel_diff.diff(local, remote) if expected is not None else None
                    verify(candidate)
                    with self._path(token).open("w", encoding="utf-8") as handle:
                        json.dump({"rel": item["rel"], "reason": item["review_reason"], "diff": diff,
                                   "pdf_only": self._pdf_only},
                                  handle, ensure_ascii=False)
                except Exception as err:
                    item.update(status="failed", error=f"无法确认上传版本：{err}")
                    debug_log.exception(err, f"上传前 Excel 对比 {item['rel']}")
                    continue
                with self._condition:
                    if self._closed:
                        return
                    self._pending[token] = candidate
                    self._condition.notify_all()
        finally:
            with self._condition:
                self._current = ""
                self._done = True
                self._condition.notify_all()

    def close(self) -> None:
        with self._condition:
            self._closed = True
            self._pending.clear()
        if self._worker is not None and self._worker.ident is not None:
            self._worker.join()
        self._results.cleanup()
