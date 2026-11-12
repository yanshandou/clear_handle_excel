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
REVIEW_TIMEOUT = 30 * 60


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
        self._candidates: list[Candidate] = []
        self._review_id = uuid.uuid4().hex
        self._cancelled = False

    @property
    def prepared_dir(self) -> Path:
        return Path(self._results.name) / "prepared"

    def _path(self, token: str) -> Path:
        return Path(self._results.name) / f"{token}.json"

    def status(self) -> tuple[list[dict[str, Any]], str, str]:
        with self._condition:
            pending = [{"token": token, "rel": item["rel"], "reason": item["review_reason"],
                        "missing": expected is None}
                       for token, (item, _, _, expected) in self._pending.items()]
            active = not self._cancelled and not self._closed and (pending or (self._candidates and not self._done))
            return pending, self._current, self._review_id if active else ""

    def result(self, token: str) -> dict[str, Any]:
        with self._condition:
            if token not in self._pending:
                raise ValueError("这份上传确认已结束，请回到上传页")
            handle = self._path(token).open(encoding="utf-8")
        with handle:
            return json.load(handle)

    def _skip_remaining(self, message: str) -> None:
        self._cancelled = True
        for item, *_ in self._candidates:
            if not item.get("status") and not item.get("review_confirmed"):
                item.update(status="skipped", error=message)
        self._pending.clear()
        self._condition.notify_all()

    def cancel(self, review_id: str) -> None:
        with self._condition:
            if review_id != self._review_id or self._cancelled or self._closed:
                raise ValueError("这批上传确认已结束，请回到上传页")
            self._skip_remaining("你取消了剩余的上传审核")

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
                require_review: bool = True, stage: Callable[[list[Candidate]], None] | None = None) -> Iterator[list[Candidate]]:
        if not require_review:
            if candidates:
                yield candidates
            return
        direct = [entry for entry in candidates if not entry[0].get("review_reason")]
        review = [entry for entry in candidates if entry[0].get("review_reason")]
        if review and stage is not None:
            stage(review)
            for entry in review:
                item = entry[0]
                if item.get("status"):
                    continue
                if item.pop("_confirmed_original", False) and item["_prepared_sha256"] == item["src_sha256"]:
                    item["review_confirmed"] = True
                    item.pop("review_reason")
                    direct.append(entry)
            review = [entry for entry in review if not entry[0].get("status") and entry[0].get("review_reason")]
        if review:
            with self._condition:
                self._candidates = review
            self._worker = threading.Thread(target=self._compare, args=(review, verify), daemon=True,
                                            name="上传前 Excel 对比")
            self._worker.start()
        if direct:
            yield direct
        if not review:
            return
        while True:
            with self._condition:
                self._condition.wait_for(lambda: self._approved or self._done or self._cancelled)
                if self._done and self._pending and not self._approved:
                    if not self._condition.wait_for(lambda: self._approved or not self._pending, REVIEW_TIMEOUT):
                        self._skip_remaining("等待上传确认超时，未上传")
                if not self._approved:
                    return
                approved, self._approved = self._approved, []
            yield approved

    def _compare(self, candidates: list[Candidate], verify: Callable[[Candidate], None]) -> None:
        try:
            for candidate in candidates:
                item, local, remote, expected = candidate
                with self._condition:
                    if self._closed or self._cancelled:
                        return
                    self._current = item["rel"]
                token = uuid.uuid4().hex
                try:
                    verify(candidate)
                    prepared = Path(item["_temp"]) if item.get("_temp") else local
                    diff = excel_diff.diff(prepared, remote) if expected is not None else None
                    verify(candidate)
                    with self._path(token).open("w", encoding="utf-8") as handle:
                        json.dump({"rel": item["rel"], "reason": item["review_reason"], "diff": diff,
                                   "pdf_only": self._pdf_only, "prepared": prepared != local},
                                  handle, ensure_ascii=False)
                except Exception as err:
                    with self._condition:
                        if self._cancelled or self._closed:
                            return
                        item.update(status="failed", error=f"无法确认上传版本：{err}")
                    debug_log.exception(err, f"上传前 Excel 对比 {item['rel']}")
                    continue
                with self._condition:
                    if self._closed or self._cancelled:
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
