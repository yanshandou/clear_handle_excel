import copy
import io
import json
import tempfile
import unittest
from email.message import Message
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, workflow as wf


class WorkbenchSnapshotTests(unittest.TestCase):
    def setUp(self):
        directory = tempfile.TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.root = Path(directory.name)
        self.book = self.root / "book.xlsx"
        self.book.write_bytes(b"workbook")
        state = {**wf.STATE, "folder": str(self.root), "items": {}, "pages": {}, "busy": False,
                 "job": None, "scan_progress": {}, "logs": [],
                 "batch_review": {"cards": {}, "rows": {}, "ignored": []}}
        for target, name, value in ((wf, "STATE", state), (fix_images, "IMAGES_DIR", self.root / "images")):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def card(self, row=1):
        item = {"id": f"item-{row}", "seq": row, "kind": "clip", "wb_path": str(self.book),
                "wb_name": self.book.name, "sheet_name": "Sheet1", "cell_addr": f"A{row}",
                "excel_row": row, "text": "review", "status": "pending", "location_confirmed": True,
                "page_index": 0, "bbox": [0, 0, 100, 20]}
        wf.STATE["items"][item["id"]] = item
        card = batch._card_for([item])
        card.update(finalized=True, waiting=False)
        return card

    def test_snapshot_reuses_index_for_old_and_handed_over_cards(self):
        card = self.card()
        card.update(handed_over=True)
        card.pop("problems")
        card.pop("error_types")
        with patch.object(batch, "_items_at", side_effect=AssertionError("unexpected full scan")):
            result = batch.snapshot()
        self.assertTrue(result["mode"]["to_auto"])
        self.assertEqual(len(result["cards"][0]["problems"]), 1)
        self.assertEqual(result["cards"][0]["auto_skip"], "")

    def test_same_file_version_is_read_once_per_snapshot_and_refreshed_next_time(self):
        self.card(1)
        self.card(2)
        current = wf._file_version(self.book)
        batch._sheet_checks()[str(self.book)] = {"*": current}
        with patch.object(wf, "_file_version", wraps=wf._file_version) as version:
            first = batch.snapshot()
            self.assertEqual(version.call_count, 1)
            self.assertTrue(all(card["auto_skip"] for card in first["cards"]))
            self.book.write_bytes(b"edited workbook")
            second = batch.snapshot()
            self.assertEqual(version.call_count, 2)
            self.assertTrue(all(not card["auto_skip"] for card in second["cards"]))

    def test_failed_version_lookup_is_cached_only_within_one_snapshot(self):
        self.card(1)
        self.card(2)
        with patch.object(wf, "_file_version", side_effect=OSError("missing")) as version:
            self.assertTrue(all(not card["auto_skip"] for card in batch.snapshot()["cards"]))
            self.assertEqual(version.call_count, 1)
            batch.snapshot()
            self.assertEqual(version.call_count, 2)

    def test_folder_resolution_is_shared_with_mode_and_refreshed_each_request(self):
        self.card(1)["handed_over"] = True
        self.card(2)["handed_over"] = True
        with patch.object(wf, "in_current_folder", wraps=wf.in_current_folder) as contains:
            self.assertTrue(batch.snapshot()["mode"]["to_auto"])
            self.assertEqual(contains.call_count, 1)
            batch.snapshot()
            self.assertEqual(contains.call_count, 2)

    def test_explicit_empty_items_preserve_optional_helper_contract(self):
        card = self.card()
        with patch.object(batch, "_items_at", side_effect=AssertionError("unexpected fallback")):
            self.assertTrue(batch._can_highlight(card, []))
        wf.STATE["items"]["item-1"]["location_confirmed"] = False
        self.assertFalse(batch._can_highlight(card))

    def test_filtered_snapshot_keeps_nested_data_isolated_in_both_directions(self):
        card = self.card()
        card["context"] = {"image_file": "missing.png", "bbox": [0, 0, 10, 20]}
        original = copy.deepcopy(card)
        result = batch.snapshot()["cards"][0]
        self.assertNotIn("signatures", result)
        self.assertNotIn("item_ids", result)
        self.assertEqual(card, original)
        card["fix_images"]["before"] = "new.png"
        card["problems"][0]["reason"] = "changed after snapshot"
        self.assertEqual(result["fix_images"], {})
        self.assertNotEqual(result["problems"][0]["reason"], "changed after snapshot")
        result["context"]["bbox"][0] = 99
        self.assertEqual(card["context"]["bbox"][0], 0)

    def test_shared_image_is_probed_once_and_not_cached_across_requests(self):
        for row in (1, 2):
            self.card(row)["context"] = {"image_file": "shared.png"}
        with patch.object(fix_images, "image_path", return_value=Path("shared.png")) as probe, \
                patch.object(fix_images, "version", return_value="version"):
            self.assertEqual(batch.snapshot()["cards"][0]["context"]["image"], "/api/batch/image/shared.png?v=version")
            self.assertEqual(probe.call_count, 1)
            batch.snapshot()
            self.assertEqual(probe.call_count, 2)


class WorkbenchConditionalHttpTests(unittest.TestCase):
    def request(self, validator=""):
        from web_clip_fixer import WebHandler
        handler = object.__new__(WebHandler)
        handler.path = "/api/batch"
        handler.headers = Message()
        if validator:
            handler.headers["If-None-Match"] = validator
        handler.wfile = io.BytesIO()
        handler.send_response = Mock()
        headers = {}
        handler.send_header = lambda key, value: headers.update({key: value})
        handler.end_headers = Mock()
        handler._handle_get()
        return handler.send_response.call_args.args[0], headers, handler.wfile.getvalue()

    def test_conditional_snapshot_uses_actual_json_bytes_and_returns_empty_304(self):
        with patch.object(batch, "snapshot", return_value={"cards": [], "folder": "测试"}) as snapshot:
            status, headers, body = self.request()
            self.assertEqual(status, 200)
            self.assertEqual(json.loads(body), snapshot.return_value)
            etag = headers["ETag"]
            self.assertEqual(headers["Content-Length"], str(len(body)))
            self.assertEqual(headers["Cache-Control"], "no-store")
            status, headers, body = self.request(etag)
            self.assertEqual((status, body, headers["ETag"]), (304, b"", etag))
            self.assertNotIn("Content-Length", headers)
            snapshot.return_value = {"cards": [{"color": "blue"}], "folder": "测试"}
            status, headers, body = self.request(etag)
            self.assertEqual(status, 200)
            self.assertNotEqual(headers["ETag"], etag)

    def test_conditional_request_accepts_weak_and_multiple_validators(self):
        with patch.object(batch, "snapshot", return_value={"cards": []}):
            etag = self.request()[1]["ETag"]
            self.assertEqual(self.request(f'"other", W/{etag}')[0], 304)
            self.assertEqual(self.request("*")[0], 304)


if __name__ == "__main__":
    unittest.main()
