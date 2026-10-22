import base64
import importlib
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import fix_images, storage


class StateImageTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        p = patch.object(storage, "DATABASE_PATH", self.root / "t.sqlite3")
        p.start()
        self.addCleanup(p.stop)

    def blob_rows(self):
        with sqlite3.connect(storage.DATABASE_PATH) as conn:
            return dict(conn.execute("SELECT key, data FROM blobs WHERE name = 'w'").fetchall())

    def test_images_are_stored_once_and_merged_back_on_load(self):
        big = "data:image/png;base64," + "A" * 1000
        state = {"pages": {"p1": {"id": "p1", "image": big}}, "items": {"i1": {"id": "i1", "image": ""}}}
        kwargs = {"blob_field": "image", "blob_collections": ("pages", "items")}
        storage.save_snapshot("w", state, **kwargs)
        with sqlite3.connect(storage.DATABASE_PATH) as conn:
            payload = conn.execute("SELECT payload FROM snapshots WHERE name = 'w'").fetchone()[0]
        self.assertNotIn("AAAA", payload)
        self.assertEqual(storage.load_snapshot("w"), state)

        with sqlite3.connect(storage.DATABASE_PATH) as conn:
            conn.execute("UPDATE blobs SET data = 'marker'")
        storage.save_snapshot("w", state, **kwargs)  # 截图没变：不再写
        self.assertEqual(list(self.blob_rows().values()), ["marker"])

        del state["pages"]["p1"]
        storage.save_snapshot("w", state, **kwargs)
        self.assertEqual(self.blob_rows(), {})
        self.assertEqual(storage.load_snapshot("w"), state)


class StateSnapshotTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        p = patch.object(storage, "DATABASE_PATH", self.root / "t.sqlite3")
        p.start()
        self.addCleanup(p.stop)
        from core import reporter, workflow
        self.workflow = importlib.reload(workflow)
        p = patch.object(reporter, "RECORD_PATH", self.root / "record.md")
        p.start()
        self.addCleanup(p.stop)
        p = patch.object(fix_images, "IMAGES_DIR", self.root / "_report_images")
        p.start()
        self.addCleanup(p.stop)

    def test_poll_carries_image_addresses_not_images(self):
        png = b"\x89PNG fake"
        image = "data:image/png;base64," + base64.b64encode(png).decode()
        wb = str(self.root / "a.xlsx")
        self.workflow.STATE.update(
            busy=False, job=None,
            pages={"p1": {"id": "p1", "wb_path": wb, "image": image}},
            items={"i1": {"id": "i1", "wb_path": wb, "image": image, "status": "pending"}},
        )
        snap = self.workflow.get_state_snapshot()
        url = snap["pages"]["p1"]["image"]
        self.assertTrue(url.startswith("/api/image/page/p1?v="))
        self.assertTrue(snap["items"]["i1"]["image"].startswith("/api/image/item/i1?v="))
        self.assertEqual(self.workflow.state_image("page", "p1"), png)
        self.assertIsNone(self.workflow.state_image("page", "missing"))
        self.assertIsNone(self.workflow.state_image("other", "p1"))

        self.workflow.STATE["pages"]["p1"]["image"] = image + "AAAA"
        self.assertNotEqual(self.workflow.get_state_snapshot()["pages"]["p1"]["image"], url)  # 换了图地址就变


if __name__ == "__main__":
    unittest.main()
