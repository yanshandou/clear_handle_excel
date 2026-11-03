import os
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from core import data_dir, storage


def make_program(folder: Path, value: str | None = None) -> Path:
    (folder / "core").mkdir(parents=True)
    (folder / "core" / "storage.py").write_text("", encoding="utf-8")
    (folder / "web_clip_fixer.py").write_text("", encoding="utf-8")
    if value is not None:
        db = sqlite3.connect(folder / data_dir.DATABASE_NAME)
        db.execute("CREATE TABLE t (v TEXT)")
        db.execute("INSERT INTO t VALUES (?)", (value,))
        db.commit()
        db.close()
    return folder


def read_value(path: Path) -> str:
    db = sqlite3.connect(path)
    try:
        return db.execute("SELECT v FROM t").fetchone()[0]
    finally:
        db.close()


class DataDirMigrationTest(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.data = self.root / "appdata" / "clear_handle_excel"
        self.database = self.data / data_dir.DATABASE_NAME
        self.images = self.data / data_dir.IMAGES_NAME
        patcher = patch.object(data_dir, "_checked", set())
        patcher.start()
        self.addCleanup(patcher.stop)
        prev = patch.object(data_dir, "previous_dirs", lambda: [])
        prev.start()
        self.addCleanup(prev.stop)

    def test_previous_data_dir_comes_first(self):
        appdata = self.root / "LocalAppData" / "clear_handle_excel"
        appdata.mkdir(parents=True)
        db = sqlite3.connect(appdata / data_dir.DATABASE_NAME)
        db.execute("CREATE TABLE t(v TEXT)")
        db.execute("INSERT INTO t VALUES ('appdata')")
        db.commit()
        db.close()
        (appdata / data_dir.IMAGES_NAME).mkdir()
        (appdata / data_dir.IMAGES_NAME / "x.png").write_bytes(b"img")
        program = make_program(self.root / "Downloads" / "clear_handle_excel-main", "program")
        desktop = self.root / "Desktop" / "handle-excel" / "程序数据"
        data_dir.ensure_migrated(desktop / data_dir.DATABASE_NAME, desktop / data_dir.IMAGES_NAME, program, [appdata])
        self.assertEqual(read_value(desktop / data_dir.DATABASE_NAME), "appdata")
        self.assertEqual((desktop / data_dir.IMAGES_NAME / "x.png").read_bytes(), b"img")
        self.assertTrue((appdata / data_dir.DATABASE_NAME).exists(), "旧文件不删")

    def tearDown(self):
        self.temp.cleanup()

    def test_new_zip_folder_picks_up_old_sibling(self):
        old = make_program(self.root / "Downloads" / "clear_handle_excel-main", "old")
        (old / data_dir.IMAGES_NAME).mkdir()
        (old / data_dir.IMAGES_NAME / "a_before.png").write_bytes(b"png")
        new = make_program(self.root / "Downloads" / "clear_handle_excel-main (1)")
        data_dir.ensure_migrated(self.database, self.images, new)
        self.assertEqual(read_value(self.database), "old")
        self.assertEqual((self.images / "a_before.png").read_bytes(), b"png")
        self.assertTrue((old / data_dir.DATABASE_NAME).exists(), "旧文件不删")

    def test_nested_zip_folder_and_newest_wins(self):
        downloads = self.root / "Downloads"
        older = make_program(downloads / "v1" / "clear_handle_excel-main", "v1")
        newer = make_program(downloads / "v2" / "clear_handle_excel-main", "v2")
        past = time.time() - 3600
        os.utime(older / data_dir.DATABASE_NAME, (past, past))
        new = make_program(downloads / "v3" / "clear_handle_excel-main")
        data_dir.ensure_migrated(self.database, self.images, new)
        self.assertEqual(read_value(self.database), "v2")
        self.assertEqual(read_value(newer / data_dir.DATABASE_NAME), "v2")

    def test_own_folder_first_and_unrelated_folders_ignored(self):
        downloads = self.root / "Downloads"
        stranger = downloads / "other"
        stranger.mkdir(parents=True)
        db = sqlite3.connect(stranger / data_dir.DATABASE_NAME)
        db.execute("CREATE TABLE t (v TEXT)")
        db.execute("INSERT INTO t VALUES ('stranger')")
        db.commit()
        db.close()
        own = make_program(downloads / "current", "own")
        make_program(downloads / "sibling", "sibling")
        data_dir.ensure_migrated(self.database, self.images, own)
        self.assertEqual(read_value(self.database), "own")

    def test_existing_data_is_never_overwritten(self):
        self.data.mkdir(parents=True)
        db = sqlite3.connect(self.database)
        db.execute("CREATE TABLE t (v TEXT)")
        db.execute("INSERT INTO t VALUES ('kept')")
        db.commit()
        db.close()
        own = make_program(self.root / "Downloads" / "current", "own")
        data_dir.ensure_migrated(self.database, self.images, own)
        self.assertEqual(read_value(self.database), "kept")

    def test_wal_content_is_copied(self):
        old = make_program(self.root / "Downloads" / "old")
        db = sqlite3.connect(old / data_dir.DATABASE_NAME)
        db.execute("PRAGMA journal_mode=WAL")
        db.execute("PRAGMA wal_autocheckpoint=0")
        db.execute("CREATE TABLE t (v TEXT)")
        db.execute("INSERT INTO t VALUES ('in-wal')")
        db.commit()
        new = make_program(self.root / "Downloads" / "new")
        data_dir.ensure_migrated(self.database, self.images, new)
        db.close()
        self.assertEqual(read_value(self.database), "in-wal")

    def test_nothing_to_migrate_creates_folder(self):
        new = make_program(self.root / "fresh")
        data_dir.ensure_migrated(self.database, self.images, new)
        self.assertTrue(self.data.is_dir())
        self.assertFalse(self.database.exists())

    def test_storage_uses_data_dir(self):
        self.assertEqual(storage.DATABASE_PATH, data_dir.DATA_DIR / data_dir.DATABASE_NAME)
        self.assertNotEqual(storage.DATABASE_PATH.parent, data_dir.PROJECT_ROOT)


if __name__ == "__main__":
    unittest.main()
