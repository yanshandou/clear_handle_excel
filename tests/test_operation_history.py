import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch
from core import operation_history, storage

class OperationHistoryTests(unittest.TestCase):
    def test_counts_folders_and_restart(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "DATABASE_PATH", Path(directory) / "history.sqlite3"):
            first = operation_history.begin(Path(directory) / "first", ["scan"])
            operation_history.finish(first, "completed")
            second = operation_history.begin(Path(directory) / "first", ["scan", "autofix"])
            operation_history.finish(second, "failed")
            operation_history.begin(Path(directory) / "second", ["colors"])
            operation_history.interrupt_running()
            groups = {Path(g["folder"]).name: g["operations"] for g in operation_history.summary()["folders"]}
            self.assertEqual(groups["first"]["scan"]["count"], 2)
            self.assertEqual(groups["first"]["scan"]["status"], "failed")
            self.assertEqual(groups["first"]["autofix"]["count"], 1)
            self.assertEqual(groups["second"]["colors"]["status"], "interrupted")

    def test_done_for_folder_and_parent(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(storage, "DATABASE_PATH", Path(directory) / "history.sqlite3"):
            root = Path(directory) / "root"
            operation_history.finish(operation_history.begin(root, ["clean"]), "completed")
            operation_history.finish(operation_history.begin(root / "sub", ["colors"]), "failed")
            operation_history.finish(operation_history.begin(root / "sub", ["colors"]), "completed")
            operation_history.finish(operation_history.begin(root / "sub", ["links"]), "failed")
            operation_history.begin(root / "sub", ["hidden"])
            operation_history.finish(operation_history.begin(Path(directory) / "rootx", ["pdf"]), "completed")
            done = operation_history.done_for(root / "sub")
            self.assertTrue(done["clean"]["completed"])          # 在上级文件夹做过也算
            self.assertEqual(done["colors"]["count"], 2)
            self.assertTrue(done["colors"]["completed"])
            self.assertFalse(done["links"]["completed"])         # 只有出错的那次
            self.assertNotIn("hidden", done)                     # 还在跑 / 中断不算
            self.assertNotIn("pdf", done)                        # 名字相近的别的文件夹不算
            self.assertNotIn("colors", operation_history.done_for(root))  # 子文件夹做过不算整个文件夹
            self.assertEqual(operation_history.done_for(""), {})

if __name__ == "__main__":
    unittest.main()
