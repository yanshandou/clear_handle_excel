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

if __name__ == "__main__":
    unittest.main()
