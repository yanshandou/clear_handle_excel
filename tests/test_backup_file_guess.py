"""备份文件名只能作为明确标示、双向唯一的人工核对线索。"""
import unittest
import tempfile
from pathlib import Path
from unittest.mock import patch

from core.backup_compare import pair_files
from core import data_dir, workbook_scope, work_task
from core.reference_folders import ReferenceFolders, ReferenceIndex


class BackupFileGuessTests(unittest.TestCase):
    def setUp(self):
        self.root = Path("backup")
        self.inferred = set()

    def pairs(self, paths, names, current_root=None):
        return pair_files(self.root, [self.root / path for path in paths], names, current_root,
                          inferred=self.inferred)

    def test_extension_change_and_common_backup_suffixes_are_marked_as_inferred(self):
        for old, new in (("book.xls", "book.xlsx"), ("book_backup.xlsx", "book.xlsx"),
                         ("book_20251002.xlsx", "book.xlsx"), ("book (備份).xlsx", "book.xlsx"),
                         ("book.xlsx", "book_2025-10-02.xlsm"), ("❓book.xlsx", "book.xlsx")):
            with self.subTest(old=old, new=new):
                self.inferred.clear()
                self.assertEqual(self.pairs([old], [new])[new], self.root / old)
                self.assertEqual(self.inferred, {new})

    def test_precise_match_wins_and_is_never_marked_as_inferred(self):
        pairs = self.pairs(["book.xlsx", "book_backup.xlsx"], ["book.xlsx"])
        self.assertEqual(pairs["book.xlsx"], self.root / "book.xlsx")
        self.assertEqual(self.inferred, set())

    def test_old_and_current_candidates_must_be_unique(self):
        pairs = self.pairs(["book_backup.xlsx", "book_old.xlsx"], ["book.xlsx"])
        self.assertIsNone(pairs["book.xlsx"])
        self.assertEqual(self.inferred, set())
        pairs = self.pairs(["book_backup.xlsx"], ["a/book.xlsx", "b/book.xlsx"])
        self.assertEqual(pairs, {"a/book.xlsx": None, "b/book.xlsx": None})

    def test_relative_directory_and_current_subfolder_are_preferred(self):
        pairs = self.pairs(["001/book_backup.xlsx", "002/book_backup.xlsx"], ["book.xlsx"], Path("current/001"))
        self.assertEqual(pairs["book.xlsx"], self.root / "001/book_backup.xlsx")
        pairs = self.pairs(["a/book_backup.xlsx", "b/book_backup.xlsx"], ["a/book.xlsx", "b/book.xlsx"])
        self.assertEqual(pairs, {"a/book.xlsx": self.root / "a/book_backup.xlsx", "b/book.xlsx": self.root / "b/book_backup.xlsx"})

    def test_precise_candidate_reserved_even_when_precise_pair_is_ambiguous(self):
        pairs = self.pairs(["book.xlsx", "book_backup.xlsx"], ["a/book.xlsx", "b/book.xlsx"])
        self.assertEqual(pairs, {"a/book.xlsx": None, "b/book.xlsx": None})
        self.assertEqual(self.inferred, set())

    def test_unrelated_or_invalid_date_names_are_not_candidates(self):
        for old in ("other_backup.xlsx", "book_20251340.xlsx", "book_0114.xlsx", "book.pdf"):
            with self.subTest(old=old):
                self.assertIsNone(self.pairs([old], ["book.xlsx"])["book.xlsx"])

    def test_guesses_are_disabled_until_caller_explicitly_handles_them(self):
        self.assertIsNone(pair_files(self.root, [self.root / "book_backup.xlsx"], ["book.xlsx"])["book.xlsx"])

    def test_many_explicit_files_share_a_recursive_reference_directory(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            reference = root / "reference"
            paths, names, expected = [], [], {}
            for i in range(50):
                name = f"workbook-{i}.xlsx"
                current = root / "current" / f"work-{i}" / name
                backup = reference / "originals" / str(i % 3) / name
                current.parent.mkdir(parents=True)
                backup.parent.mkdir(parents=True, exist_ok=True)
                current.touch()
                backup.touch()
                paths.append(current)
                label = f"work-{i}/{name}"
                names.append(label)
                expected[label] = backup
            (reference / "missing_backup.xlsx").touch()

            with patch.object(data_dir, "DATA_DIR", root / "data"):
                scope = Path(workbook_scope.resolve_input("\n".join(map(str, paths)), work_task.require_folder))
                options = ReferenceFolders.from_payload({"old_folder": f'"{reference}"'})
                index = ReferenceIndex(scope, options, {".xlsx"})
                self.assertEqual(pair_files(index, [], names, scope, inferred=self.inferred), expected)
                self.assertIsNone(index.match("work-0/missing.xlsx")[0])
                self.assertIn("没有同名 Excel", index.match("work-0/missing.xlsx")[1])


if __name__ == "__main__":
    unittest.main()
