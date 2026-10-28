"""备份文件名只能作为明确标示、双向唯一的人工核对线索。"""
import unittest
from pathlib import Path

from core.backup_compare import pair_files


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


if __name__ == "__main__":
    unittest.main()
