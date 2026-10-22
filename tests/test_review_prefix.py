import unittest
from pathlib import Path

from core.common import (
    REVIEW_PREFIX,
    has_review_prefix,
    review_marked_name,
    review_name_variants,
    strip_review_prefix,
    workbook_rel_name,
)


class ReviewPrefixTests(unittest.TestCase):
    def test_new_prefix_displays_on_japanese_windows(self):
        # 日文 Windows 10 的字体没有 🟡，前缀必须是 Shift_JIS 里就有的字
        REVIEW_PREFIX.encode("cp932")
        self.assertEqual(review_marked_name("a.xlsx"), "★a.xlsx")

    def test_legacy_prefix_is_still_recognized_and_stripped(self):
        self.assertTrue(has_review_prefix("🟡a.xlsx"))
        self.assertEqual(review_marked_name("🟡a.xlsx"), "🟡a.xlsx")
        self.assertEqual(strip_review_prefix("🟡a.xlsx"), "a.xlsx")
        self.assertEqual(strip_review_prefix("★a.xlsx"), "a.xlsx")
        self.assertEqual(workbook_rel_name(Path("/r/sub/🟡a.xlsx"), "/r"), "sub/a.xlsx")
        self.assertFalse(has_review_prefix("a.xlsx"))

    def test_name_variants_cover_both_prefixes(self):
        self.assertEqual(review_name_variants("🟡a.xlsx"), ["★a.xlsx", "🟡a.xlsx"])


if __name__ == "__main__":
    unittest.main()
