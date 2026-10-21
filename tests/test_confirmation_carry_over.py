import unittest

from core.workflow import MANUAL_CONFIRMED_REASON, MANUALLY_FIXED_REASON, _merge_redetected


def problem(item_id, text, occurrence=None, sheet="Sheet1", status="pending", seq=1):
    item = {"id": item_id, "seq": seq, "wb_path": "C:/excel/book.xlsx", "sheet_name": sheet, "text": text,
            "occurrence": occurrence, "status": status, "page_id": f"page_{item_id}"}
    if status == "accepted":
        item["manual_confirmation"] = {"confirmed_at": "2026-10-21T10:00:00", "previous_status": "pending"}
    return item


def detected(*items):
    return {item["id"]: item for item in items}


def merge(new, previous, mark_new=True, keep_unmatched=True):
    return _merge_redetected(new, previous, mark_new=mark_new, keep_unmatched=keep_unmatched)


class MergeRedetectedTests(unittest.TestCase):
    def test_same_text_and_occurrence_keeps_id_and_skip_regardless_of_position(self):
        new = detected(problem("n1", "お客様　一郎", [0, 3], seq=90), problem("n2", "お客様　三郎", [0, 1], seq=91))
        merged = merge(new, [problem("old1", "お客様　一郎", [0, 3], status="accepted", seq=5)])

        self.assertEqual(merged["old1"]["status"], "accepted")
        self.assertEqual(merged["old1"]["seq"], 5)
        self.assertEqual(merged["old1"]["review_reason"], MANUAL_CONFIRMED_REASON)
        self.assertEqual(merged["old1"]["manual_confirmation"]["confirmed_at"], "2026-10-21T10:00:00")
        self.assertFalse(merged["old1"]["is_new"])
        # 上次没有的位置：新发现
        self.assertEqual((merged["n2"]["status"], merged["n2"]["is_new"]), ("pending", True))

    def test_still_clipped_problem_stays_pending_with_same_id(self):
        merged = merge(detected(problem("n1", "k", [0, 1])), [problem("old1", "k", [0, 1], status="manual_review")])
        self.assertEqual(list(merged), ["old1"])
        self.assertEqual((merged["old1"]["status"], merged["old1"]["is_new"]), ("pending", False))

    def test_problem_no_longer_detected_becomes_fixed_record(self):
        previous = [
            problem("old1", "k", [0, 1]),
            problem("old2", "i", [0, 1], status="accepted"),
            problem("old3", "x", [0, 1], status="fixed"),
        ]
        merged = merge({}, previous)
        self.assertEqual((merged["old1"]["status"], merged["old1"]["page_id"]), ("fixed", None))
        self.assertEqual(merged["old1"]["review_reason"], MANUALLY_FIXED_REASON)
        self.assertEqual(merged["old2"]["status"], "accepted")
        self.assertEqual(merged["old3"]["status"], "fixed")
        # 完整重新扫描不保留未再检出的旧记录
        self.assertEqual(merge({}, previous, keep_unmatched=False), {})

    def test_repeated_text_follows_occurrence_order(self):
        # 同一列有 3 个「k」，只跳过了第 2 个
        new = detected(*(problem(f"n{i}", "k", [i, 3]) for i in range(3)))
        merged = merge(new, [problem("old", "k", [1, 3], status="accepted")])
        self.assertEqual(merged["old"]["status"], "accepted")
        self.assertEqual(sorted(item["status"] for item in merged.values()), ["accepted", "pending", "pending"])

    def test_changed_count_of_same_text_is_treated_as_new_problem(self):
        merged = merge(detected(problem("n1", "k", [1, 4])), [problem("old", "k", [1, 3], status="accepted")])
        self.assertEqual((merged["n1"]["status"], merged["n1"]["is_new"]), ("pending", True))
        self.assertEqual(merged["old"]["status"], "accepted")  # 旧记录保留

    def test_other_sheet_or_text_is_not_matched(self):
        new = detected(problem("n1", "k", [0, 1], sheet="Sheet2"), problem("n2", "i", [0, 1]))
        merged = merge(new, [problem("old", "k", [0, 1])], keep_unmatched=False)
        self.assertEqual(set(merged), {"n1", "n2"})

    def test_without_occurrence_only_unique_text_is_matched(self):
        self.assertIn("old", merge(detected(problem("n1", "印。")), [problem("old", "印。")], keep_unmatched=False))
        repeated = detected(problem("n1", "k"), problem("n2", "k"))
        self.assertEqual(set(merge(repeated, [problem("old", "k")], keep_unmatched=False)), {"n1", "n2"})

    def test_first_scan_does_not_mark_new(self):
        merged = merge(detected(problem("n1", "k")), [], mark_new=False)
        self.assertFalse(merged["n1"]["is_new"])


if __name__ == "__main__":
    unittest.main()
