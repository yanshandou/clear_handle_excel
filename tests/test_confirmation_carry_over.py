import unittest

from core.workflow import MANUAL_CONFIRMED_REASON, _carry_over_confirmations


def problem(text, occurrence=None, sheet="Sheet1", status="pending"):
    item = {"wb_path": "C:/excel/book.xlsx", "sheet_name": sheet, "text": text, "occurrence": occurrence,
            "status": status}
    if status == "accepted":
        item["manual_confirmation"] = {"confirmed_at": "2026-10-21T10:00:00", "previous_status": "pending"}
    return item


class CarryOverConfirmationTests(unittest.TestCase):
    def test_same_text_and_occurrence_keeps_confirmation_regardless_of_position(self):
        new = {"a": problem("お客様　一郎", [0, 3]), "b": problem("お客様　三郎", [0, 1])}
        carried = _carry_over_confirmations(new, [problem("お客様　一郎", [0, 3], status="accepted")])

        self.assertEqual(carried, 1)
        self.assertEqual(new["a"]["status"], "accepted")
        self.assertEqual(new["a"]["review_reason"], MANUAL_CONFIRMED_REASON)
        self.assertEqual(new["a"]["manual_confirmation"]["confirmed_at"], "2026-10-21T10:00:00")
        self.assertEqual(new["b"]["status"], "pending")

    def test_repeated_text_follows_occurrence_order(self):
        # 同一列有 3 个「k」，只确认了第 2 个
        new = {str(i): problem("k", [i, 3]) for i in range(3)}
        _carry_over_confirmations(new, [problem("k", [1, 3], status="accepted")])
        self.assertEqual([new[str(i)]["status"] for i in range(3)], ["pending", "accepted", "pending"])

    def test_changed_count_of_same_text_is_treated_as_new_problem(self):
        new = {"a": problem("k", [1, 4])}
        self.assertEqual(_carry_over_confirmations(new, [problem("k", [1, 3], status="accepted")]), 0)
        self.assertEqual(new["a"]["status"], "pending")

    def test_other_sheet_or_text_is_not_matched(self):
        new = {"a": problem("k", [0, 1], sheet="Sheet2"), "b": problem("i", [0, 1])}
        self.assertEqual(_carry_over_confirmations(new, [problem("k", [0, 1], status="accepted")]), 0)

    def test_without_occurrence_only_unique_text_is_matched(self):
        unique = {"a": problem("印。")}
        self.assertEqual(_carry_over_confirmations(unique, [problem("印。", status="accepted")]), 1)
        repeated = {"a": problem("k"), "b": problem("k")}
        self.assertEqual(_carry_over_confirmations(repeated, [problem("k", status="accepted")]), 0)

    def test_each_confirmation_is_used_once(self):
        new = {"a": problem("印。"), "b": problem("印。", [0, 1])}
        self.assertEqual(_carry_over_confirmations(new, [problem("印。", [0, 1], status="accepted")]), 1)


if __name__ == "__main__":
    unittest.main()
