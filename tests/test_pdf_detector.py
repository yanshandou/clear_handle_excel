import tempfile
import unittest
from pathlib import Path

from core.pdf_detector import (
    RegionTrackingError,
    find_clipped_spans_on_page,
    fitz,
    is_region_still_clipped,
    texts_match,
)


class PdfTrackingTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.pdf_path = Path(self.temp.name) / "test.pdf"

    def add_text(self, page, text, y, clipped=False):
        page.insert_text((60, y), text, fontsize=12)
        span = next(
            span for block in page.get_text("dict")["blocks"]
            for line in block.get("lines", []) for span in line["spans"]
            if span["text"] == text and abs(span["origin"][1] - y) < 1
        )
        bbox = tuple(span["bbox"])
        if clipped:
            line_y = (bbox[1] + bbox[3]) / 2
            page.draw_line((55, line_y), (250, line_y), width=0.5)
        return bbox

    def test_nearby_different_clipped_text_is_not_original_target(self):
        with fitz.open() as doc:
            page = doc.new_page()
            bbox = self.add_text(page, "Original target", 100)
            self.add_text(page, "Different text", 117, clipped=True)
            self.assertEqual(len(find_clipped_spans_on_page(page)), 1)
            doc.save(self.pdf_path)
        clipped, _, actual_bbox, resolved_page = is_region_still_clipped(
            self.pdf_path, 0, "Original target", bbox
        )
        self.assertFalse(clipped)
        self.assertEqual(actual_bbox, bbox)
        self.assertEqual(resolved_page, 0)

    def test_two_common_characters_do_not_match_different_labels(self):
        self.assertFalse(texts_match("日報集計", "月報集計"))
        self.assertFalse(texts_match("Total amount", "Total count"))
        self.assertFalse(texts_match("Amount", "Amount payable"))
        self.assertTrue(texts_match("Total amount", "Total\namount"))

    def test_missing_text_is_not_reported_as_fixed(self):
        with fitz.open() as doc:
            page = doc.new_page()
            self.add_text(page, "Other text", 100)
            doc.save(self.pdf_path)
        with self.assertRaisesRegex(RegionTrackingError, "无法定位"):
            is_region_still_clipped(self.pdf_path, 0, "Missing target", (60, 80, 140, 100))

    def test_repeated_text_is_not_guessed_from_nearest_position(self):
        with fitz.open() as doc:
            page = doc.new_page()
            bbox = self.add_text(page, "Repeated", 100, clipped=True)
            self.add_text(page, "Repeated", 160)
            doc.save(self.pdf_path)
        with self.assertRaisesRegex(RegionTrackingError, "无法唯一定位"):
            is_region_still_clipped(self.pdf_path, 0, "Repeated", bbox)

    def test_target_moved_multiple_pages_is_rechecked_on_new_page(self):
        with fitz.open() as doc:
            doc.new_page()
            doc.new_page()
            page = doc.new_page()
            bbox = self.add_text(page, "Moved target", 200, clipped=True)
            doc.save(self.pdf_path)
        clipped, _, actual_bbox, resolved_page = is_region_still_clipped(
            self.pdf_path, 0, "Moved target", (60, 80, 140, 100)
        )
        self.assertTrue(clipped)
        self.assertEqual(actual_bbox, bbox)
        self.assertEqual(resolved_page, 2)

    def test_correct_target_still_clipped_is_detected(self):
        with fitz.open() as doc:
            page = doc.new_page()
            bbox = self.add_text(page, "Clipped target", 100, clipped=True)
            doc.save(self.pdf_path)
        clipped, _, _, _ = is_region_still_clipped(self.pdf_path, 0, "Clipped target", bbox)
        self.assertTrue(clipped)


if __name__ == "__main__":
    unittest.main()
