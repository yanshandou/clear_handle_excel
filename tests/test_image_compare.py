"""按图比：红框偏了能看出来，周围格子变了不影响"""
import tempfile
import unittest
from pathlib import Path

from core import image_compare as ic
from core.pdf_detector import fitz


def picture_png(path, red_at, size=(400, 300), picture=(50, 40, 250, 190)):
    """白底区域里一张灰色「截图」，截图里一个红框；返回图片路径"""
    doc = fitz.open()
    page = doc.new_page(width=size[0], height=size[1])
    page.draw_rect(fitz.Rect(*picture), color=None, fill=(0.6, 0.6, 0.7))
    x, y = red_at
    page.draw_rect(fitz.Rect(x, y, x + 30, y + 20), color=(0.9, 0, 0), width=3)
    page.get_pixmap().save(str(path))
    return str(path)


class ImageCompareTest(unittest.TestCase):
    def test_same_box_not_suspect_moved_box_suspect(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            area, box = (0, 0, 400, 300), (50, 40, 200, 150)
            a = ic.thumb_from_png(picture_png(tmp / "a.png", (180, 150)), area, box)
            same = ic.thumb_from_png(picture_png(tmp / "b.png", (180, 150)), area, box)
            moved = ic.thumb_from_png(picture_png(tmp / "c.png", (150, 120)), area, box)
            ok = ic.compare(a, same)
            self.assertFalse(ok["suspect"])
            self.assertLess(ok["red_shift"], 1)
            bad = ic.compare(a, moved)
            self.assertTrue(bad["suspect"])
            self.assertGreater(bad["red_shift"], 10)

    def test_picture_elsewhere_in_bigger_area_still_matches(self):
        # 新版周围格子变高了：截图在区域里的位置不同，但只取截图那块，比出来一样
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            a = ic.thumb_from_png(picture_png(tmp / "a.png", (180, 150)), (0, 0, 400, 300), (50, 40, 200, 150))
            b = ic.thumb_from_png(picture_png(tmp / "b.png", (180, 190), size=(400, 360), picture=(50, 80, 250, 230)),
                                  (0, 0, 400, 360), (50, 80, 200, 150))
            self.assertFalse(ic.compare(a, b)["suspect"])

    def test_missing_side_returns_none(self):
        self.assertIsNone(ic.compare(None, [(0, 0, 0)] * ic.SIZE * ic.SIZE))

    def test_page_thumb(self):
        doc = fitz.open()
        page = doc.new_page()
        page.draw_rect(fitz.Rect(100, 100, 300, 250), color=None, fill=(0.6, 0.6, 0.7))
        page.draw_rect(fitz.Rect(230, 210, 260, 230), color=(0.9, 0, 0), width=3)
        t = ic.thumb_from_page(page, fitz.Rect(100, 100, 300, 250))
        self.assertEqual(len(t), ic.SIZE * ic.SIZE)
        self.assertGreater(ic._red_center(t)[2], ic.RED_MIN)


if __name__ == "__main__":
    unittest.main()
