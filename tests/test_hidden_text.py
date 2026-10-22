import unittest

from core.hidden_text import (
    REASON_CLIPPED,
    REASON_HASH,
    REASON_MISSING,
    REASON_TRUNCATED,
    collect_text_pieces,
    find_hidden_cells,
)
from core.pdf_detector import fitz


def add_text(page, x, y, text, clip=None):
    """写一段文字；clip=(x0, y0, x1, y1) 时像 Excel 一样用裁剪区域只显示格子内的部分"""
    if clip is None:
        page.insert_text((x, y), text, fontsize=10, fontname="helv")
        return
    page.insert_text((0, 0), " ", fontsize=1, fontname="helv")  # 确保字体资源存在
    x0, y0, x1, y1 = clip
    height = page.rect.height
    stream = (
        f"q {x0} {height - y1} {x1 - x0} {y1 - y0} re W n "
        f"BT /helv 10 Tf {x} {height - y} Td ({text}) Tj ET Q"
    ).encode()
    xref = page.get_contents()[-1]
    page.parent.update_stream(xref, page.parent.xref_stream(xref) + b"\n" + stream)


def detect(cells, build, displayed=None, printed=lambda r, c: True):
    with fitz.open() as doc:
        page = doc.new_page()
        build(page)
        data = doc.tobytes()
    with fitz.open("pdf", data) as doc:
        pieces = collect_text_pieces(doc)
    displayed = displayed or {}
    return find_hidden_cells(
        cells, pieces,
        lambda r, c: displayed.get((r, c), next(str(v) for rr, cc, v in cells if (rr, cc) == (r, c))),
        printed,
    )


class HiddenTextTests(unittest.TestCase):
    def test_fully_printed_cells_are_not_reported(self):
        def build(page):
            add_text(page, 50, 100, "hihokensha_no")
            add_text(page, 125, 100, "DbExKhplg00")  # 自动换行成两行
            add_text(page, 125, 112, "00Mapper")
            add_text(page, 185, 100, "hihokensha_no")
            add_text(page, 50, 130, "12")
        cells = [(1, 1, "hihokensha_no"), (1, 2, "DbExKhplg0000Mapper"), (1, 3, "hihokensha_no"), (2, 1, 12.0)]
        self.assertEqual(detect(cells, build, {(2, 1): "12"}), [])

    def test_wrapped_line_hidden_by_row_height_is_reported(self):
        def build(page):
            add_text(page, 50, 100, "hihokensha_no")
            add_text(page, 125, 100, "DbExKhplg0")  # 第二行没画出来
            add_text(page, 185, 100, "hihokensha_no")
        cells = [(23, 1, "hihokensha_no"), (23, 2, "DbExKhplg0000Mapper"), (23, 3, "hihokensha_no")]
        result = detect(cells, build)
        self.assertEqual(len(result), 1)
        self.assertEqual((result[0]["row"], result[0]["col"]), (23, 2))
        self.assertEqual(result[0]["reason"], REASON_TRUNCATED)
        self.assertEqual(result[0]["visible"], "DbExKhplg0")
        self.assertEqual(result[0]["page"], 0)

    def test_text_cut_off_by_clip_region_is_reported(self):
        def build(page):
            add_text(page, 52, 100, "VeryLongTextBehindNeighbour", clip=(50, 90, 100, 104))
            add_text(page, 104, 100, "Next")
        cells = [(1, 1, "VeryLongTextBehindNeighbour"), (1, 2, "Next")]
        result = detect(cells, build)
        self.assertEqual([(r["col"], r["reason"]) for r in result], [(1, REASON_CLIPPED)])
        self.assertTrue("VeryLongTextBehindNeighbour".startswith(result[0]["visible"]))
        self.assertLess(len(result[0]["visible"]), len("VeryLongTextBehindNeighbour"))

    def test_hash_number_and_missing_text(self):
        def build(page):
            add_text(page, 50, 100, "###")
        cells = [(1, 1, 123456.0), (2, 1, "Gone"), (3, 1, "Hidden row")]
        result = detect(
            cells, build, {(1, 1): "###"}, printed=lambda r, c: r != 3,
        )
        self.assertEqual([(r["row"], r["reason"]) for r in result], [(1, REASON_HASH), (2, REASON_MISSING)])
        self.assertEqual(result[0]["text"], "123456")
        self.assertIsNone(result[1]["bbox"])

    def test_duplicate_texts_need_one_printed_copy_each(self):
        def build(page):
            add_text(page, 50, 100, "Same")
        cells = [(1, 1, "Same"), (2, 1, "Same")]
        result = detect(cells, build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_MISSING)

    def test_short_text_does_not_steal_longer_cell(self):
        def build(page):
            add_text(page, 50, 100, "No")
            add_text(page, 50, 120, "hihokensha_no")
        cells = [(1, 1, "No"), (2, 1, "hihokensha_no")]
        self.assertEqual(detect(cells, build), [])

    def test_formatted_text_uses_displayed_value(self):
        def build(page):
            add_text(page, 50, 100, "No.15")
        self.assertEqual(detect([(1, 1, "15x")], build, {(1, 1): "No.15"}), [])


if __name__ == "__main__":
    unittest.main()
