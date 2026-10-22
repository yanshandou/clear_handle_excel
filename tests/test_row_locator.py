import unittest

from core.hidden_text import collect_text_pieces
from core.pdf_detector import fitz
from core.row_locator import (
    build_anchors,
    candidate_cells,
    colored_span_cells,
    excel_color,
    locate_cell,
    tag_color,
    tag_index,
)


class RowLocatorTests(unittest.TestCase):
    def build(self, heights, rows, title_rows=0):
        """按行高把每行文字画到 PDF 上；返回 (pieces, 每行文字的 bbox)"""
        boxes = {}
        with fitz.open() as doc:
            page = doc.new_page()
            y = 40.0
            for row in range(1, title_rows + 1):  # 打印标题：每页重复的表头
                page.insert_text((50, y + heights[row] - 3), rows[row][0], fontsize=9)
                y += heights[row]
            for row, texts in rows.items():
                if row <= title_rows:
                    continue
                for i, text in enumerate(texts):
                    if text:
                        page.insert_text((50 + 120 * i, y + heights[row] - 3), text, fontsize=9)
                        boxes[(row, i)] = (50 + 120 * i, y + heights[row] - 12, 100 + 120 * i, y + heights[row] - 1)
                y += heights[row]
            data = doc.tobytes()
        with fitz.open("pdf", data) as doc:
            return collect_text_pieces(doc), boxes

    def test_repeated_text_is_located_by_nearby_unique_texts(self):
        # 第 2 列每行都是「非表示」，行高忽高忽低，按比例估算会差好几行
        heights = {r: (40.0 if r % 3 == 0 else 13.5) for r in range(1, 31)}
        rows = {r: [f"field_{r}", "非表示"] for r in range(1, 31)}
        pieces, boxes = self.build(heights, rows)
        cells = [(r, c + 1, v) for r, texts in rows.items() for c, v in enumerate(texts)]
        anchors = build_anchors(cells, pieces)
        for target in (5, 23, 29):
            self.assertEqual(
                locate_cell(cells, anchors, "非表示", 0, boxes[(target, 1)], heights.get),
                (target, 2),
            )

    def test_row_without_matching_text_uses_row_heights_between_anchors(self):
        heights = {1: 13.5, 2: 60.0, 3: 13.5, 4: 13.5, 5: 13.5}
        rows = {1: ["first", ""], 2: ["", ""], 3: ["", ""], 4: ["", ""], 5: ["last", ""]}
        pieces, boxes = self.build(heights, rows)
        cells = [(1, 1, "first"), (5, 1, "last")]
        anchors = build_anchors(cells, pieces)
        # PDF 上的文字跟表里对不上（比如显示格式不同），只能按位置换算：落在第 3 行
        row3_box = (50, 40 + 13.5 + 60 + 2, 100, 40 + 13.5 + 60 + 12)
        self.assertEqual(locate_cell(cells, anchors, "???", 0, row3_box, heights.get), (3, 1))

    def test_same_text_on_one_row_is_told_apart_by_column(self):
        heights = {r: 13.5 for r in range(1, 6)}
        rows = {r: [f"id_{r}", f"name_{r}", "-"] for r in range(1, 6)}
        rows[3] = ["id_3", "-", "-"]
        pieces, boxes = self.build(heights, rows)
        cells = [(r, c + 1, v) for r, texts in rows.items() for c, v in enumerate(texts)]
        anchors = build_anchors(cells, pieces)
        self.assertEqual(locate_cell(cells, anchors, "-", 0, boxes[(3, 1)], heights.get), (3, 2))
        self.assertEqual(locate_cell(cells, anchors, "-", 0, boxes[(3, 2)], heights.get), (3, 3))

    def test_page_without_anchors_falls_back(self):
        self.assertIsNone(locate_cell([(1, 1, "x")], {}, "x", 0, (0, 0, 10, 10), lambda r: 15.0))


class ColorTagTests(unittest.TestCase):
    def test_palette_round_trips_through_pdf_colors(self):
        for index in (1, 2, 31, 32, 1000, 32767):
            rgb = tag_color(index)
            self.assertEqual(tag_index(rgb), index)
            # PDF 里颜色是 0~1 的小数，读回来会有舍入
            channels = [round(round((rgb >> shift & 0xFF) / 255, 3) * 255) for shift in (16, 8, 0)]
            self.assertEqual(tag_index(channels[0] << 16 | channels[1] << 8 | channels[2]), index)
        self.assertIsNone(tag_index(0x000000))  # 普通黑字
        self.assertIsNone(tag_index(0xFF0000))
        self.assertEqual(excel_color(0x112233), 0x332211)

    def test_colored_span_identifies_cell_among_same_texts(self):
        tagged = [(3, 2), (7, 2), (9, 4)]
        with fitz.open() as doc:
            page = doc.new_page()
            for i, (y, _) in enumerate(zip((100, 140, 180), tagged), start=1):
                rgb = tag_color(i)
                page.insert_text((60, y), "非表示", fontsize=10, fontname="japan",
                                 color=tuple((rgb >> s & 0xFF) / 255 for s in (16, 8, 0)))
            data = doc.tobytes()
        with fitz.open("pdf", data) as doc:
            spans = [
                span for block in doc[0].get_text("dict")["blocks"]
                for line in block.get("lines", []) for span in line["spans"]
            ]
            targets = [(0, tuple(span["bbox"]), "非表示") for span in spans]
            self.assertEqual(colored_span_cells(doc, targets, tagged), tagged)
            # 坐标对不上（版面变了）时认不出，不乱猜
            self.assertEqual(colored_span_cells(doc, [(0, (0, 0, 10, 10), "非表示")], tagged), [None])

    def test_candidates_fall_back_to_rows_near_estimate(self):
        cells = [(1, 1, "abc"), (5, 1, "xyz"), (6, 2, 12.0), (20, 1, "far")]
        self.assertEqual(candidate_cells(cells, "ab", 10), [(1, 1)])
        self.assertEqual(candidate_cells(cells, "???", 6), [(5, 1), (6, 2)])


if __name__ == "__main__":
    unittest.main()
