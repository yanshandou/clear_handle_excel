import unittest

from core.hidden_text import (
    Piece,
    REASON_CLIPPED,
    REASON_HASH,
    REASON_MISSING,
    REASON_TRUNCATED,
    collect_text_pieces,
    find_hidden_cells,
    locate_text,
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
    def test_halfwidth_voiced_table_header_is_visible(self):
        variants = ("ﾃｰﾌﾞﾙ", "テーブル", "テーフ\u3099ル")
        for source in variants:
            for pdf_text in ("ﾃｰﾌﾞﾙ", "テーブル"):
                with self.subTest(source=source, pdf_text=pdf_text):
                    def build(page):
                        # 两行合并后的表头，文字居中，边框没有切到文字。
                        page.draw_rect(fitz.Rect(50, 80, 150, 120))
                        page.insert_text((70, 105), pdf_text, fontsize=10, fontname="japan")

                    self.assertEqual(detect([(46, 26, source)], build), [])

    def test_combining_marks_and_compatibility_glyphs_match_complete_text(self):
        for text in ("ﾊﾟﾋﾟﾌﾟﾍﾟﾎﾟ", "カ\u3099", "e\u0301", "A\u0301\u0327", "\u1100\u1161", "㍿"):
            with self.subTest(text=text):
                pieces = [Piece(0, [(c, (i * 10, 0, i * 10 + 10, 10), True)
                                    for i, c in enumerate(text)])]
                self.assertEqual(find_hidden_cells(
                    [(1, 1, text)], pieces, lambda r, c: text, lambda r, c: True,
                ), [])

    def test_hidden_dakuten_is_still_reported_with_its_location(self):
        pieces = [Piece(0, [("ﾌ", (50, 90, 60, 100), True),
                            ("ﾞ", (60, 90, 65, 100), False)])]
        result = find_hidden_cells([(46, 26, "ﾌﾞ")], pieces, lambda r, c: "ﾌﾞ", lambda r, c: True)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_CLIPPED)
        self.assertEqual(result[0]["visible"], "ﾌ")
        self.assertEqual(result[0]["bbox"], (50, 90, 65, 100))

    def test_halfwidth_duplicate_headers_still_need_separate_printed_copies(self):
        def build(page):
            page.insert_text((50, 100), "ﾃｰﾌﾞﾙ", fontsize=10, fontname="japan")

        result = detect([(46, 26, "ﾃｰﾌﾞﾙ"), (56, 26, "ﾃｰﾌﾞﾙ")], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_MISSING)

    def test_halfwidth_header_with_missing_suffix_is_still_truncated(self):
        def build(page):
            page.insert_text((50, 100), "ﾃｰﾌﾞ", fontsize=10, fontname="japan")

        result = detect([(46, 26, "ﾃｰﾌﾞﾙ")], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_TRUNCATED)
        self.assertEqual(result[0]["visible"], "ﾃｰﾌﾞ")

    def test_halfwidth_header_can_be_located_in_pdf(self):
        with fitz.open() as doc:
            page = doc.new_page()
            page.insert_text((70, 105), "ﾃｰﾌﾞﾙ", fontsize=10, fontname="japan")
            location = locate_text(collect_text_pieces(doc), "ﾃｰﾌﾞﾙ")
        self.assertIsNotNone(location)
        self.assertEqual(location[0], 0)
        self.assertLess(location[1][1], 105)
        self.assertGreater(location[1][3], 105)

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

    def test_visible_paragraphs_separated_by_empty_lines_are_not_truncated(self):
        for newline in ("\n", "\r\n"):
            for blank_lines in (1, 2):
                with self.subTest(newline=newline, blank_lines=blank_lines):
                    lines = ["Registration check", "chkErrorRyouyouhiMeisai()",
                             "Claim error check", "getKokuhoKindEntry()"]
                    text = newline.join(lines[:2]) + newline * (blank_lines + 1) + newline.join(lines[2:])

                    def build(page):
                        second_y = 112 + 12 * (blank_lines + 2)
                        for y, line in zip((100, 112, second_y, second_y + 12), lines):
                            add_text(page, 50, y, line)

                    self.assertEqual(detect([(47, 33, text)], build), [])

    def test_japanese_check_descriptions_with_blank_line_are_not_truncated(self):
        lines = ["●登録前チェック", "BIIryouChk::chkErrorRyouyouhiMeisai()",
                 "●請求エラーチェック", "BIIryouKokuhoError::getKokuhoKindEntry()"]
        text = "\n".join(lines[:2]) + "\n\n" + "\n".join(lines[2:])

        def build(page):
            for y, line in zip((100, 112, 148, 160), lines):
                page.insert_text((50, y), line, fontsize=10, fontname="japan")

        self.assertEqual(detect([(47, 33, text)], build), [])

    def test_missing_last_line_after_blank_line_is_still_truncated(self):
        text = "Registration\ncheckMethod()\n\nClaim\nmissingMethod()"

        def build(page):
            for y, line in ((100, "Registration"), (112, "checkMethod()"), (148, "Claim")):
                add_text(page, 50, y, line)

        result = detect([(47, 33, text)], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_TRUNCATED)
        self.assertEqual(result[0]["visible"], "RegistrationcheckMethod()Claim")

    def test_large_gap_without_source_blank_line_is_not_joined(self):
        def build(page):
            add_text(page, 50, 100, "First")
            add_text(page, 50, 148, "Second")

        result = detect([(1, 1, "First\nSecond")], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_TRUNCATED)

    def test_source_blank_line_does_not_join_another_column_or_distant_row(self):
        for x, y in ((300, 136), (50, 300)):
            with self.subTest(x=x, y=y):
                def build(page):
                    add_text(page, 50, 100, "First")
                    add_text(page, x, y, "Second")

                result = detect([(1, 1, "First\n\nSecond")], build)
                self.assertEqual(len(result), 1)
                self.assertEqual(result[0]["reason"], REASON_TRUNCATED)

    def test_blank_line_paragraphs_still_check_character_clipping(self):
        def build(page):
            add_text(page, 50, 100, "First")
            add_text(page, 52, 136, "LongSecondParagraph", clip=(50, 126, 100, 140))

        result = detect([(1, 1, "First\n\nLongSecondParagraph")], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_CLIPPED)
        self.assertLess(len(result[0]["visible"]), len("FirstLongSecondParagraph"))

    def test_location_includes_visible_paragraph_after_blank_line(self):
        with fitz.open() as doc:
            page = doc.new_page()
            add_text(page, 50, 100, "First")
            add_text(page, 50, 136, "Second")
            location = locate_text(collect_text_pieces(doc), "First\n\nSecond")
        self.assertIsNotNone(location)
        self.assertGreater(location[1][3], 136)

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

    def test_spaced_paragraphs_cannot_be_reused_for_two_cells(self):
        def build(page):
            add_text(page, 50, 100, "First")
            add_text(page, 50, 136, "Second")

        text = "First\n\nSecond"
        result = detect([(1, 1, text), (2, 1, text)], build)
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["reason"], REASON_MISSING)

    def test_short_text_does_not_steal_longer_cell(self):
        def build(page):
            add_text(page, 50, 100, "No")
            add_text(page, 50, 120, "hihokensha_no")
        cells = [(1, 1, "No"), (2, 1, "hihokensha_no")]
        self.assertEqual(detect(cells, build), [])

    def test_single_digit_lines_of_truncated_cell_stay_with_it(self):
        # 多行格子里单独一行的「4」「3」，不能被表里更靠下的「4」「3」格子先占走，
        # 否则只认出第一行，「还能看到什么」和「还缺几行」都会算错
        lines = ["桁数④が1以上の場合、", "4", "桁数③が1以上の場合、", "3", "桁数②が1以上の場合、"]
        text = "桁数④が1以上の場合、4\n桁数③が1以上の場合、3\n桁数②が1以上の場合、2\n上記以外の場合、0"

        def build(page):
            for i, line in enumerate(lines):
                page.insert_text((200, 100 + 12 * i), line, fontsize=10, fontname="japan")
            page.insert_text((50, 400), "4", fontsize=10, fontname="japan")
            page.insert_text((50, 412), "3", fontsize=10, fontname="japan")

        cells = [(104, 10, text), (200, 1, 4.0), (201, 1, 3.0)]
        result = detect(cells, build, {(200, 1): "4", (201, 1): "3"})
        self.assertEqual(len(result), 1)
        self.assertEqual((result[0]["row"], result[0]["reason"]), (104, REASON_TRUNCATED))
        self.assertEqual(result[0]["visible"], "".join(lines))
        self.assertEqual(result[0]["missing_lines"], 2)  # 「2」和「上記以外の場合、0」
        self.assertAlmostEqual(result[0]["line_pitch"], 12.0, places=1)

    def test_only_copy_of_short_text_is_not_taken_away(self):
        def build(page):
            page.insert_text((200, 100), "桁数④が1以上の場合、", fontsize=10, fontname="japan")
            page.insert_text((200, 112), "4", fontsize=10, fontname="japan")

        cells = [(104, 10, "桁数④が1以上の場合、4\n上記以外の場合、0"), (200, 1, 4.0)]
        result = detect(cells, build, {(200, 1): "4"})
        self.assertEqual([r["row"] for r in result], [104])
        self.assertEqual(result[0]["visible"], "桁数④が1以上の場合、")

    def test_formatted_text_uses_displayed_value(self):
        def build(page):
            add_text(page, 50, 100, "No.15")
        self.assertEqual(detect([(1, 1, "15x")], build, {(1, 1): "No.15"}), [])

    def test_display_text_prefix_does_not_shorten_expected_cell_content(self):
        def build(page):
            page.insert_text((50, 100), "文字列", fontsize=10, fontname="japan")

        result = detect([(67, 18, "文字列(100)")], build, {(67, 18): "文字列"})
        self.assertEqual(len(result), 1)
        self.assertEqual(result[0]["text"], "文字列(100)")
        self.assertEqual(result[0]["reason"], REASON_TRUNCATED)


if __name__ == "__main__":
    unittest.main()
