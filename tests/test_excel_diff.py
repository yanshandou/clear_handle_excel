"""仿 Git 的 Excel 对比（用户 2026-11-11）：左 = 远程，右 = 要上传的"""
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

from core import excel_diff


def book(path: Path, new: bool) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "画面仕様"
    ws["A1"], ws["B1"] = "画面ID", "KS0101"
    ws["B2"] = "居宅支援業務一覧" if new else "居宅支援一覧"
    ws["A5"], ws["B5"], ws["C5"] = 10, 20, "=SUM(A5:B5)" if new else "=A5+B5"
    ws["A1"].font = Font(name="ＭＳ Ｐゴシック" if new else "游ゴシック", size=11, bold=True)
    ws["A1"].fill = PatternFill("solid", fgColor="FFFFFF00" if new else "FFD9E1F2")
    ws["B1"].border = Border(bottom=Side(style="medium" if new else "thin"))
    if new:
        ws["B2"].alignment = Alignment(wrap_text=True)
        ws.row_dimensions[2].height = 36
        ws.column_dimensions["B"].width = 40
        ws.print_area = "A1:C10"
        ws["A7"] = "備考"
        from openpyxl.comments import Comment
        ws["A1"].comment = Comment("確認済み", "XU")
        ws["B1"].hyperlink = "https://example.com/spec"
        ws.freeze_panes = "A2"
        wb.create_sheet("補足")["A1"] = "x"
    else:
        ws.print_area = "A1:C12"
        ws.merge_cells("A10:C10")
        ws["A8"] = "削除予定"
        wb.create_sheet("変更履歴")["A1"] = "初版"
        wb.create_sheet("帳票定義")["A1"] = "x"
    wb.create_sheet("項目一覧")["A1"] = "No"
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)
    return path


class ExcelDiffTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def test_lists_every_kind_of_difference(self):
        mine, theirs = book(self.root / "mine.xlsx", True), book(self.root / "theirs.xlsx", False)
        result = excel_diff.diff(mine, theirs)
        self.assertFalse(result["identical"])
        sheets = {s["name"]: s for s in result["sheets"]}
        main = sheets["画面仕様"]
        self.assertEqual(main["status"], "changed")
        self.assertEqual({(t["cell"], t["kind"]) for t in main["text"]}, {("B2", "mod"), ("A7", "add"), ("A8", "del")})
        self.assertEqual(main["formulas"], [{"cell": "C5", "old": "=A5+B5", "new": "=SUM(A5:B5)", "kind": "mod"}])
        styles = {(g["kind"], g["old"], g["new"]) for g in main["styles"]}
        self.assertIn(("字体", "游ゴシック 11 粗体", "ＭＳ Ｐゴシック 11 粗体"), styles)
        self.assertIn(("填充色", "#D9E1F2", "#FFFF00"), styles)
        self.assertIn(("边框", "下细", "下中"), styles)
        self.assertIn(("对齐", "", "自动换行"), styles)
        whole = {(w["kind"], w["what"], w["old"], w["new"]) for w in main["whole"]}
        self.assertIn(("行高", "第 2 行", "默认", "36"), whole)
        self.assertIn(("列宽", "B 列", "默认", "40"), whole)
        self.assertIn(("打印设置", "打印区域", "A1:C12", "A1:C10"), whole)
        self.assertIn(("合并单元格", "合并的范围", "A10:C10", ""), whole)
        self.assertIn(("冻结窗格", "冻结到", "", "A2"), whole)
        notes = {(n["kind"], n["cell"], n["old"], n["new"]) for n in main["notes"]}
        self.assertEqual(notes, {("批注", "A1", "", "確認済み"), ("超链接", "B1", "", "https://example.com/spec")})
        self.assertEqual(sheets["項目一覧"]["status"], "same")
        self.assertEqual((sheets["帳票定義"]["status"], sheets["帳票定義"]["doomed"]), ("only_old", ""))
        self.assertEqual(sheets["補足"]["status"], "only_new")
        # 精简会删的表单独标出，不算差异
        self.assertEqual(sheets["変更履歴"]["status"], "only_old")
        self.assertTrue(sheets["変更履歴"]["doomed"])
        self.assertEqual((result["summary"]["只有远程有"], result["summary"]["只有你这份有"], result["summary"]["精简会删不算"]), (1, 1, 1))

    def test_same_content_is_identical(self):
        a, b = book(self.root / "a.xlsx", True), book(self.root / "b.xlsx", True)
        result = excel_diff.diff(a, b)
        self.assertTrue(result["identical"], result["summary"])
        self.assertTrue(all(s["status"] == "same" for s in result["sheets"]))

    def test_only_doomed_sheet_differs_is_identical(self):
        a = book(self.root / "a.xlsx", False)
        wb = Workbook()
        b = book(self.root / "b.xlsx", False)
        from openpyxl import load_workbook
        wb = load_workbook(b)
        del wb["変更履歴"]
        wb.save(b)
        self.assertTrue(excel_diff.diff(b, a)["identical"])

    def test_unsupported_or_missing_file(self):
        xls = self.root / "a.xls"
        xls.write_text("x")
        with self.assertRaisesRegex(ValueError, "只支持"):
            excel_diff.diff(xls, xls)
        with self.assertRaisesRegex(ValueError, "找不到"):
            excel_diff.diff(self.root / "none.xlsx", self.root / "none.xlsx")


if __name__ == "__main__":
    unittest.main()


def test_view_only_draws_changed_rows_with_context(tmp_path):
    """效果图：100 多行里只有一行不一样，只画那一行和上下几行；一样的表不画"""
    from openpyxl import Workbook
    books = []
    for text in ("旧", "新"):
        book = Workbook()
        sheet = book.active
        sheet.title = "本体"
        for r in range(1, 201):
            sheet.cell(r, 1, f"行{r}")
        sheet["B100"] = text
        same = book.create_sheet("一样的表")
        same["A1"] = "同じ"
        path = tmp_path / f"{text}.xlsx"
        book.save(path)
        books.append(path)
    result = excel_diff.diff(books[1], books[0])
    sheets = {s["name"]: s for s in result["sheets"]}
    view = sheets["本体"]["view"]
    assert view["rows"] == list(range(100 - excel_diff.VIEW_CONTEXT, 100 + excel_diff.VIEW_CONTEXT + 1))
    assert "文字" in view["marks"]["B100"]
    old_texts = {c[2] for c in view["old"]["cells"]}
    new_texts = {c[2] for c in view["new"]["cells"]}
    assert "旧" in old_texts and "新" in new_texts and "行1" not in old_texts
    assert not sheets["一样的表"].get("view")


def _picture(path: Path, color: tuple[int, int, int]) -> Path:
    from PIL import Image
    Image.new("RGB", (120, 80), color).save(path)
    return path


def _picture_book(folder: Path, name: str, pictures: list[tuple[Path, str]]) -> Path:
    from openpyxl.drawing.image import Image as SheetImage
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "画面"
    for r in range(1, 60):
        sheet.cell(r, 1, f"行{r}")
    for picture, anchor in pictures:
        sheet.add_image(SheetImage(str(picture)), anchor)
    path = folder / name
    workbook.save(path)
    return path


def test_pictures_replaced_moved_added_are_listed(tmp_path):
    """图片换了一张、挪了位置、多了一张都要列出来；不是只比个数（用户 2026-11-12）"""
    red, blue, green = (_picture(tmp_path / f"{n}.png", c) for n, c in
                        (("red", (200, 0, 0)), ("blue", (0, 0, 200)), ("green", (0, 160, 0))))
    old = _picture_book(tmp_path, "old.xlsx", [(red, "C3"), (blue, "C20")])
    same = _picture_book(tmp_path, "same.xlsx", [(red, "C3"), (blue, "C20")])
    new = _picture_book(tmp_path, "new.xlsx", [(green, "C3"), (blue, "C30"), (red, "H50")])
    assert excel_diff.diff(same, old)["identical"]
    sheet = excel_diff.diff(new, old)["sheets"][0]
    rows = [row for row in sheet["whole"] if row["kind"] == "图片图形"]
    assert sheet["counts"]["图片图形"] == 3 and len(rows) == 3
    text = " | ".join(f"{row['what']} {row['old']} → {row['new']}" for row in rows)
    assert "换成了另一张图片" in text and "在 C20" in text and "在 C30" in text
    assert any(not row["old"] and "H50" in row["what"] for row in rows)  # 多出来的那张
    boxes = sheet["view"]["boxes"]
    assert len(boxes["old"]) == 2 and len(boxes["new"]) == 3
    assert 30 in sheet["view"]["rows"] and 50 in sheet["view"]["rows"]


def test_shape_text_and_size_changes_are_listed():
    """图形 / 文本框里的字、大小变了要列出来；只是名字不同不算"""
    def drawing(text: str, cx: int, name: str = "テキスト 1") -> bytes:
        return (f'<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                f'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                f'<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>'
                f'<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>8</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>'
                f'<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="2" name="{name}"/><xdr:cNvSpPr txBox="1"/></xdr:nvSpPr>'
                f'<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="720000"/></a:xfrm><a:prstGeom prst="rect"/></xdr:spPr>'
                f'<xdr:txBody><a:bodyPr/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
                f'</xdr:wsDr>').encode("utf-8")
    parse = lambda xml: excel_diff._parse_drawing(xml, lambda rid: None)  # noqa: E731
    base = parse(drawing("旧い説明", 3600000))
    assert base[0]["kind"] == "文本框" and base[0]["from"] == (5, 2) and base[0]["size"] == (10.0, 2.0)
    assert excel_diff._diff_drawings(base, parse(drawing("旧い説明", 3600000, "テキスト 9"))) == []
    rows = excel_diff._diff_drawings(base, parse(drawing("新しい説明", 1800000)))
    assert len(rows) == 1 and rows[0]["changes"] == ["文字", "大小"]
    assert "旧い説明" in rows[0]["old"] and "新しい説明" in rows[0]["new"] and "宽 5" in rows[0]["new"]
