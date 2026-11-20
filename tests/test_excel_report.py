"""Excel 对比报告（core/excel_report.py）：任选两份 Excel 生成单个 HTML 报告"""
import base64
import json
import re
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

from core import excel_diff, excel_report


def make(path: Path, new: bool) -> Path:
    wb = Workbook()
    ws = wb.active
    ws.title = "一覧"
    for r in range(1, 601):
        ws.cell(r, 1, f"項目{r}")
    if new:
        for r in range(1, 601):  # 600 格都改：清单不能截断（对比页最多 400）
            ws.cell(r, 1, f"項目{r}改")
        ws.oddHeader.center.text = "</script><b>見出し</b>"
        dv = DataValidation(type="list", formula1='"A,B"')
        ws.add_data_validation(dv)
        dv.add("B1:B5")
        ws.auto_filter.ref = "A1:A600"
        ws.protection.sheet = True
        wb.defined_names["範囲"] = DefinedName("範囲", attr_text="'一覧'!$A$1:$A$10")
        wb.create_sheet("追加")["A1"] = "x"
    wb.save(path)
    return path


class ExcelReportTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)
        self.a, self.b = make(self.root / "a.xlsx", False), make(self.root / "b.xlsx", True)

    def test_full_lists_everything_and_new_kinds(self):
        result = excel_diff.diff(self.b, self.a, full=True)
        sheet = next(s for s in result["sheets"] if s["name"] == "一覧")
        self.assertEqual(len(sheet["text"]), 600)
        self.assertEqual(sheet["text_more"], 0)
        for kind in ("页眉页脚", "数据验证", "自动筛选", "工作表保护"):
            self.assertIn(kind, sheet["counts"])
        self.assertEqual([r["kind"] for r in result["book"]], ["定义的名称"])
        # 对比页（不是 full）照旧最多列 400 条
        self.assertEqual(len(next(s for s in excel_diff.diff(self.b, self.a)["sheets"] if s["name"] == "一覧")["text"]), 400)

    def test_report_is_single_self_contained_html(self):
        page = excel_report.render_html(excel_diff.diff(self.b, self.a, full=True))
        self.assertNotIn('src="', page)  # 不引外部脚本 / 样式
        self.assertNotIn("<link", page)
        self.assertEqual(page.count("</script>"), 2)  # 数据里的 </script> 不会提前结束
        data = re.search(r'<script id="xr-data" type="application/json">(.*?)</script>', page, re.S)[1]
        self.assertEqual(json.loads(data)["mine"], str(self.b))

    def test_build_from_uploaded_files(self):
        side = lambda p: {"name": "同名.xlsx", "data": base64.b64encode(p.read_bytes()).decode()}  # noqa: E731
        result = excel_report.build_from_request(side(self.a), side(self.b))
        self.assertFalse(result["identical"])
        self.assertEqual(result["sheets"]["only_new"], 1)
        self.assertTrue(Path(result["saved"]).is_file())
        self.assertEqual(excel_report.saved_report(result["filename"]), Path(result["saved"]).resolve())
        with self.assertRaises(ValueError):
            excel_report.saved_report("../../secret.txt")

    def test_same_file_is_identical(self):
        self.assertTrue(excel_report.build(self.a, self.a, save=False)["identical"])


class SheetPairTests(unittest.TestCase):
    """挑两张表对比（用户 2026-11-19）：同一份文件的 Sheet1 vs Sheet2，表名不同也能比"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.path = Path(tmp.name) / "一份.xlsx"
        wb = Workbook()
        first = wb.active
        first.title = "Sheet1"
        second = wb.create_sheet("Sheet2")
        twin = wb.create_sheet("コピー")
        hidden = wb.create_sheet("隠し")
        hidden.sheet_state = "hidden"
        for r in range(1, 6):
            first.cell(r, 1, f"行{r}")
            twin.cell(r, 1, f"行{r}")
            second.cell(r, 1, f"行{r}" if r != 3 else "変更")
        second["B2"] = "追加"
        second.column_dimensions["A"].width = 30
        wb.save(self.path)

    def test_lists_sheets_with_state(self):
        names = excel_diff.sheet_names(self.path)
        self.assertEqual([n["name"] for n in names], ["Sheet1", "Sheet2", "コピー", "隠し"])
        self.assertEqual(names[3]["state"], "隐藏")
        side = {"name": "x.xlsx", "data": base64.b64encode(self.path.read_bytes()).decode()}
        self.assertEqual(len(excel_report.sheets_from_request(side)["sheets"]), 4)

    def test_same_file_two_sheets(self):
        result = excel_diff.diff_sheets(self.path, "Sheet1", self.path, "Sheet2")
        self.assertFalse(result["identical"])
        self.assertEqual(result["pair"], {"a": "Sheet1", "b": "Sheet2"})
        sheet = result["sheets"][0]
        self.assertEqual(sheet["name"], "Sheet1 ↔ Sheet2")
        self.assertEqual({(r["cell"], r["kind"]) for r in sheet["text"]}, {("A3", "mod"), ("B2", "add")})
        self.assertIn("列宽", sheet["counts"])
        self.assertEqual(result["book"], [])
        # 效果图左右各画自己那张表
        cells_old = {(c[0], c[1]): c[2] for c in sheet["view"]["old"]["cells"]}
        cells_new = {(c[0], c[1]): c[2] for c in sheet["view"]["new"]["cells"]}
        self.assertEqual(cells_old[(3, 1)], "行3")
        self.assertEqual(cells_new[(3, 1)], "変更")

    def test_identical_sheets_with_different_names(self):
        self.assertTrue(excel_diff.diff_sheets(self.path, "Sheet1", self.path, "コピー")["identical"])

    def test_report_and_errors(self):
        result = excel_report.build(self.path, self.path, save=False, sheet_a="Sheet1", sheet_b="Sheet2")
        self.assertEqual(result["pair"], {"a": "Sheet1", "b": "Sheet2"})
        self.assertEqual(result["filename"], "Excel对比报告_一份_Sheet1_vs_Sheet2.html")
        self.assertIn("「Sheet1」 vs", result["html"])
        with self.assertRaises(ValueError):
            excel_report.build(self.path, self.path, save=False, sheet_a="Sheet1", sheet_b="ない")
        with self.assertRaises(ValueError):
            excel_report.build(self.path, self.path, save=False, sheet_a="Sheet1")
        req = excel_report.build_from_request({"path": str(self.path)}, {"path": f'"{self.path}"'}, "Sheet1", "Sheet2")
        self.assertFalse(req["identical"])


if __name__ == "__main__":
    unittest.main()
