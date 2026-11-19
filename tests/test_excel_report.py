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


if __name__ == "__main__":
    unittest.main()
