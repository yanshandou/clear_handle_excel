import base64
import copy
import importlib
import io
import re
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch

from core import fix_images, report_export, reporter, storage
from core.pdf_detector import fitz


def png_data_url(color) -> str:
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 4, 4), False)
    pix.set_rect(pix.irect, color)
    return "data:image/png;base64," + base64.b64encode(pix.tobytes("png")).decode("ascii")


class ReportImagesTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.folder = self.root / "excel"
        self.folder.mkdir()
        for target, name, value in [(storage, "DATABASE_PATH", self.root / "t.sqlite3"),
                                    (reporter, "RECORD_PATH", self.root / "record.md"),
                                    (fix_images, "IMAGES_DIR", self.root / "_report_images")]:
            p = patch.object(target, name, value)
            p.start()
            self.addCleanup(p.stop)
        saved = copy.deepcopy(reporter.EXECUTION_STATE)
        self.addCleanup(lambda: (reporter.EXECUTION_STATE.clear(), reporter.EXECUTION_STATE.update(saved)))
        reporter.EXECUTION_STATE.update(folder="", files={}, steps=[], batch=[], started_at="", status="idle")
        from core import workflow
        self.workflow = importlib.reload(workflow)
        self.book = self.folder / "a.xlsx"

    def item(self, iid, row, **extra):
        return {"id": iid, "seq": 1, "kind": "clip", "wb_path": str(self.book), "wb_name": "a.xlsx",
                "sheet_name": "S1", "excel_row": row, "cell_addr": f"C{row}", "status": "pending", **extra}

    def record_fixes(self, items):
        reporter.init_execution_batch("切字", self.folder, [self.book])
        self.workflow.STATE.update(folder=str(self.folder), items={it["id"]: it for it in items})
        reporter.update_file_progress("a.xlsx", status="done", **self.workflow._problem_counts(str(self.book)))
        reporter.finish_execution_batch()

    def test_before_image_is_kept_when_screenshot_is_replaced(self):
        item = self.item("i1", 5, image=png_data_url((255, 0, 0)))
        fix_images.remember_before(item)
        item["image"] = png_data_url((0, 255, 0))
        fix_images.remember_before(item)  # 已经存过：不能被后来的截图覆盖
        before = fix_images.image_path(item["fix_images"]["before"]).read_bytes()
        self.assertEqual(before, base64.b64decode(png_data_url((255, 0, 0)).split(",", 1)[1]))

    def test_download_zip_has_html_markdown_and_shared_row_pictures_once(self):
        items = [self.item("i1", 5, status="fixed"), self.item("i2", 5, status="fixed"),
                 self.item("i3", 9, status="fixed")]
        for it in items[:2]:  # 同一行两处：共用一组截图
            fix_images.save(it, "before", png_data_url((255, 0, 0)), key="row5")
            fix_images.save(it, "after", png_data_url((0, 255, 0)), key="row5")
        self.record_fixes(items)  # 第 9 行没有截图

        plain = reporter.render_execution_markdown()
        self.assertNotIn("![", plain)  # 自动保存的 Markdown 不带图

        archive = zipfile.ZipFile(io.BytesIO(report_export.build_report_zip()))
        names = sorted(archive.namelist())
        self.assertEqual(len([n for n in names if n.startswith("images/")]), 2)
        self.assertIn("execution_report.html", names)
        md = archive.read("execution_report.md").decode("utf-8")
        self.assertEqual(md.count("![修复前](images/"), 1)
        self.assertEqual(md.count("![修复后](images/"), 1)
        self.assertIn("**#1、#2　S1 C5（第 5 行）**", md)
        for link in re.findall(r"\]\((images/[^)]+)\)", md):
            self.assertIn(link, names)

        page = archive.read("execution_report.html").decode("utf-8")
        self.assertEqual(page.count('src="data:image/png;base64,'), 2)
        self.assertIn("修复前", page)
        self.assertIn("修复后", page)
        for word in ("人工", "自动", "跳过", "复检"):
            self.assertNotIn(word, page)
        self.assertIsNone(re.search(r"\d{1,2}:\d{2}|\d{4}-\d{2}-\d{2}", page))

    def test_download_includes_sheet_renames_with_html_escaping(self):
        reporter.init_execution_batch("精简 Sheet", self.folder, [self.book])
        reporter.update_file_progress("a.xlsx", sheet_change={"at": "2026-10-28T12:00:00+09:00",
                                     "renamed": {"新<Layout>": "<Layout>"}, "deleted": []})
        archive = zipfile.ZipFile(io.BytesIO(report_export.build_report_zip()))
        page = archive.read("execution_report.html").decode("utf-8")
        self.assertIn("新&lt;Layout&gt; → &lt;Layout&gt;", page)
        self.assertNotIn("新<Layout>", page)
        self.assertIn("新<Layout> → <Layout>", archive.read("execution_report.md").decode("utf-8"))
        self.assertNotIn("2026-10-28", page)

    def test_manually_fixed_problem_gets_after_picture_from_new_pdf(self):
        pdf_path = self.folder / "a_S1.pdf"
        with fitz.open() as doc:
            doc.new_page().insert_text((60, 100), "Target", fontsize=12)
            doc.save(pdf_path)
        old = self.item("i1", 5, text="Target", pdf_path=str(pdf_path), page_index=0,
                        bbox=[60, 90, 100, 102], occurrence=None, image=png_data_url((255, 0, 0)))
        merged = self.workflow._merge_redetected({}, [old], mark_new=True, keep_unmatched=True)
        record = merged["i1"]
        self.assertEqual(record["status"], "fixed")
        self.assertIsNotNone(fix_images.image_path(record["fix_images"]["before"]))
        self.assertIsNotNone(fix_images.image_path(record["fix_images"]["after"]))

    def test_other_folder_clears_old_pictures(self):
        item = self.item("i1", 5, image=png_data_url((255, 0, 0)))
        fix_images.remember_before(item)
        reporter.init_execution_batch("切字", self.folder, [self.book])
        self.assertIsNotNone(fix_images.image_path(item["fix_images"]["before"]))
        other = self.root / "other"
        other.mkdir()
        reporter.init_execution_batch("切字", other, [other / "b.xlsx"])
        self.assertIsNone(fix_images.image_path(item["fix_images"]["before"]))


if __name__ == "__main__":
    unittest.main()
