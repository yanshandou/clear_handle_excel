"""没有 hash 记录 / 记录太旧时，上传前按全部文字 + 打印区域格子数核对（用户 2026-11-11）"""
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from core import content_check, distribute, storage, upload_prep, work_task, workflow as wf


def make(path: Path, cells: dict, area: str | None = "A1:C10", sheets: tuple = ("仕様",), image_note: str = "") -> Path:
    """cells: {"A1": "文字"}；image_note 只改格式（模拟修图片），不改文字"""
    path.parent.mkdir(parents=True, exist_ok=True)
    book = Workbook()
    book.remove(book.active)
    for name in sheets:
        sheet = book.create_sheet(name)
        for addr, value in cells.items():
            sheet[addr] = value
        if area:
            sheet.print_area = area
        if image_note:
            sheet.column_dimensions["A"].width = 30
            from openpyxl.styles import Font
            sheet["A1"].font = Font(bold=True)
    book.save(path)
    return path


def edit(path: Path, **cells) -> None:
    book = load_workbook(path)
    for addr, value in cells.items():
        book.worksheets[0][addr] = value
    book.save(path)


class CompareTests(unittest.TestCase):
    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.root = Path(tmp.name)

    def test_same_text_and_area_even_if_format_differs(self):
        a = make(self.root / "a.xlsx", {"A1": "画面", "B2": 12})
        b = make(self.root / "b.xlsx", {"A1": "画面", "B2": 12.0}, image_note="改过格式")
        self.assertNotEqual(a.read_bytes(), b.read_bytes())
        result = content_check.compare(a, b)
        self.assertTrue(result["checked"])
        self.assertTrue(result["same"], result["diffs"])

    def test_text_difference(self):
        a = make(self.root / "a.xlsx", {"A1": "画面", "B2": "旧"})
        b = make(self.root / "b.xlsx", {"A1": "画面", "B2": "新", "C3": "追加"})
        result = content_check.compare(a, b, "本地", "共享")
        self.assertFalse(result["same"])
        self.assertIn("文字有 2 处不一样", result["diffs"][0])
        self.assertIn("B2 本地「旧」共享「新」", result["diffs"][0])

    def test_print_area_cell_count(self):
        a = make(self.root / "a.xlsx", {"A1": "画面"}, area="A1:C10")
        b = make(self.root / "b.xlsx", {"A1": "画面"}, area="A1:C12")
        result = content_check.compare(a, b)
        self.assertFalse(result["same"])
        self.assertIn("打印区域格子数不一样：本地 30（A1:C10），远程 36（A1:C12）", result["diffs"][0])

    def test_no_print_area_uses_content_range(self):
        a = make(self.root / "a.xlsx", {"A1": "x", "B3": "y"}, area=None)
        b = make(self.root / "b.xlsx", {"A1": "x", "B3": "y"}, area=None)
        self.assertTrue(content_check.compare(a, b)["same"])
        snap = content_check.snapshot(a)["仕様"]
        self.assertEqual(snap["area_cells"], 6)

    def test_sheet_only_on_one_side(self):
        a = make(self.root / "a.xlsx", {"A1": "x"}, sheets=("表1",))
        b = make(self.root / "b.xlsx", {"A1": "x"}, sheets=("表1", "表2"))
        result = content_check.compare(a, b, "本地", "共享")
        self.assertEqual(result["diffs"], ["表「表2」只有共享有"])

    def test_multi_range_print_area(self):
        self.assertEqual(content_check._area_cells("'仕様 書'!$A$1:$B$2,'仕様 書'!$D$1:$D$3"), 7)
        self.assertEqual(content_check._area_cells("Sheet1!$1:$2"), 2 * content_check.MAX_COL)

    def test_unreadable_formats(self):
        fake = self.root / "a.xls"
        fake.write_text("x")
        result = content_check.compare(fake, fake)
        self.assertFalse(result["checked"])
        bad = self.root / "b.xlsx"
        bad.write_text("not a zip")
        self.assertFalse(content_check.compare(bad, bad)["checked"])


def plan(folders):
    return json.dumps({"label": "测试", "people": [{"name": "呉", "folders": folders}]}, ensure_ascii=False)


class JsonCopyTests(unittest.TestCase):
    """按 JSON 复制：没有拿取记录时按内容决定能不能覆盖"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        self.src, self.dst = root / "src", root / "dst"
        for obj, name, value in ((storage, "DATABASE_PATH", root / "t.sqlite3"), (distribute, "_STATUS", {"running": False})):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        wf.STATE["busy"] = False
        self.folders = [{"source": ["A"], "destination": ["S"], "files": 1}]

    def run_copy(self):
        distribute.start()
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def test_no_record_same_content_overwrites(self):
        make(self.src / "A" / "a.xlsx", {"A1": "画面"}, image_note="修了图片")
        make(self.dst / "S" / "a.xlsx", {"A1": "画面"})
        prev = distribute.preview(plan(self.folders), str(self.src), str(self.dst))
        f = prev["folders"][0]["files"][0]
        self.assertEqual(f["guard"], "content_same")
        self.assertEqual(f["conflict"], "")
        result = self.run_copy()
        self.assertEqual(result["folders"][0]["files"][0]["status"], "ok")
        self.assertEqual((self.dst / "S" / "a.xlsx").read_bytes(), (self.src / "A" / "a.xlsx").read_bytes())

    def test_no_record_different_content_blocks(self):
        make(self.src / "A" / "a.xlsx", {"A1": "画面"})
        make(self.dst / "S" / "a.xlsx", {"A1": "画面", "B5": "别人加的"})
        before = (self.dst / "S" / "a.xlsx").read_bytes()
        prev = distribute.preview(plan(self.folders), str(self.src), str(self.dst))
        f = prev["folders"][0]["files"][0]
        self.assertEqual(f["guard"], "content_diff")
        self.assertIn("拿最新版", f["conflict"])
        self.assertIn("B5", f["conflict"])
        self.run_copy()
        self.assertEqual((self.dst / "S" / "a.xlsx").read_bytes(), before)

    def test_destination_changed_after_preview_blocks(self):
        make(self.src / "A" / "a.xlsx", {"A1": "画面"}, image_note="修了图片")
        make(self.dst / "S" / "a.xlsx", {"A1": "画面"})
        distribute.preview(plan(self.folders), str(self.src), str(self.dst))
        edit(self.dst / "S" / "a.xlsx", B9="刚改")
        result = self.run_copy()
        self.assertEqual(result["folders"][0]["files"][0]["status"], "conflict")

    def test_old_record_but_same_content_overwrites(self):
        (self.src / "A").mkdir(parents=True)
        make(self.dst / "S" / "a.xlsx", {"A1": "画面"})
        distribute.preview_fetch(plan(self.folders), str(self.src), str(self.dst))
        self.run_copy()  # 记下版本
        make(self.dst / "S" / "a.xlsx", {"A1": "画面"}, image_note="别人只改了格式")
        make(self.src / "A" / "a.xlsx", {"A1": "画面"}, area="A1:C10", image_note="我修了图片")
        edit(self.src / "A" / "a.xlsx")  # 让两份字节不同
        prev = distribute.preview(plan(self.folders), str(self.src), str(self.dst))
        self.assertIn(prev["folders"][0]["files"][0]["guard"], ("content_same", "identical"))


class FreeUploadTests(JsonCopyTests):
    """起点 → 终点：没有记录时不再只是「要确认」，按内容决定"""

    def test_free_no_record(self):
        make(self.src / "a.xlsx", {"A1": "画面"}, image_note="修了图片")
        make(self.src / "b.xlsx", {"A1": "画面"})
        make(self.dst / "a.xlsx", {"A1": "画面"})
        make(self.dst / "b.xlsx", {"A1": "画面"}, area="A1:C20")
        prev = distribute.preview_free(str(self.src), str(self.dst))
        files = {f["rel"]: f for f in prev["files"]}
        self.assertEqual((files["a.xlsx"]["action"], files["a.xlsx"]["guard"]), ("overwrite", "content"))
        self.assertIn("核对一致", files["a.xlsx"]["note"])
        self.assertEqual(files["b.xlsx"]["action"], "conflict")
        self.assertIn("打印区域格子数", files["b.xlsx"]["error"])
        self.assertEqual(prev["summary"]["要确认"], 0)

    def test_free_content_only_ignores_manifest_and_baseline(self):
        make(self.dst / "a.xlsx", {"A1": "画面"})
        make(self.src / "a.xlsx", {"A1": "画面", "B2": "我改的"})
        auto = distribute.preview_free(str(self.src), str(self.dst))
        self.assertEqual(auto["files"][0]["action"], "conflict")  # 没记录、内容不一样
        make(self.src / "a.xlsx", {"A1": "画面"}, image_note="修了图片")
        strict = distribute.preview_free(str(self.src), str(self.dst), "content")
        self.assertEqual(strict["guard"], "content")
        self.assertEqual((strict["files"][0]["action"], strict["files"][0]["guard"]), ("overwrite", "content"))
        self.assertIn("选了只比内容", strict["files"][0]["note"])

    test_no_record_same_content_overwrites = test_no_record_different_content_blocks = None
    test_destination_changed_after_preview_blocks = test_old_record_but_same_content_overwrites = None


PREP = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [], "reference": False, "normalized": True}


class TaskUploadTests(unittest.TestCase):
    """当前任务上传：共享版本和记录对不上（拿很久以前的版本修的）时，内容一样就照常传"""

    def setUp(self):
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        root = Path(tmp.name)
        self.shared = root / "共享" / "仕様書"
        make(self.shared / "a.xlsx", {"A1": "画面", "B2": "説明"})
        for obj, name, value in ((storage, "DATABASE_PATH", root / "t.sqlite3"), (distribute, "_STATUS", {"running": False})):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        wf.STATE["busy"] = False
        self.task = work_task.create(str(self.shared), str(root / "handle-excel"))
        distribute.preview_fetch("", "", "", self.task["id"])
        distribute.start()
        self.wait()
        self.work = Path(self.task["work"])
        for p in (patch.object(upload_prep, "prepare", side_effect=lambda *a: dict(PREP)),
                  patch("core.excel_ops.start_background_excel", Mock()), patch("core.excel_ops.quit_background_excel", Mock()),
                  patch("core.file_problem.ask", Mock(return_value="skip"))):
            p.start()
            self.addCleanup(p.stop)

    def wait(self):
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def upload(self, guard=None):
        distribute.start_task_upload(self.task["id"], ["a.xlsx"], {}, None, guard)
        return self.wait()

    def test_content_only_blocks_even_when_hash_matches(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明（改）"})
        result = self.upload("content")
        item = result["files"][0]
        self.assertEqual(item["status"], "failed")
        self.assertIn("只比内容", item["error"])
        self.assertIn("B2", item["error"])
        # 默认（先比 hash）：共享没人动过，照常上传
        self.assertEqual(self.upload()["files"][0]["status"], "ok")

    def test_content_only_same_content_uploads(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明"}, image_note="我修了图片")
        item = self.upload("content")["files"][0]
        self.assertEqual(item["status"], "ok", item.get("error"))
        self.assertIn("只比内容", item["content_check"])

    def test_shared_reformatted_but_same_text_uploads(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明"}, image_note="我修了图片")
        make(self.shared / "a.xlsx", {"A1": "画面", "B2": "説明"}, image_note="别人另存过")
        edit(self.shared / "a.xlsx")
        result = self.upload()
        item = result["files"][0]
        self.assertEqual(item["status"], "ok", item.get("error"))
        self.assertIn("核对一致", item["content_check"])
        self.assertEqual((self.shared / "a.xlsx").read_bytes(), (self.work / "a.xlsx").read_bytes())

    def test_shared_text_changed_is_synced_not_overwritten(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明"}, image_note="我修了图片")
        mine = (self.work / "a.xlsx").read_bytes()
        make(self.shared / "a.xlsx", {"A1": "画面", "B2": "説明（別の人が修正）"})
        theirs = (self.shared / "a.xlsx").read_bytes()
        result = self.upload()
        item = result["files"][0]
        self.assertEqual(item["status"], "review")
        self.assertIn("按内容核对和共享不一样", item["error"])
        self.assertEqual((self.shared / "a.xlsx").read_bytes(), theirs)  # 共享没被盖
        self.assertEqual(Path(item["backup_path"]).read_bytes(), mine)  # 我改的先备份


if __name__ == "__main__":
    unittest.main()
