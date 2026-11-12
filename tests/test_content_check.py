"""轻量内容核对，以及上传时需要确认的完整 Excel 差异。"""
import json
import tempfile
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from openpyxl import Workbook, load_workbook

from core import content_check, distribute, excel_diff, storage, upload_prep, work_task, workflow as wf


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

    def test_sheets_clean_will_delete_are_ignored(self):
        a = make(self.root / "a.xlsx", {"A1": "x"}, sheets=("仕様",))
        b = make(self.root / "b.xlsx", {"A1": "x"}, sheets=("仕様", "変更履歴", "仕様_20220101"))
        result = content_check.compare(a, b, "本地", "共享")
        self.assertTrue(result["same"], result["diffs"])
        self.assertEqual(set(result["skipped"]), {"変更履歴", "仕様_20220101"})
        self.assertIn("変更履歴", content_check.skipped_note(result))
        # 正式的表多了一张照样拦
        c = make(self.root / "c.xlsx", {"A1": "x"}, sheets=("仕様", "変更履歴", "画面一覧"))
        self.assertEqual(content_check.compare(a, c, "本地", "共享")["diffs"], ["表「画面一覧」只有共享有"])

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


def wait_until(predicate):
    deadline = time.monotonic() + 5
    while time.monotonic() < deadline:
        result = predicate()
        if result:
            return result
        time.sleep(0.02)
    raise AssertionError(f"上传状态没有及时更新：{distribute.status()}")


def next_review(rel):
    return wait_until(lambda: next((r for r in distribute.status().get("reviews", []) if r["rel"] == rel), None))


class FreeUploadTests(JsonCopyTests):
    def test_no_base_requires_each_review_and_uploads_while_next_diff_runs(self):
        from web_clip_fixer import WebHandler

        def request(action, token, **payload):
            handler = object.__new__(WebHandler)
            handler.path = "/api/distribute/upload_" + action
            handler._send_json = Mock()
            handler._handle_post({"token": token, **payload})
            return handler._send_json.call_args.args[0]

        a = make(self.src / "a.xlsx", {"A1": "同样文字"}, image_note="修了格式")
        b = make(self.src / "b.xlsx", {"A1": "本地 B"})
        bad = self.src / "c.xlsx"
        bad.write_text("broken", encoding="utf-8")
        target_a = make(self.dst / "a.xlsx", {"A1": "同样文字"})
        target_b = make(self.dst / "b.xlsx", {"A1": "远程 B"})
        make(self.dst / "c.xlsx", {"A1": "远程 C"})
        old_a, old_b = target_a.read_bytes(), target_b.read_bytes()
        preview = distribute.preview_free(str(self.src), str(self.dst))
        self.assertEqual([f["action"] for f in preview["files"]], ["review"] * 3)
        self.assertEqual(preview["summary"]["要确认"], 3)
        self.assertEqual(len(distribute.pdf_check(preview_id=preview["id"])["files"]), 3)

        continue_b = threading.Event()
        b_started = threading.Event()
        compare = excel_diff.diff

        def hold_b(mine, theirs):
            if mine.name == "b.xlsx":
                b_started.set()
                if not continue_b.wait(8):
                    raise AssertionError("timed out waiting to compare b")
            return compare(mine, theirs)

        with (patch.object(excel_diff, "diff", side_effect=hold_b),
              patch.object(upload_prep, "prepare", return_value=dict(PREP))):
            distribute.start_free(preview["id"])
            try:
                a_review = next_review("a.xlsx")
                self.assertTrue(b_started.wait(5))
                self.assertEqual(target_a.read_bytes(), old_a)
                self.assertFalse(request("review", a_review["token"])["diff"]["identical"])
                request("decision", a_review["token"], decision="upload")
                wait_until(lambda: target_a.read_bytes() == a.read_bytes())
                self.assertTrue(distribute.status()["running"])
                self.assertEqual(target_b.read_bytes(), old_b)
                continue_b.set()
                b_review = next_review("b.xlsx")
                self.assertFalse(request("review", b_review["token"])["diff"]["identical"])
                request("decision", b_review["token"], decision="skip")
                wait_until(lambda: not distribute.status()["running"])
            finally:
                continue_b.set()
        self.assertEqual(target_b.read_bytes(), old_b)
        self.assertEqual(load_workbook(b).active["A1"].value, "本地 B")
        self.assertEqual([f["status"] for f in distribute.status()["applied"]["files"]], ["ok", "skipped", "failed"])

    def test_single_file_from_diff_page_uploads_and_keeps_rest_of_list(self):
        """起点 → 终点：对比页「上传这一份」只传那一份、不再逐份确认，清单里其余的留着（用户 2026-11-12）"""
        a = make(self.src / "a.xlsx", {"A1": "本地 A"})
        make(self.src / "b.xlsx", {"A1": "本地 B"})
        target_a = make(self.dst / "a.xlsx", {"A1": "远程 A"})
        target_b = make(self.dst / "b.xlsx", {"A1": "远程 B"})
        old_b = target_b.read_bytes()
        preview = distribute.preview_free(str(self.src), str(self.dst))
        self.assertEqual(len(distribute.pdf_check(None, ["a.xlsx"], preview["id"])["files"]), 1)
        seen = {"mine": distribute._digest(a), "theirs": distribute._digest(target_a)}
        with self.assertRaisesRegex(ValueError, "又变了"):
            distribute.start_free(preview["id"], None, "a.xlsx", {**seen, "theirs": "别的版本"})
        with (patch.object(excel_diff, "diff") as compare, patch.object(upload_prep, "prepare", return_value=dict(PREP)),
              patch("core.excel_ops.start_background_excel", Mock()), patch("core.excel_ops.quit_background_excel", Mock())):
            distribute.start_free(preview["id"], None, "a.xlsx", seen)
            wait_until(lambda: not distribute.status()["running"])
        compare.assert_not_called()
        status = distribute.status()
        self.assertEqual(status["single"]["status"], "ok", status["single"])
        self.assertEqual(load_workbook(target_a).active["A1"].value, "本地 A")
        self.assertEqual(target_b.read_bytes(), old_b)
        self.assertEqual([f["rel"] for f in status["preview"]["files"]], ["b.xlsx"])  # 清单还在，只少了传成的那份
        self.assertEqual(status["preview"]["id"], preview["id"])
        self.assertEqual((status["preview"]["summary"]["要上传"], status["preview"]["summary"]["要确认"]), (1, 1))
        with self.assertRaisesRegex(ValueError, "清单里没有这份文件"):
            distribute.start_free(preview["id"], None, "a.xlsx", seen)

    test_no_record_same_content_overwrites = test_no_record_different_content_blocks = None
    test_destination_changed_after_preview_blocks = test_old_record_but_same_content_overwrites = None


PREP = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [], "reference": False, "normalized": True}


class TaskUploadTests(unittest.TestCase):
    """当前任务上传的哈希快路径和逐文件版本确认。"""

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

    def start_upload(self, guard=None):
        distribute.start_task_upload(self.task["id"], ["a.xlsx"], {}, None, guard)

    def test_matching_base_skips_diff_for_local_edits(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明（改）"})
        with patch.object(excel_diff, "diff") as compare:
            self.start_upload()
            item = self.wait()["files"][0]
        compare.assert_not_called()
        self.assertEqual(item["status"], "ok", item.get("error"))

    def test_always_compare_requires_confirmation_even_with_matching_base(self):
        local = make(self.work / "a.xlsx", {"A1": "画面", "B2": "説明（改）"})
        remote = self.shared / "a.xlsx"
        before = remote.read_bytes()
        self.start_upload("content")
        token = next_review("a.xlsx")["token"]
        self.assertEqual(remote.read_bytes(), before)
        self.assertFalse(distribute.upload_review(token)["diff"]["identical"])
        distribute.decide_upload_review(token, "upload")
        self.assertEqual(self.wait()["files"][0]["status"], "ok")
        self.assertEqual(remote.read_bytes(), local.read_bytes())

    def test_changed_remote_and_local_invalidate_confirmation(self):
        for change_remote in (True, False):
            with self.subTest(change_remote=change_remote):
                local = make(self.work / "a.xlsx", {"A1": "画面", "B2": "我的修改"})
                remote = make(self.shared / "a.xlsx", {"A1": "画面", "B2": "共享变过"})
                self.start_upload()
                token = next_review("a.xlsx")["token"]
                self.assertFalse(distribute.upload_review(token)["diff"]["identical"])
                edit(remote if change_remote else local, B2="确认前刚改")
                remote_expected, local_expected = remote.read_bytes(), local.read_bytes()
                distribute.decide_upload_review(token, "upload")
                item = self.wait()["files"][0]
                self.assertIn("发生变化", item["error"])
                self.assertEqual(remote.read_bytes(), remote_expected)
                self.assertEqual(local.read_bytes(), local_expected)

    def test_confirmed_on_diff_page_uploads_without_second_review(self):
        """对比页看过差异点「上传这一份」：还是那两个版本就直接传，不再逐份确认（用户 2026-11-12）"""
        local = make(self.work / "a.xlsx", {"A1": "画面", "B2": "我的修改"})
        remote = make(self.shared / "a.xlsx", {"A1": "画面", "B2": "共享变过"})
        seen = {"a.xlsx": {"mine": distribute._digest(local), "theirs": distribute._digest(remote)}}
        with patch.object(excel_diff, "diff") as compare:
            distribute.start_task_upload(self.task["id"], ["a.xlsx"], {}, None, None, seen)
            item = self.wait()["files"][0]
        compare.assert_not_called()
        self.assertEqual(item["status"], "ok", item.get("error"))
        self.assertTrue(item["review_confirmed"])
        self.assertEqual(remote.read_bytes(), local.read_bytes())

    def test_confirmation_from_diff_page_is_void_after_remote_changes(self):
        make(self.work / "a.xlsx", {"A1": "画面", "B2": "我的修改"})
        remote = make(self.shared / "a.xlsx", {"A1": "画面", "B2": "共享变过"})
        seen = {"a.xlsx": {"mine": distribute._digest(self.work / "a.xlsx"), "theirs": distribute._digest(remote)}}
        edit(remote, B2="看完差异后别人又改了")
        before = remote.read_bytes()
        distribute.start_task_upload(self.task["id"], ["a.xlsx"], {}, None, None, seen)
        token = next_review("a.xlsx")["token"]  # 没直接传，回到逐份确认
        self.assertEqual(remote.read_bytes(), before)
        distribute.decide_upload_review(token, "skip")
        self.assertEqual(self.wait()["files"][0]["status"], "skipped")

    def test_review_covers_prepared_bytes_when_preparation_changes_the_workbook(self):
        local = make(self.work / "a.xlsx", {"A1": "处理前"})
        remote = self.shared / "a.xlsx"
        before_local, before_remote = local.read_bytes(), remote.read_bytes()
        seen = {"a.xlsx": {"mine": distribute._digest(local), "theirs": distribute._digest(remote)}}

        def prepare(_get_excel, path, _reference, _where):
            edit(path, A1="处理后")
            return dict(PREP)

        with patch.object(upload_prep, "prepare", side_effect=prepare) as prep:
            distribute.start_task_upload(self.task["id"], ["a.xlsx"], {}, None, None, seen)
            token = next_review("a.xlsx")["token"]
            try:
                review = distribute.upload_review(token)
                diff = review["diff"]
                prepared = Path(diff["mine"]).read_bytes()
                self.assertTrue(review["prepared"])
                self.assertIn("处理后", [change["new"] for change in diff["sheets"][0]["text"]])
                self.assertEqual(local.read_bytes(), before_local)
                self.assertEqual(remote.read_bytes(), before_remote)
            finally:
                distribute.decide_upload_review(token, "upload")
                item = self.wait()["files"][0]
        self.assertEqual(item["status"], "ok", item.get("error"))
        self.assertEqual(local.read_bytes(), prepared)
        self.assertEqual(remote.read_bytes(), prepared)
        prep.assert_called_once()

    def test_cancelling_review_releases_upload_and_invalidates_tokens(self):
        local = make(self.work / "a.xlsx", {"A1": "本地"})
        remote = self.shared / "a.xlsx"
        before = local.read_bytes(), remote.read_bytes()
        self.start_upload("content")
        token = next_review("a.xlsx")["token"]
        distribute.cancel_upload_review(distribute.status()["review_id"])
        self.assertEqual(self.wait()["files"][0]["status"], "skipped")
        with self.assertRaises(ValueError):
            distribute.decide_upload_review(token, "upload")
        self.assertEqual((local.read_bytes(), remote.read_bytes()), before)
        self.assertFalse(wf.STATE["busy"])



if __name__ == "__main__":
    unittest.main()
