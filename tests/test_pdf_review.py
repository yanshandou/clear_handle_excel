"""PDF 人工复检：列出 PDF、按页打开 Excel（和原始备份）、第几页 → Excel 里的范围、打开标黄 / 删除标红、重新生成。"""
import copy
import tempfile
import time
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from core import page_breaks, pdf_review, workflow as wf
from core.pdf_detector import fitz


class ReviewFolderTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.book = self.tmp / "営業" / "売上.xlsx"
        self.book.parent.mkdir()
        self.book.write_bytes(b"x")
        self.pdf = self.book.parent / "売上_Sheet1.pdf"
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        patcher = patch.object(wf, "STATE", self.state)
        patcher.start()
        self.addCleanup(patcher.stop)
        persist = patch.object(wf, "_persist_state")
        persist.start()
        self.addCleanup(persist.stop)

    def test_list_guesses_owner(self):
        listed = pdf_review.list_pdfs(str(self.tmp))
        entry = listed["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "Sheet1", True))
        self.assertEqual(entry["pages"], [[595.0, 842.0]])
        self.assertEqual(self.state["pdf_review"], {"folder": str(self.tmp.resolve())})

    def test_export_record_wins_over_guess(self):
        self.state["pdf_exports"][str(self.pdf)] = {"wb_path": str(self.book), "sheet": "Sheet/1"}
        entry = pdf_review.list_pdfs(str(self.tmp))["pdfs"][0]
        self.assertEqual((entry["sheet"], entry["guessed"]), ("Sheet/1", False))

    def test_pdfs_folder_is_listed_alone_and_manifest_finds_excel(self):
        # 2026-11-02 起 PDF 放在 根目录/PDFS/<子文件夹>；有 PDFS 时只看里面的，旁边混放的旧 PDF 不列
        from core import pdf_store
        inside = self.tmp / "PDFS" / "営業" / "売上_表_1.pdf"
        inside.parent.mkdir(parents=True)
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(inside)
        pdf_store.record(self.tmp, self.book, "h", [(inside, "表/1")])
        listed = pdf_review.list_pdfs(str(self.tmp))
        self.assertEqual([e["name"] for e in listed["pdfs"]], ["営業/売上_表_1.pdf"])
        entry = listed["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "表/1", False))
        self.assertTrue(listed["in_pdfs"])

    def test_pdfs_folder_without_record_guesses_excel_in_matching_subfolder(self):
        inside = self.tmp / "PDFS" / "営業" / "売上_Sheet2.pdf"
        inside.parent.mkdir(parents=True)
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(inside)
        entry = pdf_review.list_pdfs(str(self.tmp))["pdfs"][0]
        self.assertEqual((entry["wb_name"], entry["sheet"], entry["guessed"]), ("営業/売上.xlsx", "Sheet2", True))

    def opener(self, address="A21:H40", note=""):
        """假的可见 Excel：记下切到哪张表、第几页的范围"""
        calls = []

        def fake(path, sheet_name, locate):
            book = FakeBook(["Sheet1"])
            name = sheet_name(book)
            calls.append((path, name, locate))
            return name, address, note
        return patch.object(pdf_review, "open_workbook_at", side_effect=fake), calls

    def test_open_page_selects_that_pages_range(self):
        pdf_review.list_pdfs(str(self.tmp))
        opener, calls = self.opener()
        with opener, patch.object(pdf_review.repair_compare, "open_comparison") as compare, \
                patch.object(page_breaks, "page_area", return_value="A21:H40") as area:
            result = pdf_review.open_at(str(self.pdf), 1, 3)
            calls[0][2]("sheet")
        self.assertEqual(result, {"sheet": "Sheet1", "cell": "A21:H40", "backup": False, "note": ""})
        self.assertEqual(calls[0][:2], (self.book, "Sheet1"))
        area.assert_called_once_with("sheet", 1, 3)
        compare.assert_not_called()

    def test_page_not_matched_still_opens_sheet_with_note(self):
        pdf_review.list_pdfs(str(self.tmp))
        opener, _ = self.opener(address="")
        with opener:
            result = pdf_review.open_at(str(self.pdf), 0)
        self.assertEqual(result["cell"], "")
        self.assertIn("第 1 页在 Excel 里的位置没对上", result["note"])

    def test_open_with_backup_jumps_to_same_range(self):
        old = Path(tempfile.mkdtemp())
        (old / "営業").mkdir()
        backup = old / "営業" / "売上.xlsx"
        backup.write_bytes(b"old")
        pdf_review.list_pdfs(str(self.tmp))
        opener, _ = self.opener()
        with opener, patch.object(pdf_review.repair_compare, "open_comparison",
                                  return_value={"opened": ["current", "backup"], "note": ""}) as compare:
            result = pdf_review.open_at(str(self.pdf), 0, 1, str(old))
        self.assertTrue(result["backup"])
        target = compare.call_args[0][0]
        self.assertEqual((target["current"], target["backup"], target["sheet"], target["cell"]),
                         (self.book.resolve(), backup.resolve(), "Sheet1", "A21:H40"))

    def test_backup_not_found_still_opens_current(self):
        old = Path(tempfile.mkdtemp())
        pdf_review.list_pdfs(str(self.tmp))
        opener, calls = self.opener()
        with opener:
            result = pdf_review.open_at(str(self.pdf), 0, 1, str(old))
            missing = pdf_review.open_at(str(self.pdf), 0, 1, str(old / "nope"))
        self.assertFalse(result["backup"])
        self.assertIn("找不到这份文件", result["note"])
        self.assertIn("不存在", missing["note"])
        self.assertEqual(len(calls), 2)

    def test_open_is_blocked_while_scan_still_needs_the_file(self):
        pdf_review.list_pdfs(str(self.tmp))
        self.state.update(busy=True, job={"kind": "batch_commit"})
        opener, calls = self.opener()
        with opener, self.assertRaises(ValueError), \
                patch.object(pdf_review.batch, "open_block_reason", return_value="正在处理这个文件"):
            pdf_review.open_at(str(self.pdf), 0, 1)
        self.assertEqual(calls, [])

    def test_rejects_pdf_outside_folder(self):
        pdf_review.list_pdfs(str(self.tmp))
        other = Path(tempfile.mkdtemp()) / "x.pdf"
        other.write_bytes(self.pdf.read_bytes())
        with self.assertRaises(ValueError):
            pdf_review.open_at(str(other), 0, 1)
        with self.assertRaises(ValueError):
            pdf_review.open_at(str(self.pdf), "0", 1)


class ForgetTests(unittest.TestCase):
    def test_forget_pdfs_drops_export_records(self):
        tmp = Path(tempfile.mkdtemp())
        pdf = str(tmp / "a.pdf")
        state = copy.deepcopy(wf.STATE)
        state["pdf_exports"] = {pdf: {"wb_path": "x", "sheet": "S"}}
        with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"):
            pdf_review.forget_pdfs([Path(pdf)])
        self.assertEqual(state["pdf_exports"], {})


class FakeRange:
    def __init__(self, row, col, rows, cols):
        self.Row, self.Column = row, col
        self.Rows, self.Columns = SimpleNamespace(Count=rows), SimpleNamespace(Count=cols)


class FakePageSheet:
    """分页：横向分页在第 21、41 行上方，纵向分页在 F 列左边；打印区域 A1:H50"""
    def __init__(self, order=1, area="$A$1:$H$50"):
        self.PageSetup = SimpleNamespace(PrintArea=area, Order=order)
        self.UsedRange = FakeRange(1, 1, 60, 8)

    def Range(self, address):
        assert address == "$A$1:$H$50"
        return FakeRange(1, 1, 50, 8)


class PageAreaTests(unittest.TestCase):
    """PDF 第几页 → Excel 里那一页打印的范围"""

    def area(self, page, pages=6, **kw):
        breaks = {"h": [(21, False), (41, True)], "v": [6]}
        with patch.object(page_breaks, "read_breaks", return_value=breaks):
            return page_breaks.page_area(FakePageSheet(**kw), page, pages)

    def test_down_then_over(self):
        # 默认「先列后行」：先往下 3 页，再到右边那一列
        self.assertEqual([self.area(p) for p in range(6)],
                         ["A1:E20", "A21:E40", "A41:E50", "F1:H20", "F21:H40", "F41:H50"])

    def test_over_then_down(self):
        self.assertEqual([self.area(p, order=2) for p in range(6)],
                         ["A1:E20", "F1:H20", "A21:E40", "F21:H40", "A41:E50", "F41:H50"])

    def test_page_count_mismatch_or_several_areas_is_not_guessed(self):
        self.assertIsNone(self.area(0, pages=5))
        self.assertIsNone(self.area(6))
        self.assertIsNone(self.area(0, area="$A$1:$B$2,$D$1:$E$2"))


class ClearOldPdfTests(unittest.TestCase):
    def test_deletes_only_this_workbooks_old_pdfs(self):
        tmp = Path(tempfile.mkdtemp())
        book = tmp / "★売上.xlsx"
        book.write_bytes(b"x")
        old = [tmp / "売上_Sheet1.pdf", tmp / "売上_削除した表_2.PDF"]
        keep = [tmp / "取扱説明書.pdf", tmp / "売上.pdf", tmp / "その他_売上_Sheet1.pdf"]
        for path in old + keep:
            path.write_bytes(b"%PDF")
        state = copy.deepcopy(wf.STATE)
        state.update(pdf_exports={str(old[0]): {"wb_path": str(book), "sheet": "Sheet1"}})
        with patch.object(wf, "STATE", state), patch.object(wf, "_persist_state"), patch.object(wf, "log_msg"):
            wf._clear_old_pdfs([book])
        self.assertTrue(all(not p.exists() for p in old))
        self.assertTrue(all(p.exists() for p in keep))
        self.assertEqual(state["pdf_exports"], {})


class FakeSheet:
    def __init__(self, name, visible=-1):
        self.Name, self.Visible = name, visible


class FakeBook:
    def __init__(self, names, hidden=()):
        self.sheets = [FakeSheet(n, 0 if n in hidden else -1) for n in names]
        self.saved = 0
        self.closed = False
        book = self

        class Sheets(list):
            def __call__(self, key):
                return next(s for s in book.sheets if s.Name == key)
        self.Sheets = Sheets(self.sheets)
        self.Worksheets = self.Sheets

    def Save(self):
        self.saved += 1

    def Close(self, SaveChanges=False):
        self.closed = True


class ProcessTests(unittest.TestCase):
    """一键处理：标红的表在后台删掉并保存、对应 PDF 删掉；标黄的打开 Excel"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.book_path = self.tmp / "売上.xlsx"
        self.book_path.write_bytes(b"x")
        self.pdfs = {}
        for sheet in ("一覧", "JSONサンプル", "明細"):
            pdf = self.tmp / f"売上_{sheet}.pdf"
            doc = fitz.open()
            doc.new_page(width=595, height=842)
            doc.save(pdf)
            self.pdfs[sheet] = pdf
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        for target, name, value in ((wf, "STATE", self.state), (pdf_review, "_PROCESS", {"running": False})):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        for target, name in ((wf, "_persist_state"), (wf, "log_msg"),
                             (pdf_review.file_problem, "Guard"), (pdf_review, "start_background_excel"),
                             (pdf_review.sheet_history, "record")):
            patcher = patch.object(target, name)
            patcher.start()
            self.addCleanup(patcher.stop)
        pdf_review.file_problem.Guard.return_value.problem.return_value = (False, "打不开")
        pdf_review.file_problem.Guard.return_value.closed = False
        self.fp = patch.object(pdf_review.sheet_history, "fingerprint", return_value="h")
        self.fp.start()
        self.addCleanup(self.fp.stop)
        pdf_review.list_pdfs(str(self.tmp))

    def run_process(self, red=(), yellow=(), book=None, deleted=None):
        book = book or FakeBook(["一覧", "JSONサンプル", "明細"])
        calls = []

        def fake_delete(b, names):
            calls.append(list(names))
            b.sheets[:] = [s for s in b.sheets if s.Name not in names]
            return []
        with patch.object(pdf_review, "open_workbook_checked", return_value=book), \
                patch.object(pdf_review, "delete_sheets_preserving_values", side_effect=fake_delete), \
                patch.object(pdf_review, "open_page", side_effect=lambda p, sheet, page, n: (sheet, "A1:H20" if page else "", "")) as opened, \
                patch.object(pdf_review, "workbook_locked", return_value=False):
            pdf_review.process([{"pdf": str(self.pdfs[s])} for s in red],
                               [{"pdf": str(self.pdfs[s]), **extra} for s, extra in yellow])
            for _ in range(200):
                if not pdf_review.process_status()["running"]:
                    break
                time.sleep(0.02)
        status = pdf_review.process_status()
        self.assertFalse(status["running"])
        return status["result"], calls, book, opened

    def test_red_sheets_are_deleted_and_saved_with_pdf_removed(self):
        result, calls, book, _ = self.run_process(red=["JSONサンプル"])
        self.assertEqual(calls, [["JSONサンプル"]])
        self.assertEqual(book.saved, 1)
        self.assertTrue(book.closed)
        self.assertEqual([(e["sheet"], e["pdf_removed"]) for e in result["deleted"]], [("JSONサンプル", True)])
        self.assertFalse(self.pdfs["JSONサンプル"].exists())
        self.assertTrue(self.pdfs["一覧"].exists())
        self.assertFalse(self.state["busy"])
        pdf_review.sheet_history.record.assert_called_once()

    def test_two_red_sheets_in_one_file_open_it_once(self):
        result, calls, book, _ = self.run_process(red=["JSONサンプル", "明細"])
        self.assertEqual(calls, [["JSONサンプル", "明細"]])
        self.assertEqual(book.saved, 1)
        self.assertEqual(len(result["deleted"]), 2)

    def test_last_visible_sheet_is_not_deleted(self):
        book = FakeBook(["一覧", "JSONサンプル", "明細"], hidden=("一覧", "明細"))
        result, calls, book, _ = self.run_process(red=["JSONサンプル"], book=book)
        self.assertEqual(calls, [])
        self.assertEqual(book.saved, 0)
        self.assertIn("唯一", result["failed"][0]["reason"])
        self.assertTrue(self.pdfs["JSONサンプル"].exists())

    def test_yellow_opens_each_file_once_at_the_first_marked_page(self):
        result, calls, _, opened = self.run_process(
            yellow=[("一覧", {"spots": [{"page": 1}, {"page": 0, "x": 210, "y": 110}]}),
                    ("明細", {"spots": [{"page": 0}]})])
        opened.assert_called_once_with(self.book_path, "一覧", 0, 1)  # 第 1 页那处先，PDF 只有 1 页
        self.assertEqual([(e["sheet"], e["page"]) for e in result["opened"]],
                         [("一覧", 0), ("一覧", 1), ("明細", 0)])
        self.assertEqual([e["first"] for e in result["opened"]], [True, False, False])
        self.assertEqual(calls, [])

    def test_red_then_yellow_in_same_run(self):
        result, calls, _, opened = self.run_process(red=["JSONサンプル"], yellow=[("明細", {})])
        self.assertEqual(calls, [["JSONサンプル"]])
        opened.assert_called_once_with(self.book_path, "明細", None, 1)

    def test_more_than_limit_files_wait_for_next_press(self):
        with patch.object(pdf_review, "OPEN_LIMIT", 0):
            result, _, _, opened = self.run_process(yellow=[("明細", {})])
        opened.assert_not_called()
        self.assertEqual(len(result["later"]), 1)

    def test_locked_file_asks_and_records_failure(self):
        with patch.object(pdf_review, "open_workbook_checked") as open_book, \
                patch.object(pdf_review, "workbook_locked", return_value=True):
            pdf_review.process([{"pdf": str(self.pdfs["明細"])}], [])
            for _ in range(200):
                if not pdf_review.process_status()["running"]:
                    break
                time.sleep(0.02)
        open_book.assert_not_called()
        guard = pdf_review.file_problem.Guard.return_value
        self.assertIn("Excel 中打开", str(guard.problem.call_args[0][0]))
        self.assertEqual(pdf_review.process_status()["result"]["failed"][0]["reason"], "打不开")
        self.assertTrue(self.pdfs["明細"].exists())

    def run_floor(self, sheets, found, unchanged=True, book=None, tools=None):
        book = book or FakeBook(["一覧", "JSONサンプル", "明細"])
        with patch.object(pdf_review, "open_workbook_checked", return_value=book), \
                patch.object(pdf_review, "workbook_locked", return_value=False), \
                patch.object(pdf_review.pdf_store, "unchanged", return_value=unchanged), \
                patch.object(pdf_review.page_fill, "floor_pages", side_effect=found) as floor, \
                patch.object(wf, "_export_workbook_pdfs") as export, \
                patch("core.reporter.suspend_updates"):
            pdf_review.process([], [], tools if tools is not None else [{"pdf": str(self.pdfs[s])} for s in sheets])
            for _ in range(200):
                if not pdf_review.process_status()["running"]:
                    break
                time.sleep(0.02)
        self.assertFalse(self.state["busy"])
        return pdf_review.process_status()["result"], book, floor, export

    def test_floor_saves_and_regenerates_pdfs(self):
        ok = {"ok": True, "reason": "", "before": 5, "after": 4, "scale": 0.96, "manual": False}
        bad = {"ok": False, "reason": "要把字缩到约 80% 才挤得进 4 页，太小了"}
        result, book, floor, export = self.run_floor(["一覧", "明細"], [ok, bad])
        self.assertEqual(floor.call_count, 2)
        self.assertEqual(book.saved, 1)
        export.assert_called_once()
        self.assertEqual(export.call_args.kwargs, {"force": True})
        self.assertEqual([(e["sheet"], e["reason"]) for e in result["changed"]], [("一覧", "5 页 → 4 页，字缩到约 96%")])
        self.assertEqual([e["sheet"] for e in result["change_failed"]], ["明細"])

    def test_floor_nothing_ok_is_not_saved(self):
        bad = {"ok": False, "reason": "只有 1 页，不能再少了"}
        result, book, _, export = self.run_floor(["一覧"], [bad])
        self.assertEqual(book.saved, 0)
        export.assert_not_called()
        self.assertEqual(result["change_failed"][0]["reason"], "只有 1 页，不能再少了")

    def test_floor_refuses_when_excel_changed_after_export(self):
        result, book, floor, _ = self.run_floor(["一覧"], [], unchanged=False)
        floor.assert_not_called()
        self.assertIn("重新生成", result["change_failed"][0]["reason"])

    def test_floor_dirty_sheet_blocks_save(self):
        ok = {"ok": True, "reason": "", "before": 5, "after": 4, "scale": 0.96, "manual": False}
        dirty = {"ok": False, "reason": "改不回去", "dirty": True}
        result, book, _, export = self.run_floor(["一覧", "明細"], [ok, dirty])
        self.assertEqual(book.saved, 0)
        export.assert_not_called()
        self.assertEqual(result["changed"], [])
        self.assertEqual(len(result["change_failed"]), 2)

    def test_orientation_and_auto_tall_are_saved_and_regenerated(self):
        # 调整页面（用户 2026-11-07）：改方向、页高改回自动，和取整一样改完保存、重新导出
        book = FakeBook(["一覧", "JSONサンプル", "明細"])
        book.sheets[0].PageSetup = SimpleNamespace(Orientation=1, Zoom=False, FitToPagesWide=1, FitToPagesTall=4)
        book.sheets[2].PageSetup = SimpleNamespace(Orientation=1, Zoom=False, FitToPagesWide=1, FitToPagesTall=3)
        tools = [{"pdf": str(self.pdfs["一覧"]), "tool": "orient", "value": 2},
                 {"pdf": str(self.pdfs["明細"]), "tool": "auto"}]
        result, book, floor, export = self.run_floor([], [], book=book, tools=tools)
        floor.assert_not_called()
        self.assertEqual(book.saved, 1)
        export.assert_called_once()
        first, third = book.sheets[0].PageSetup, book.sheets[2].PageSetup
        self.assertEqual((first.Orientation, first.FitToPagesTall), (2, False))
        self.assertEqual((third.Zoom, third.FitToPagesWide, third.FitToPagesTall), (False, 1, False))
        reasons = {e["sheet"]: (e["tool"], e["reason"]) for e in result["changed"]}
        self.assertEqual(reasons["一覧"][0], "orient")
        self.assertIn("竖向 → 横向", reasons["一覧"][1])
        self.assertIn("改回自动", reasons["一覧"][1])
        self.assertEqual(reasons["明細"], ("auto", "页高改回自动（宽一页、高由 Excel 定）"))

    def test_clear_manual_breaks_is_saved_and_regenerated(self):
        cleared = {"ok": True, "count": 3, "reason": ""}
        tools = [{"pdf": str(self.pdfs["明細"]), "tool": "breaks"}]
        with patch.object(pdf_review.page_fill, "clear_manual_breaks", return_value=cleared) as clear:
            result, book, floor, export = self.run_floor([], [], tools=tools)
        clear.assert_called_once()
        floor.assert_not_called()
        self.assertEqual(book.saved, 1)
        export.assert_called_once()
        self.assertEqual([(e["tool"], e["reason"]) for e in result["changed"]],
                         [("breaks", "清掉了 3 处手动分页，改成 Excel 自动分页")])

    def test_clear_manual_breaks_partly_left_is_not_saved(self):
        left = {"ok": False, "count": 1, "dirty": True, "reason": "还剩 1 处手动分页清不掉，请在 Excel 里手动处理"}
        tools = [{"pdf": str(self.pdfs["明細"]), "tool": "breaks"}]
        with patch.object(pdf_review.page_fill, "clear_manual_breaks", return_value=left):
            result, book, _, export = self.run_floor([], [], tools=tools)
        self.assertEqual(book.saved, 0)
        export.assert_not_called()
        self.assertIn("还剩 1 处", result["change_failed"][0]["reason"])

    def test_floor_waits_when_same_sheet_changes_orientation(self):
        book = FakeBook(["一覧", "JSONサンプル", "明細"])
        book.sheets[0].PageSetup = SimpleNamespace(Orientation=1, Zoom=False, FitToPagesWide=1, FitToPagesTall=False)
        tools = [{"pdf": str(self.pdfs["一覧"]), "tool": "floor"},
                 {"pdf": str(self.pdfs["一覧"]), "tool": "orient", "value": 2}]
        result, book, floor, _ = self.run_floor([], [], book=book, tools=tools)
        floor.assert_not_called()
        self.assertEqual([e["tool"] for e in result["changed"]], ["orient"])
        self.assertIn("改完方向看新的 PDF", result["change_failed"][0]["reason"])

    def test_unknown_tool_is_rejected(self):
        with self.assertRaises(ValueError):
            pdf_review.process([], [], [{"pdf": str(self.pdfs["一覧"]), "tool": "zoom"}])
        with self.assertRaises(ValueError):
            pdf_review.process([], [], [{"pdf": str(self.pdfs["一覧"]), "tool": "orient", "value": 3}])

    def test_rejects_busy_or_outside_pdf(self):
        self.state["busy"] = True
        with self.assertRaises(ValueError):
            pdf_review.process([{"pdf": str(self.pdfs["明細"])}], [])
        self.state["busy"] = False
        other = Path(tempfile.mkdtemp()) / "x.pdf"
        other.write_bytes(self.pdfs["明細"].read_bytes())
        with self.assertRaises(ValueError):
            pdf_review.process([{"pdf": str(other)}], [])
        with self.assertRaises(ValueError):
            pdf_review.process([], [])


class RegenerateTests(unittest.TestCase):
    """复检时只重新生成一份 Excel 的 PDF"""

    def setUp(self):
        from core import pdf_store
        self.store = pdf_store
        self.tmp = Path(tempfile.mkdtemp())
        self.book = self.tmp / "売上.xlsx"
        self.book.write_bytes(b"x")
        self.pdf = self.tmp / "PDFS" / "売上_一覧.pdf"
        self.pdf.parent.mkdir()
        doc = fitz.open()
        doc.new_page(width=595, height=842)
        doc.save(self.pdf)
        pdf_store.record(self.tmp, self.book, pdf_review.sheet_history.fingerprint(self.book), [(self.pdf, "一覧")])
        self.state = copy.deepcopy(wf.STATE)
        self.state.update(pdf_review={}, pdf_exports={}, busy=False, job=None)
        for target, name, value in ((wf, "STATE", self.state), (pdf_review, "_REGEN", {"running": False})):
            patcher = patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        for target, name in ((wf, "_persist_state"), (wf, "log_msg"),
                             (pdf_review.file_problem, "Guard"), (pdf_review, "start_background_excel")):
            patcher = patch.object(target, name)
            patcher.start()
            self.addCleanup(patcher.stop)
        pdf_review.list_pdfs(str(self.tmp))

    def wait(self):
        for _ in range(200):
            if not pdf_review.regen_status()["running"]:
                break
            time.sleep(0.02)
        status = pdf_review.regen_status()
        self.assertFalse(status["running"])
        return status["result"]

    def test_unchanged_excel_asks_first(self):
        with patch.object(wf, "_export_workbook_pdfs") as export:
            self.assertEqual(pdf_review.regenerate(str(self.pdf)),
                             {"started": False, "unchanged": True, "name": "売上.xlsx"})
        export.assert_not_called()

    def test_changed_or_confirmed_exports_only_this_workbook(self):
        self.book.write_bytes(b"edited and saved")
        with patch.object(wf, "_export_workbook_pdfs", return_value=[self.pdf]) as export:
            self.assertEqual(pdf_review.regenerate(str(self.pdf)), {"started": True, "name": "売上.xlsx"})
            result = self.wait()
        self.assertEqual(result["ok"], True)
        self.assertEqual(result["pdfs"], [str(self.pdf)])
        args, kwargs = export.call_args
        self.assertEqual((args[1], args[2], kwargs), (self.book, self.tmp.resolve(), {"force": True}))
        self.assertFalse(self.state["busy"])

    def test_failure_is_reported_and_busy_released(self):
        pdf_review.file_problem.Guard.return_value.problem.return_value = (False, "文件打不开")
        pdf_review.file_problem.Guard.return_value.closed = False
        with patch.object(wf, "_export_workbook_pdfs", side_effect=RuntimeError("boom")):
            pdf_review.regenerate(str(self.pdf), True)
            result = self.wait()
        self.assertEqual((result["ok"], result["reason"]), (False, "文件打不开"))
        self.assertFalse(self.state["busy"])

    def test_busy_is_rejected(self):
        self.state["busy"] = True
        with self.assertRaises(ValueError):
            pdf_review.regenerate(str(self.pdf), True)


if __name__ == "__main__":
    unittest.main()
