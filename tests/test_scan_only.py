import importlib
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import batch, fix_images, reporter, storage, workflow


class FakeCell:
    def __init__(self, addr, pattern=-4142, color=16777215, font=0):
        self.Address = addr
        self.Interior = Mock(Pattern=pattern, Color=color, TintAndShade=0)
        self.Font = Mock(Color=font)


def _shared(members, part, name):
    """整块读：每格一样返回那个值，不一样返回 None（和 Excel 一样）"""
    values = {getattr(getattr(m, part), name) for m in members}
    return values.pop() if len(values) == 1 else None


class FakeArea:
    def __init__(self, members):
        self.Cells = members
        self.Address = members[0].Address if len(members) == 1 else f"{members[0].Address}:{members[-1].Address}"
        self.MergeArea = self
        area = self

        class Part:
            def __init__(self, part):
                object.__setattr__(self, "part", part)

            def __getattr__(self, name):
                return _shared(area.Cells, self.part, name)

            def __setattr__(self, name, value):
                for m in area.Cells:
                    setattr(getattr(m, self.part), name, value)
        self.Interior, self.Font = Part("Interior"), Part("Font")


class FakeSheet:
    def __init__(self):
        self.cells = {}
        self.merges = {}

    def cell(self, addr):
        key = "$" + addr[0] + "$" + addr[1:] if not addr.startswith("$") else addr
        return self.cells.setdefault(key, FakeCell(key))

    def Range(self, addr, *corner):
        if corner:  # Range(左上格, 右下格)
            addr = f"{addr.Address}:{corner[0].Address}"
        cell = self.cell(addr.split(":")[0])
        area = FakeArea([self.cell(a) for a in self.merges.get(cell.Address, [cell.Address])])
        if ":" in addr:
            return area
        cell.MergeArea = area
        return cell


class ScanOnlyTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for obj, name, value in (
            (storage, "DATABASE_PATH", self.root / "state.sqlite3"),
            (reporter, "RECORD_PATH", self.root / "record.md"),
            (fix_images, "IMAGES_DIR", self.root / "images"),
        ):
            p = patch.object(obj, name, value)
            p.start()
            self.addCleanup(p.stop)
        self.wf = importlib.reload(workflow)
        self.path = str(self.root / "book.xlsx")
        Path(self.path).write_bytes(b"workbook")
        self.wf.STATE.update(folder=str(self.root), items={}, pages={},
                             batch_review={"cards": {}, "rows": {}, "ignored": []},
                             job={"kind": "batch_scan", "options": {}, "completed": [], "pending_adjustment": None})
        self.sheet = FakeSheet()
        self.sheet.Name = "S1"
        self.book = Mock(ReadOnly=False)
        self.book.Sheets.return_value = self.sheet
        self.book.Names.side_effect = KeyError  # 默认当作文件里没有这个名称，按原地址还原
        self.excel = Mock()
        self.excel.Workbooks.Open.return_value = self.book
        patch.object(batch, "normalize_workbook_view").start()
        self.real_matches = batch._marked_matches
        self.matches = patch.object(batch, "_marked_matches", return_value=[]).start()
        patch.object(self.wf, "_sync_problem_counts").start()
        self.addCleanup(patch.stopall)

    def add_item(self, address, row=5, **extra):
        iid = f"i{address}"
        self.wf.STATE["items"][iid] = {
            "id": iid, "seq": len(self.wf.STATE["items"]) + 1, "kind": "hidden", "text": address,
            "wb_path": self.path, "wb_name": "book.xlsx", "sheet_name": "S1", "cell_addr": address,
            "excel_row": row, "status": "pending", "image": "", **extra}
        return iid

    def scan(self, style=None):
        batch._data().setdefault("highlight_files", {})[self.path] = batch.highlight_style(style)
        batch._hand_over(self.path)
        batch._sync_highlights(lambda: self.excel, self.path)

    def card(self, address):
        return batch._data()["cards"][batch._key(self.path, "S1", address)]

    def test_all_problems_go_to_manual_and_cells_are_colored(self):
        self.add_item("A5")
        self.add_item("B7", row=7)
        self.scan()
        for addr in ("A5", "B7"):
            card = self.card(addr)
            self.assertEqual((card["color"], card["manual"], card["finalized"], card["waiting"]),
                             ("yellow", True, True, False))
            cell = self.sheet.cell(addr)
            self.assertEqual(cell.Interior.Pattern, batch.XL_PATTERN_SOLID)
            self.assertEqual(cell.Interior.Color, batch._excel_color("#C0007A"))
            self.assertEqual(cell.Font.Color, batch._excel_color("#FFFFFF"))
        self.assertEqual(self.wf.STATE["items"]["iA5"]["status"], "manual_review")
        self.book.Save.assert_called_once()
        # 行高没动：只扫描不写行高
        self.assertEqual(batch._data()["rows"], {})

    def test_custom_colors_and_validation(self):
        self.add_item("A5")
        self.scan({"fill": "#1d4ed8", "font": "#ffff00"})
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#1D4ED8"))
        self.assertEqual(self.sheet.cell("A5").Font.Color, 0x00FFFF)
        with self.assertRaises(ValueError):
            batch.highlight_style({"fill": "#000000", "font": "#000000"})
        self.assertEqual(batch.highlight_style({"fill": "red"}), batch.DEFAULT_HIGHLIGHT)

    def test_recheck_pass_restores_original_colors_but_keeps_user_changes(self):
        self.sheet.cells["$A$5"] = FakeCell("$A$5", pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        # 用户人工处理时自己把 B5 改成别的底色
        self.sheet.cell("B5").Interior.Color = 0x123456
        for addr in ("A5", "B5"):
            self.card(addr).update(color="blue", verified=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        a5, b5 = self.sheet.cell("A5"), self.sheet.cell("B5")
        self.assertEqual((a5.Interior.Pattern, a5.Interior.Color, a5.Font.Color), (1, 0xCCFFCC, 0x0000FF))
        self.assertEqual(b5.Interior.Color, 0x123456)
        self.assertEqual(b5.Font.Color, 0)  # 字色还是标记色时照常还原
        self.assertEqual(batch._marks(), {})

    def test_recheck_checks_with_original_colors_then_marks_this_rounds_problems(self):
        """复检：先把标记色全部还原再检测，查完只给这次查出的问题重新标色"""
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        seen = {}
        # 只读打开、内存里还原颜色、不保存就关：关掉时文件还是原来标着颜色的样子
        saved = {k: (c.Interior.Pattern, c.Interior.Color, c.Font.Color) for k, c in self.sheet.cells.items()}
        def discard(SaveChanges):
            for k, (pattern, color, font) in saved.items():
                cell = self.sheet.cells[k]
                cell.Interior.Pattern, cell.Interior.Color, cell.Font.Color = pattern, color, font
            self.book.Close.side_effect = None
        self.book.Close.side_effect = discard
        def refresh(*_, workbook=None, **__):
            self.assertIs(workbook, self.book)  # 检测用的就是还原过颜色的这一份，不再另外打开
            seen["A5"] = self.sheet.cell("A5").Interior.Pattern
            seen["B5"] = self.sheet.cell("B5").Interior.Pattern
            self.wf.STATE["items"]["iA5"]["status"] = "fixed"   # A5 改好了
            self.wf.STATE["items"]["iB5"]["status"] = "pending"  # B5 还有问题
            self.add_item("C5")                                  # 又冒出 C5
        saves = self.book.Save.call_count
        with patch.object(batch, "_refresh", side_effect=refresh):
            batch._recheck_file(self.excel, self.path, [self.card("A5")])
        self.assertEqual(self.book.Save.call_count, saves)  # 检测前不保存
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#C0007A"))  # 文件里还是标着的
        none = batch.XL_PATTERN_NONE
        self.assertEqual(seen, {"A5": none, "B5": none})  # 检测时是原来的颜色（无填充）
        batch._sync_highlights(lambda: self.excel, self.path)
        marked = batch._excel_color("#C0007A")
        self.assertTrue(self.card("A5")["dismissed"])
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, none)
        self.assertEqual(self.sheet.cell("B5").Interior.Color, marked)
        self.assertEqual(self.sheet.cell("C5").Interior.Color, marked)
        self.assertEqual((self.card("C5")["color"], self.card("C5")["finalized"]), ("yellow", True))

    def test_restore_follows_cells_moved_by_inserted_rows(self):
        """人工在上面插了 2 行：标色的 A5 跑到 A7，按标色时建的隐藏名称找到 A7 还原，不动现在的 A5"""
        self.add_item("A5")
        self.scan()
        mark = next(iter(batch._marks().values()))
        self.assertEqual(self.book.Names.Add.call_args.kwargs,
                         {"Name": mark["name"], "RefersTo": "='S1'!$A$5", "Visible": False})
        marked = batch._excel_color("#C0007A")
        moved, stays = self.sheet.cell("A7"), self.sheet.cell("A5")
        moved.Interior.Pattern, moved.Interior.Color = batch.XL_PATTERN_SOLID, marked
        moved.Font.Color = batch._excel_color("#FFFFFF")
        stays.Interior.Pattern, stays.Interior.Color, stays.Font.Color = batch.XL_PATTERN_SOLID, marked, 0
        self.sheet.Cells = lambda row, col: self.sheet.cell(f"{chr(64 + col)}{row}")
        names = {mark["name"]: Mock(RefersToRange=Mock(Row=7, Column=1, Worksheet=self.sheet))}
        self.book.Names.side_effect = lambda name: names[name]
        self.card("A5").update(color="blue")
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(moved.Interior.Pattern, batch.XL_PATTERN_NONE)
        self.assertEqual(moved.Font.Color, 0)
        self.assertEqual(stays.Interior.Pattern, batch.XL_PATTERN_SOLID)  # 现在的 A5 是别的格子，不动
        names[mark["name"]].Delete.assert_called_once()  # 还原后名称删掉，交出去的文件里不留

    def test_cell_deleted_by_user_needs_no_restore(self):
        self.add_item("A5")
        self.scan()
        mark = next(iter(batch._marks().values()))
        broken = Mock()
        type(broken).RefersToRange = property(lambda _: (_ for _ in ()).throw(RuntimeError("#REF!")))
        self.book.Names.side_effect = lambda name: broken
        self.card("A5").update(color="blue")
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(batch._marks(), {})
        broken.Delete.assert_called_once()

    def moved_mark(self, *others):
        """旧记录没有隐藏名称、被插了 2 行：A5 的标记色挪到了 A7，现在的 A5 是插进来的空行"""
        for addr in ("A5",) + others:
            self.add_item(addr)
        self.scan()
        mark = next(m for m in batch._marks().values() if m["addr"].replace("$", "") == "A5")
        mark.pop("name", None)
        mark["value"] = "合計"
        a5 = self.sheet.cell("A5")
        a5.Interior.Pattern, a5.Font.Color = batch.XL_PATTERN_NONE, 0
        moved = self.sheet.cell("A7")
        moved.Interior.Pattern, moved.Interior.Color = batch.XL_PATTERN_SOLID, batch._excel_color("#C0007A")
        moved.Font.Color = 0xFFFFFF
        self.sheet.Cells = lambda row, col: self.sheet.cell(f"{chr(64 + col)}{row}")
        self.card("A5").update(color="blue")
        return mark, moved

    def test_moved_marker_color_is_found_by_content_and_restored(self):
        mark, moved = self.moved_mark()
        self.matches.return_value = [(7, 1)]
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.matches.assert_called_once_with(self.sheet, "合計", mark["fill"])
        self.assertEqual((moved.Interior.Pattern, moved.Font.Color), (batch.XL_PATTERN_NONE, 0))
        self.assertTrue(any("按格子内容找到并还原" in c.args[0] for c in log.call_args_list))

    def test_several_matching_cells_are_left_alone_and_listed(self):
        _, moved = self.moved_mark()
        self.matches.return_value = [(7, 1), (9, 3)]
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(moved.Interior.Pattern, batch.XL_PATTERN_SOLID)  # 不确定就不改
        messages = [c.args[0] for c in log.call_args_list]
        self.assertTrue(any("不止一个" in m and "S1!A7、C9" in m for m in messages), messages)

    def test_cells_still_marked_on_purpose_do_not_count(self):
        _, moved = self.moved_mark("B5")
        self.matches.return_value = [(5, 2), (7, 1)]  # B5 还要继续标着，不算
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(moved.Interior.Pattern, batch.XL_PATTERN_NONE)

    def test_no_search_when_every_mark_was_restored_in_place(self):
        """平常都在原位置还原了：不做任何查找，不多花时间，也不报残留"""
        self.add_item("A5")
        self.scan()
        self.card("A5").update(color="blue")
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.matches.assert_not_called()
        self.assertFalse(any("⚠️" in c.args[0] for c in log.call_args_list))

    def test_one_mark_that_cannot_be_removed_does_not_fail_the_recheck(self):
        self.add_item("A5")
        self.scan()
        with patch.object(batch, "_restore_cells", side_effect=RuntimeError("例外が発生しました。")), \
                patch.object(batch.wf, "log_msg") as log:
            book = batch._open_without_marks(self.excel, self.path)
        self.assertIsNotNone(book)
        self.assertTrue(any("没能在检测前去掉" in c.args[0] and "S1!" in c.args[0] for c in log.call_args_list))

    def test_colors_left_behind_after_a_crash_are_found_later(self):
        """复检通过后 Excel 崩溃、颜色没还原：记录还在，之后提交 / 复检时顺带还原"""
        self.add_item("A5")
        self.scan()
        self.assertEqual(batch._leftover_mark_paths([]), [])  # 黄卡还要标着
        self.card("A5").update(dismissed=True, color="blue")
        self.assertEqual(batch._leftover_mark_paths([]), [self.path])
        self.assertEqual(batch._leftover_mark_paths([self.path]), [])  # 这次任务本来就会处理

    def test_marked_matches_keeps_only_cells_still_in_marker_color(self):
        fill = batch._excel_color("#C0007A")
        def cell(addr, row, col, color):
            return Mock(Address=addr, MergeArea=Mock(Row=row, Column=col, Interior=Mock(
                Color=color, Pattern=batch.XL_PATTERN_SOLID)))
        hits = [cell("$A$7", 7, 1, fill), cell("$B$7", 7, 1, fill), cell("$C$9", 9, 3, 0xFFFF00)]
        used = Mock()
        used.Find.return_value = hits[0]
        used.FindNext.side_effect = hits[1:] + [hits[0]]
        sheet = Mock(UsedRange=used)
        self.assertEqual(self.real_matches(sheet, "*合計?", fill), [(7, 1)])  # 合并区域只算一次，黄底的不算
        self.assertEqual(used.Find.call_args.kwargs["What"], "~*合計~?")  # 通配符按原字找
        self.assertEqual(self.real_matches(sheet, "", fill), [])

    def test_merged_area_and_mixed_font_color(self):
        self.sheet.merges["$A$5"] = ["$A$5", "$B$5"]
        self.sheet.cells["$A$5"] = FakeCell("$A$5", font=None)
        self.add_item("A5")
        self.scan()
        self.assertEqual(self.sheet.cell("B5").Interior.Color, batch._excel_color("#C0007A"))
        self.assertIsNone(self.sheet.cell("A5").Font.Color)  # 混合字色不动
        self.card("A5").update(dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.cell("B5").Interior.Pattern, batch.XL_PATTERN_NONE)

    def test_uniform_merged_area_is_read_once_and_restored_after_shift(self):
        """方眼紙的大合并区域颜色一致：整块记一条，不一格一格读；插行挪了位置也整块还原"""
        block = [f"${c}$5" for c in "ABCDEF"]
        self.sheet.merges["$A$5"] = block
        for addr in block:
            self.sheet.cells[addr] = FakeCell(addr, pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.scan()
        mark = next(iter(batch._marks().values()))
        self.assertEqual(mark["cells"], [{"addr": "$A$5:$F$5", "pattern": 1, "color": 0xCCFFCC,
                                          "tint": 0.0, "font": 0x0000FF}])
        self.assertTrue(mark["saved"])
        marked = batch._excel_color("#C0007A")
        self.assertTrue(all(self.sheet.cell(a).Interior.Color == marked for a in block))
        # 人工在上面插了 2 行：整块挪到第 7 行
        moved = [f"${c}$7" for c in "ABCDEF"]
        for old, new in zip(block, moved):
            src = self.sheet.cell(old)
            self.sheet.cells[new] = FakeCell(new, pattern=src.Interior.Pattern, color=src.Interior.Color,
                                             font=src.Font.Color)
        self.sheet.merges["$A$7"] = moved
        self.sheet.Cells = lambda row, col: self.sheet.cell(f"{chr(64 + col)}{row}")
        names = {mark["name"]: Mock(RefersToRange=Mock(Row=7, Column=1, Worksheet=self.sheet))}
        self.book.Names.side_effect = lambda name: names[name]
        self.card("A5").update(color="blue")
        batch._sync_highlights(lambda: self.excel, self.path)
        for addr in moved:
            cell = self.sheet.cell(addr)
            self.assertEqual((cell.Interior.Pattern, cell.Interior.Color, cell.Font.Color), (1, 0xCCFFCC, 0x0000FF))

    def test_partly_recolored_merged_area_restores_only_marked_cells(self):
        block = ["$A$5", "$B$5"]
        self.sheet.merges["$A$5"] = block
        self.add_item("A5")
        self.scan()
        self.sheet.cell("B5").Interior.Color = 0x123456  # 人工改了其中一格
        self.card("A5").update(color="blue")
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, batch.XL_PATTERN_NONE)
        self.assertEqual(self.sheet.cell("B5").Interior.Color, 0x123456)

    def test_interrupted_highlight_is_colored_again_next_time(self):
        """标色到一半被打断（看门狗结束 Excel）：记录存了、颜色没存进文件。重扫时要重新标，不能当作已标过"""
        self.add_item("A5")
        self.book.Save.side_effect = RuntimeError("Excel 被结束")
        with self.assertRaises(RuntimeError):
            self.scan()
        mark = next(iter(batch._marks().values()))
        self.assertFalse(mark["saved"])
        self.assertFalse(batch.snapshot()["cards"][0]["highlighted"])
        self.sheet.cells.clear()  # 没保存：文件里还是原来的颜色
        self.book.Save.side_effect = None
        with patch.object(batch.wf, "log_msg") as log:
            self.scan()
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#C0007A"))
        self.assertTrue(next(iter(batch._marks().values()))["saved"])
        self.assertTrue(any("没有存进文件" in c.args[0] for c in log.call_args_list))
        self.assertTrue(batch.snapshot()["cards"][0]["highlighted"])

    def test_old_record_still_colored_is_confirmed_without_saving(self):
        self.add_item("A5")
        self.scan()
        next(iter(batch._marks().values())).pop("saved")  # 旧版本的记录没有这一项
        saves = self.book.Save.call_count
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.book.Save.call_count, saves)  # 颜色还在：只确认，不改文件
        self.assertTrue(next(iter(batch._marks().values()))["saved"])
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.excel.Workbooks.Open.call_count, 2)  # 确认过以后不用再打开

    def test_every_cell_tells_watchdog_it_is_alive(self):
        for row in range(5, 65):
            self.add_item(f"A{row}", row=row)
        batch._data().setdefault("highlight_files", {})[self.path] = batch.highlight_style(None)
        batch._hand_over(self.path)
        tick = Mock()
        batch._sync_highlights(lambda: self.excel, self.path, tick=tick)
        self.assertGreaterEqual(tick.call_count, 120)  # 记原色、标色各一次
        self.assertTrue(self.excel.ScreenUpdating)  # 结束后恢复

    def test_unconfirmed_location_is_not_colored(self):
        self.add_item("A5", location_confirmed=False)
        self.scan()
        self.assertIn("未在 Excel 中标色", self.card("A5")["message"])
        self.excel.Workbooks.Open.assert_not_called()

    def test_clear_restores_everything(self):
        self.add_item("A5")
        self.scan()
        batch._sync_highlights(lambda: self.excel, self.path, clear=True)
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, batch.XL_PATTERN_NONE)
        self.assertEqual(batch._marks(), {})

    def use_named_sheets(self, sheets):
        """book.Sheets(name) 像 Excel 一样：名字不存在时抛 DISP_E_BADINDEX，也能遍历"""
        class Sheets:
            def __call__(_, name):
                for sheet in sheets:
                    if sheet.Name.lower() == name.lower():
                        return sheet
                raise Exception(-2147352567, "例外が発生しました。", (0, None, None, None, 0, -2147352565), None)

            def __iter__(_):
                return iter(sheets)
        self.book.Sheets = Sheets()

    def test_deleted_sheet_skips_restore_without_warning(self):
        self.sheet.Name = "S1"
        self.use_named_sheets([self.sheet])
        self.add_item("A5")
        self.scan()
        self.assertEqual(len(batch._marks()), 1)
        # 人工处理时整张表被删掉，复检后卡片变蓝
        self.use_named_sheets([])
        self.card("A5").update(color="blue", verified=True)
        with patch.object(self.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        messages = [c.args[0] for c in log.call_args_list]
        self.assertFalse(any("⚠️" in m for m in messages), messages)
        self.assertTrue(any("工作表「S1」已不在文件里" in m and "1 格" in m for m in messages), messages)
        self.assertEqual(batch._marks(), {})

    def test_sheet_name_with_extra_spaces_is_still_found(self):
        self.sheet.Name = "S1\u3000"
        self.use_named_sheets([self.sheet])
        self.add_item("A5")
        self.scan()
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#C0007A"))

    def test_highlight_on_missing_sheet_explains_why(self):
        self.use_named_sheets([])
        self.add_item("A5")
        self.scan()
        self.assertIn("工作表「S1」已不在文件里", self.card("A5")["message"])
        self.assertEqual(batch._marks(), {})

    def test_user_marked_cards_keep_their_color(self):
        self.add_item("A5")
        batch._card_for([self.wf.STATE["items"]["iA5"]]).update(color="green", marked=True)
        self.scan()
        self.assertEqual(self.card("A5")["color"], "green")
        self.excel.Workbooks.Open.assert_not_called()

    def test_run_scan_only_does_not_repair(self):
        self.add_item("A5")
        self.wf.STATE["job"]["options"] = {"folder": str(self.root), "scan_only": True,
                                           "highlight": {"fill": "#6D28D9", "font": "#FFFFFF"}}
        session = Mock()
        session.app.return_value = self.excel
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock(timed_out=False)), \
                patch.object(self.wf, "list_workbooks", return_value=(self.root, [Path(self.path)])), \
                patch.object(self.wf.pythoncom, "CoInitialize", create=True), \
                patch.object(self.wf.pythoncom, "CoUninitialize", create=True), \
                patch.object(batch, "init_execution_batch"), patch.object(batch, "finish_execution_batch"), \
                patch.object(batch, "_refresh"), patch.object(batch, "_repair_file") as repair:
            self.assertEqual(batch.run(), [self.path])
        repair.assert_not_called()
        self.assertEqual(self.card("A5")["color"], "yellow")
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#6D28D9"))
        self.assertIn(self.path, self.wf.STATE["job"]["completed"])

    def test_full_auto_colors_yellow_cards_only_after_commit(self):
        self.add_item("A5")
        self.wf.STATE["job"]["options"] = {"folder": str(self.root), "highlight": {"fill": "#006633", "font": "#FFFFFF"}}
        session = Mock()
        session.app.return_value = self.excel
        mocks = dict(ExcelWatchdog=Mock(return_value=Mock(timed_out=False)), init_execution_batch=Mock(),
                     finish_execution_batch=Mock(), _refresh=Mock())
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.multiple(batch, **mocks), \
                patch.object(self.wf, "list_workbooks", return_value=(self.root, [Path(self.path)])), \
                patch.object(self.wf.pythoncom, "CoInitialize", create=True), \
                patch.object(self.wf.pythoncom, "CoUninitialize", create=True), \
                patch.object(batch, "_repair_file", side_effect=lambda *a: batch._card_for(
                    [self.wf.STATE["items"]["iA5"]]).update(color="yellow", manual=True, waiting=False)):
            batch.run()
            self.excel.Workbooks.Open.assert_not_called()  # 处理过程中不改颜色
            card = self.card("A5")
            self.wf.STATE["job"] = {"kind": "batch_commit", "options": {"card_ids": [card["id"]]},
                                    "completed": [], "pending_adjustment": None}
            batch.run()
        self.assertEqual(self.sheet.cell("A5").Interior.Color, batch._excel_color("#006633"))

    def test_full_auto_without_highlight_stops_coloring(self):
        self.add_item("A5")
        self.scan()
        batch._data()["highlight_files"].pop(self.path)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, batch.XL_PATTERN_NONE)


if __name__ == "__main__":
    unittest.main()
