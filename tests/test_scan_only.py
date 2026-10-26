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


class FakeRule:
    """一条条件格式规则：sheet.rules 里越靠前优先级越高（和 Excel 一样，新加的排在最后）"""
    def __init__(self, sheet, cells, formula):
        self.sheet, self.cells, self.Formula1 = sheet, cells, formula
        self.Interior, self.Font, self.StopIfTrue = Mock(Color=None), Mock(Color=None), False

    def SetFirstPriority(self):
        self.sheet.rules.remove(self)
        self.sheet.rules.insert(0, self)

    def Delete(self):
        if self.sheet.protected:
            raise RuntimeError("例外が発生しました。")
        self.sheet.rules.remove(self)


class FakeRules:
    def __init__(self, sheet, cells=None):
        self.sheet, self.cells = sheet, cells

    @property
    def Count(self):
        return len(self.sheet.rules)

    def Item(self, index):
        return self.sheet.rules[index - 1]

    def Add(self, Type, Formula1):
        assert Type == 2  # xlExpression
        if self.sheet.protected or self.sheet.blocked:
            raise RuntimeError("例外が発生しました。")
        rule = FakeRule(self.sheet, list(self.cells), Formula1)
        self.sheet.rules.append(rule)
        return rule


class FakeArea:
    def __init__(self, sheet, members):
        self.Cells = members
        self.Address = members[0].Address if len(members) == 1 else f"{members[0].Address}:{members[-1].Address}"
        self.MergeArea = self
        self.FormatConditions = FakeRules(sheet, [m.Address for m in members])
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


class FakeCells:
    """sheet.Cells：可以 Cells(行, 列) 取一格，也有整张表的 FormatConditions"""
    def __init__(self, sheet):
        self.sheet = sheet
        self.FormatConditions = FakeRules(sheet)

    def __call__(self, row, col):
        return self.sheet.cell(f"{chr(64 + col)}{row}")


class FakeSheet:
    def __init__(self):
        self.cells = {}
        self.merges = {}
        self.rules = []
        self.protected = False
        self.blocked = False  # 共享工作簿等：不受保护也加不了条件格式
        self.Cells = FakeCells(self)

    @property
    def ProtectContents(self):
        return self.protected

    def cell(self, addr):
        key = "$" + addr[0] + "$" + addr[1:] if not addr.startswith("$") else addr
        return self.cells.setdefault(key, FakeCell(key))

    def Range(self, addr, *corner):
        if corner:  # Range(左上格, 右下格)
            addr = f"{addr.Address}:{corner[0].Address}"
        cell = self.cell(addr.split(":")[0])
        area = FakeArea(self, [self.cell(a) for a in self.merges.get(cell.Address, [cell.Address])])
        if ":" in addr:
            return area
        cell.MergeArea = area
        return cell

    def shown(self, addr):
        """这一格现在显示的标记（底色, 字色）：排在最前、盖着这一格的规则；没有返回 None"""
        key = "$" + addr[0] + "$" + addr[1:]
        for rule in self.rules:
            if key in rule.cells:
                return rule.Interior.Color, rule.Font.Color
        return None


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
        self.book.Worksheets = [self.sheet]
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

    def old_scan(self, style=None):
        """旧版本标的颜色：直接改单元格底色、字色，记下原色和隐藏名称。升级后这些记录还要能还原"""
        style = batch.highlight_style(style)
        batch._data().setdefault("highlight_files", {})[self.path] = style
        batch._hand_over(self.path)
        fill, font = batch._excel_color(style["fill"]), batch._excel_color(style["font"])
        for cid in batch._wanted_marks(self.path):
            card = batch._data()["cards"][cid]
            addr = card["cell_addr"]
            area = self.sheet.Range(addr).MergeArea
            mark = {"path": self.path, "sheet": "S1", "addr": "$" + addr[0] + "$" + addr[1:], "fill": fill,
                    "font": font, "name": batch._mark_name(cid), "saved": True, "cells": []}
            for cell in area.Cells:
                mark["cells"].append({"addr": cell.Address, "pattern": cell.Interior.Pattern,
                                      "color": cell.Interior.Color, "tint": 0.0, "font": cell.Font.Color})
                cell.Interior.Pattern, cell.Interior.Color = batch.XL_PATTERN_SOLID, fill
                if cell.Font.Color is not None:
                    cell.Font.Color = font
            batch._marks()[cid] = mark

    def user_rule(self, addr="A5"):
        """表里本来就有的条件格式"""
        rule = self.sheet.Range(addr).MergeArea.FormatConditions.Add(Type=2, Formula1="=$C$1>0")
        rule.Interior.Color = 0x00FFFF
        return rule

    def test_all_problems_go_to_manual_and_cells_are_colored(self):
        self.add_item("A5")
        self.add_item("B7", row=7)
        self.scan()
        for addr in ("A5", "B7"):
            card = self.card(addr)
            self.assertEqual((card["color"], card["manual"], card["finalized"], card["waiting"]),
                             ("yellow", True, True, False))
            self.assertEqual(self.sheet.shown(addr), (batch._excel_color("#C0007A"), batch._excel_color("#FFFFFF")))
            # 单元格本身的颜色不动，标记是条件格式规则
            self.assertEqual(self.sheet.cell(addr).Interior.Pattern, batch.XL_PATTERN_NONE)
        rule = self.sheet.rules[0]
        self.assertTrue(rule.StopIfTrue)
        self.assertRegex(rule.Formula1, r'^=N\("CHE_MARK_[0-9a-f]+"\)=0$')
        self.assertEqual(self.wf.STATE["items"]["iA5"]["status"], "manual_review")
        self.book.Save.assert_called_once()
        # 行高没动：只扫描不写行高
        self.assertEqual(batch._data()["rows"], {})

    def test_custom_colors_and_validation(self):
        self.add_item("A5")
        self.scan({"fill": "#1d4ed8", "font": "#ffff00"})
        self.assertEqual(self.sheet.shown("A5"), (batch._excel_color("#1D4ED8"), 0x00FFFF))
        with self.assertRaises(ValueError):
            batch.highlight_style({"fill": "#000000", "font": "#000000"})
        self.assertEqual(batch.highlight_style({"fill": "red"}), batch.DEFAULT_HIGHLIGHT)

    def test_recheck_pass_restores_original_colors_but_keeps_user_changes(self):
        self.sheet.cells["$A$5"] = FakeCell("$A$5", pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.add_item("B5")
        self.old_scan()
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

    def test_recheck_checks_without_marks_then_marks_this_rounds_problems(self):
        """复检：先把标记规则全部去掉再检测，查完只给这次查出的问题重新标"""
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        mine = self.user_rule("D5")
        seen = {}
        # 只读打开、内存里删规则、不保存就关：关掉时文件还是原来标着的样子
        saved = list(self.sheet.rules)
        def discard(SaveChanges):
            self.sheet.rules[:] = saved
            self.book.Close.side_effect = None
        self.book.Close.side_effect = discard
        def refresh(*_, workbook=None, **__):
            self.assertIs(workbook, self.book)  # 检测用的就是去掉标记的这一份，不再另外打开
            seen.update(rules=list(self.sheet.rules))
            self.wf.STATE["items"]["iA5"]["status"] = "fixed"   # A5 改好了
            self.wf.STATE["items"]["iB5"]["status"] = "pending"  # B5 还有问题
            self.add_item("C5")                                  # 又冒出 C5
        saves = self.book.Save.call_count
        with patch.object(batch, "_refresh", side_effect=refresh):
            batch._recheck_file(self.excel, self.path, [self.card("A5")])
        self.assertEqual(self.book.Save.call_count, saves)  # 检测前不保存
        self.assertIsNotNone(self.sheet.shown("A5"))  # 文件里还是标着的
        self.assertEqual(seen["rules"], [mine])  # 检测时只剩表里原有的规则
        batch._sync_highlights(lambda: self.excel, self.path)
        marked = (batch._excel_color("#C0007A"), batch._excel_color("#FFFFFF"))
        self.assertTrue(self.card("A5")["dismissed"])
        self.assertIsNone(self.sheet.shown("A5"))
        self.assertEqual(self.sheet.shown("B5"), marked)
        self.assertEqual(self.sheet.shown("C5"), marked)
        self.assertIn(mine, self.sheet.rules)
        self.assertEqual((self.card("C5")["color"], self.card("C5")["finalized"]), ("yellow", True))

    def test_only_our_rule_for_that_card_is_removed(self):
        """表里原有的条件格式不动、顺序不变；我们的规则排最前；只删这张卡自己的规则"""
        mine = self.user_rule("A5")
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        self.assertEqual(self.sheet.rules[-1], mine)
        self.assertEqual(self.sheet.shown("A5")[0], batch._excel_color("#C0007A"))  # 盖住原有规则
        self.card("A5").update(color="blue", dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.shown("A5"), (0x00FFFF, None))  # 原有规则又露出来
        self.assertEqual(self.sheet.shown("B5")[0], batch._excel_color("#C0007A"))
        self.assertEqual(len(self.sheet.rules), 2)
        self.assertEqual(list(batch._marks()), [self.card("B5")["id"]])

    def test_copied_rules_and_renamed_sheet_are_all_found_by_tag(self):
        """复制粘贴把规则带到了别的表、拆成几条，表也改了名：按标识整份找，一起删掉"""
        other = FakeSheet()
        other.Name = "コピー"
        self.add_item("A5")
        self.scan()
        tag = self.sheet.rules[0].Formula1
        other.Range("C9").MergeArea.FormatConditions.Add(Type=2, Formula1=tag)
        other.Range("D9").MergeArea.FormatConditions.Add(Type=2, Formula1=tag)
        self.sheet.Name = "改名後"
        self.book.Worksheets = [self.sheet, other]
        self.card("A5").update(color="blue", dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual((self.sheet.rules, other.rules), ([], []))
        self.assertEqual(batch._marks(), {})

    def test_rule_that_cannot_be_deleted_is_kept_and_retried(self):
        self.add_item("A5")
        self.scan()
        self.card("A5").update(color="blue", dismissed=True)
        self.sheet.protected = True
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        messages = [c.args[0] for c in log.call_args_list]
        self.assertEqual(len(batch._marks()), 1)
        self.assertTrue(any("记录留着下次再试" in m for m in messages), messages)
        self.sheet.protected = False
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual((batch._marks(), self.sheet.rules), ({}, []))
        self.assertTrue(any("去掉标记 1 格" in c.args[0] for c in log.call_args_list))

    def test_rule_deleted_by_hand_needs_nothing_and_card_stays_yellow(self):
        self.add_item("A5")
        self.scan()
        self.sheet.rules.clear()  # 用户自己在 Excel 里删了规则
        self.card("A5").update(color="blue", dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(batch._marks(), {})

    def test_protected_sheet_cannot_be_marked_and_says_so(self):
        self.sheet.protected = True
        self.add_item("A5")
        self.scan()
        self.assertIn("未能在 Excel 中标色", self.card("A5")["message"])
        self.assertIn("工作表受保护", self.card("A5")["message"])
        self.assertEqual(batch._marks(), {})

    def test_shared_workbook_cannot_be_marked_and_says_so(self):
        """共享工作簿：Excel 只回一个空的「例外が発生しました」，卡片和日志要说清楚是共享工作簿"""
        self.book.MultiUserEditing = True
        self.sheet.blocked = True  # 工作表没保护，但 Excel 不让加条件格式
        self.add_item("A5")
        with patch.object(batch.wf, "log_msg") as log:
            self.scan()
        self.assertIn("共享工作簿", self.card("A5")["message"])
        self.assertTrue(any("1 格未能标色（这是共享工作簿" in c.args[0] for c in log.call_args_list))

    def test_grouped_sheets_are_ungrouped_before_marking(self):
        """文件保存时几张表组合着（同时选中）：先只选中第一张可见的表，再标色"""
        self.sheet.Visible = -1
        self.sheet.Select = Mock()
        self.book.Windows.return_value.SelectedSheets.Count = 2
        self.add_item("A5")
        with patch.object(batch.wf, "log_msg") as log:
            self.scan()
        self.sheet.Select.assert_called_once()
        self.assertTrue(any("已解除组合" in c.args[0] for c in log.call_args_list))
        self.assertIsNotNone(self.sheet.shown("A5"))

    def test_restore_follows_cells_moved_by_inserted_rows(self):
        """人工在上面插了 2 行：标色的 A5 跑到 A7，按标色时建的隐藏名称找到 A7 还原，不动现在的 A5"""
        self.add_item("A5")
        self.old_scan()
        mark = next(iter(batch._marks().values()))
        marked = batch._excel_color("#C0007A")
        moved, stays = self.sheet.cell("A7"), self.sheet.cell("A5")
        moved.Interior.Pattern, moved.Interior.Color = batch.XL_PATTERN_SOLID, marked
        moved.Font.Color = batch._excel_color("#FFFFFF")
        stays.Interior.Pattern, stays.Interior.Color, stays.Font.Color = batch.XL_PATTERN_SOLID, marked, 0
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
        self.old_scan()
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
        self.old_scan()
        mark = next(m for m in batch._marks().values() if m["addr"].replace("$", "") == "A5")
        mark.pop("name", None)
        mark["value"] = "合計"
        a5 = self.sheet.cell("A5")
        a5.Interior.Pattern, a5.Font.Color = batch.XL_PATTERN_NONE, 0
        moved = self.sheet.cell("A7")
        moved.Interior.Pattern, moved.Interior.Color = batch.XL_PATTERN_SOLID, batch._excel_color("#C0007A")
        moved.Font.Color = 0xFFFFFF
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
        self.old_scan()
        with patch.object(batch, "_restore_cells", side_effect=RuntimeError("例外が発生しました。")), \
                patch.object(batch.wf, "log_msg") as log:
            book = batch._open_without_marks(self.excel, self.path)
        self.assertIsNotNone(book)
        self.assertTrue(any("没能在检测前去掉" in c.args[0] and "S1!" in c.args[0] for c in log.call_args_list))

    def test_rule_that_cannot_be_removed_does_not_fail_the_recheck(self):
        self.add_item("A5")
        self.scan()
        self.sheet.protected = True
        with patch.object(batch.wf, "log_msg") as log:
            book = batch._open_without_marks(self.excel, self.path)
        self.assertIsNotNone(book)
        self.assertTrue(any("没能在检测前去掉" in c.args[0] and "S1" in c.args[0] for c in log.call_args_list))

    def test_old_colors_and_new_rules_in_one_file(self):
        """升级前旧方式标的颜色还在，升级后又加了规则：清除时两种一起去掉"""
        self.sheet.cells["$A$5"] = FakeCell("$A$5", pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.old_scan()
        self.add_item("B5")
        self.scan()
        self.assertEqual(len(self.sheet.rules), 1)  # A5 已有旧记录，只给 B5 加规则
        batch._sync_highlights(lambda: self.excel, self.path, clear=True)
        a5 = self.sheet.cell("A5")
        self.assertEqual((a5.Interior.Pattern, a5.Interior.Color, a5.Font.Color), (1, 0xCCFFCC, 0x0000FF))
        self.assertEqual((self.sheet.rules, batch._marks()), ([], {}))

    def test_colors_left_behind_after_a_crash_are_found_later(self):
        """复检通过后 Excel 崩溃、颜色没还原：记录还在，之后提交 / 复检时顺带还原"""
        self.add_item("A5")
        self.scan()
        self.assertEqual(batch._leftover_mark_paths([]), [])  # 黄卡还要标着
        self.card("A5").update(dismissed=True, color="blue")
        self.assertEqual(batch._leftover_mark_paths([]), [self.path])
        self.assertEqual(batch._leftover_mark_paths([self.path]), [])  # 这次任务本来就会处理

    def unhighlight(self, sync):
        self.wf.STATE["job"] = {"kind": "batch_unhighlight", "options": {"paths": [self.path]},
                                "completed": [], "pending_adjustment": None, "status": "running"}
        session = Mock()
        session.app.return_value = self.excel
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock(timed_out=False)) as dog, \
                patch.object(batch, "close_saved_workbooks", return_value={}), \
                patch.object(batch, "_sync_highlights", side_effect=sync), \
                patch.object(batch, "_restart_excel") as restart, \
                patch.object(batch, "finish_execution_batch"), \
                patch.object(batch.wf, "log_msg") as log:
            batch.run()
        return dog, restart, [c.args[0] for c in log.call_args_list]

    def test_unhighlight_says_which_file_it_is_on_and_watches_excel(self):
        dog, restart, messages = self.unhighlight(None)
        self.assertEqual(dog.call_args.kwargs["timeout"], batch.UNHIGHLIGHT_STALL_SECONDS)
        restart.assert_not_called()
        self.assertTrue(any("正在清除（1 / 1）" in m and "book.xlsx" in m for m in messages), messages)
        self.assertIn(self.path, self.wf.STATE["job"]["completed"])

    def test_clear_reports_every_cell_for_the_result_window(self):
        """清除时逐格记下结果：去掉了的、删不掉的（受保护）、旧颜色原位置已不是标记色的"""
        self.add_item("B7", row=7)
        self.old_scan()
        self.sheet.cell("B7").Interior.Color = 0x123456  # 旧方式标的颜色被人改过
        self.add_item("A5")
        self.scan()
        self.sheet.Range("E9").MergeArea.FormatConditions.Add(Type=2, Formula1='=N("CHE_MARK_0123abcd")=0')
        report = {}
        batch._sync_highlights(lambda: self.excel, self.path, clear=True, report=report)
        cells = {c["addr"]: (c["sheet"], c["status"]) for c in report["cells"]}
        self.assertEqual(cells, {"A5": ("S1", "ok"), "B7": ("S1", "check")})
        self.assertEqual(report["leftovers"], 1)
        # 删不掉的规则（工作表受保护）：记成没清掉，写原因
        self.scan()
        self.sheet.protected = True
        report = {}
        batch._sync_highlights(lambda: self.excel, self.path, clear=True, report=report)
        self.assertEqual(sorted((c["addr"], c["status"]) for c in report["cells"]), [("A5", "failed"), ("B7", "failed")])
        self.assertTrue(all("受保护" in c["reason"] for c in report["cells"]))

    def test_unhighlight_job_keeps_results_per_file_and_logs_a_total(self):
        def sync(*_, report, **__):
            report["cells"] = [{"sheet": "S1", "addr": "A5", "status": "ok", "reason": "已去掉标记"},
                               {"sheet": "S1", "addr": "B7", "status": "failed", "reason": "删不掉"}]
            report["leftovers"] = 2
        _, _, messages = self.unhighlight(sync)
        result = self.wf.STATE["job"]["result"][self.path]
        self.assertEqual((result["name"], len(result["cells"]), result["error"]), ("book.xlsx", 2, None))
        self.assertTrue(any(m == "🧽 清除完成：1 份文件，已清除 1 格，没清掉 1 格，另清掉残留标记 2 条"
                            for m in messages), messages)

    def test_unhighlight_result_explains_a_file_still_open_in_excel(self):
        self.wf.STATE["job"] = {"kind": "batch_unhighlight", "options": {"paths": [self.path]},
                                "completed": [], "pending_adjustment": None, "status": "running"}
        session = Mock()
        with patch.object(batch.autofix, "_ExcelSession", return_value=session), \
                patch.object(batch, "ExcelWatchdog", return_value=Mock(timed_out=False)), \
                patch.object(batch, "close_saved_workbooks", return_value={Path(self.path): "文件在 Excel 中有未保存的修改"}), \
                patch.object(batch, "_sync_highlights") as sync, \
                patch.object(batch, "finish_execution_batch"):
            batch.run()
        sync.assert_not_called()
        self.assertIn("未保存的修改", self.wf.STATE["job"]["result"][self.path]["error"])

    def test_unhighlight_restarts_a_stuck_or_crashed_excel_and_tries_again(self):
        _, restart, _ = self.unhighlight([RuntimeError("RPC サーバーを利用できません。"), None])
        restart.assert_called_once()
        self.assertIn(self.path, self.wf.STATE["job"]["completed"])

    def test_close_saved_workbooks_gives_up_when_your_excel_does_not_answer(self):
        from core import excel_ops
        gate = __import__("threading").Event()
        with patch.object(excel_ops, "_close_saved_workbooks", side_effect=lambda paths: gate.wait(5)):
            with self.assertRaisesRegex(ValueError, "没有响应"):
                excel_ops.close_saved_workbooks([Path(self.path)], timeout=0.1)
        gate.set()
        with patch.object(excel_ops, "_close_saved_workbooks", return_value={"x": "y"}):
            self.assertEqual(excel_ops.close_saved_workbooks([Path(self.path)]), {"x": "y"})

    def test_failed_restore_keeps_the_record_and_is_retried(self):
        """还原出错不能当成已还原、把记录删掉：那样颜色留在 Excel 里，程序再也不认得"""
        self.add_item("A5")
        self.old_scan()
        self.card("A5").update(color="blue", dismissed=True)
        with patch.object(batch, "_restore_cells", side_effect=RuntimeError("例外が発生しました。")), \
                patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        messages = [c.args[0] for c in log.call_args_list]
        self.assertEqual(len(batch._marks()), 1)
        self.assertFalse(any("还原颜色 1 格" in m for m in messages), messages)
        self.assertTrue(any("记录留着下次再试" in m for m in messages), messages)
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(batch._marks(), {})
        self.assertEqual(self.sheet.cell("A5").Interior.Pattern, batch.XL_PATTERN_NONE)
        self.assertTrue(any("还原颜色 1 格" in c.args[0] for c in log.call_args_list))

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
        self.assertEqual(self.sheet.shown("B5")[0], batch._excel_color("#C0007A"))  # 合并格整块一条规则
        self.assertEqual(len(self.sheet.rules), 1)
        self.assertIsNone(self.sheet.cell("A5").Font.Color)  # 单元格本身的字色不动
        self.card("A5").update(dismissed=True)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.rules, [])

    def test_uniform_merged_area_is_read_once_and_restored_after_shift(self):
        """方眼紙的大合并区域颜色一致：整块记一条，不一格一格读；插行挪了位置也整块还原"""
        block = [f"${c}$5" for c in "ABCDEF"]
        self.sheet.merges["$A$5"] = block
        for addr in block:
            self.sheet.cells[addr] = FakeCell(addr, pattern=1, color=0xCCFFCC, font=0x0000FF)
        self.add_item("A5")
        self.old_scan()
        mark = next(iter(batch._marks().values()))
        mark["cells"] = [{"addr": "$A$5:$F$5", "pattern": 1, "color": 0xCCFFCC, "tint": 0.0, "font": 0x0000FF}]
        marked = batch._excel_color("#C0007A")
        self.assertTrue(all(self.sheet.cell(a).Interior.Color == marked for a in block))
        # 人工在上面插了 2 行：整块挪到第 7 行
        moved = [f"${c}$7" for c in "ABCDEF"]
        for old, new in zip(block, moved):
            src = self.sheet.cell(old)
            self.sheet.cells[new] = FakeCell(new, pattern=src.Interior.Pattern, color=src.Interior.Color,
                                             font=src.Font.Color)
        self.sheet.merges["$A$7"] = moved
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
        self.old_scan()
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
        self.sheet.rules.clear()  # 没保存：文件里没有规则
        self.book.Save.side_effect = None
        with patch.object(batch.wf, "log_msg") as log:
            self.scan()
        self.assertEqual(self.sheet.shown("A5")[0], batch._excel_color("#C0007A"))
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
        self.assertGreaterEqual(tick.call_count, 120)  # 找格子、加规则各一次
        self.assertTrue(self.excel.ScreenUpdating)  # 结束后恢复

    def test_unconfirmed_location_is_not_colored(self):
        self.add_item("A5", location_confirmed=False)
        self.scan()
        self.assertIn("未在 Excel 中标色", self.card("A5")["message"])
        self.excel.Workbooks.Open.assert_not_called()

    def test_clear_restores_everything(self):
        mine = self.user_rule("A5")
        self.add_item("A5")
        self.add_item("B5")
        self.scan()
        # 上次崩溃、记录没了，规则还留在文件里：按前缀一起清掉
        self.sheet.Range("E9").MergeArea.FormatConditions.Add(Type=2, Formula1='=N("CHE_MARK_0123abcd")=0')
        with patch.object(batch.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path, clear=True)
        self.assertEqual(self.sheet.rules, [mine])
        self.assertEqual(batch._marks(), {})
        self.assertTrue(any("去掉标记 2 格" in c.args[0] and "残留标记 1 条" in c.args[0]
                            for c in log.call_args_list))

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
        self.book.Worksheets = list(sheets)

    def test_deleted_sheet_skips_restore_without_warning(self):
        self.sheet.Name = "S1"
        self.use_named_sheets([self.sheet])
        self.add_item("A5")
        self.old_scan()
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

    def test_deleted_sheet_takes_its_rules_with_it(self):
        self.use_named_sheets([self.sheet])
        self.add_item("A5")
        self.scan()
        self.use_named_sheets([])
        self.card("A5").update(color="blue", verified=True)
        with patch.object(self.wf, "log_msg") as log:
            batch._sync_highlights(lambda: self.excel, self.path)
        self.assertFalse(any("⚠️" in c.args[0] for c in log.call_args_list))
        self.assertEqual(batch._marks(), {})

    def test_sheet_name_with_extra_spaces_is_still_found(self):
        self.sheet.Name = "S1\u3000"
        self.use_named_sheets([self.sheet])
        self.add_item("A5")
        self.scan()
        self.assertEqual(self.sheet.shown("A5")[0], batch._excel_color("#C0007A"))

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
        self.assertEqual(self.sheet.shown("A5")[0], batch._excel_color("#6D28D9"))
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
        self.assertEqual(self.sheet.shown("A5")[0], batch._excel_color("#006633"))

    def test_full_auto_without_highlight_stops_coloring(self):
        self.add_item("A5")
        self.scan()
        batch._data()["highlight_files"].pop(self.path)
        batch._sync_highlights(lambda: self.excel, self.path)
        self.assertEqual(self.sheet.rules, [])


if __name__ == "__main__":
    unittest.main()
