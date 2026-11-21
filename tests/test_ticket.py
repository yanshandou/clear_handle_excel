import stat
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import icon_library, ticket
from core.pdf_detector import fitz


class TicketTests(unittest.TestCase):
    def setUp(self):
        desktop = tempfile.TemporaryDirectory()
        self.addCleanup(desktop.cleanup)
        self.desktop = Path(desktop.name).resolve()
        desktop_lookup = patch.object(ticket, "_desktop", return_value=self.desktop)
        desktop_lookup.start()
        self.addCleanup(desktop_lookup.stop)

    def add_icon(self):
        library = tempfile.TemporaryDirectory()
        self.addCleanup(library.cleanup)
        library_path = patch.object(icon_library, "LIB_DIR", Path(library.name))
        library_path.start()
        self.addCleanup(library_path.stop)
        image = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 8, 8), 0)
        image.clear_with(120)
        return icon_library.add("传票.png", image.tobytes("png"))["icon"]

    def test_saves_confirmed_fields_as_readable_text(self):
        task = ticket.create({
            "ticket_id": "＃５５３４５", "title": "TBSNG0000_手書きサイン/?",
            "project": "開発（UI）", "parent_project": "上级项目", "parent": "55022",
            "parent_category": "機能追加", "parent_title": "案件标题", "assignee": "白家绫", "assignee_short": "白", "progress": "0%", "estimated_hours": "9.00",
        })
        folder = Path(task["folder"])
        self.assertEqual(folder, self.desktop / "Ticket #55345 - TBSNG0000_手書きサイン__")
        self.assertEqual((folder / task["filename"]).read_text(encoding="utf-8-sig"),
                         "Ticket #55345\n\n任务标题：TBSNG0000_手書きサイン/?\n所属项目：開発（UI）\n上级项目名称：上级项目\n"
                         "关联案件编号：55022\n关联案件分类：機能追加\n关联案件标题：案件标题\n"
                         "负责人：白家绫\n负责人简称：白\n进度：0%\n预计工数（人日）：9.00\n")

    def test_retries_and_preserves_different_tickets_with_the_same_name(self):
        first = self.desktop / "Ticket - 手動"
        first.mkdir()
        (first / "Ticket.txt").write_text("existing", encoding="utf-8")
        fields = {"title": "手動", "status": "進行中"}
        saved = ticket.create(fields)
        self.assertEqual(Path(saved["folder"]), self.desktop / "Ticket - 手動 (2)")
        self.assertEqual(ticket.create(fields), saved)
        self.assertEqual((first / "Ticket.txt").read_text(encoding="utf-8"), "existing")

    def test_invalid_values_create_no_folder(self):
        for fields in ({"title": " "}, {"title": "Task", "ticket_id": "invalid"}, {"title": "two\nlines"}):
            with self.subTest(fields=fields), self.assertRaises(ValueError):
                ticket.create(fields)
        self.assertEqual(list(self.desktop.iterdir()), [])

    def test_failed_write_removes_incomplete_folder(self):
        with patch.object(Path, "open", side_effect=OSError("disk error")), self.assertRaises(OSError):
            ticket.create({"title": "Task"})
        self.assertEqual(list(self.desktop.iterdir()), [])

    def test_applies_library_icon_to_the_ticket_folder(self):
        icon = self.add_icon()
        fields = {"title": "Task", "icon_id": icon["id"]}
        saved = ticket.create(fields)
        self.assertEqual(saved["icon_id"], icon["id"])
        folder = Path(saved["folder"])
        icon_copy = folder / f"{icon_library.ICON_PREFIX}{icon['id']}_{icon['version']}.ico"
        ini = folder / "desktop.ini"
        self.assertEqual(icon_copy.read_bytes(), icon_library.icon_file(icon["id"], "ico")[0])
        self.assertIn(f"IconResource={icon_copy.name},0", ini.read_text())
        self.assertNotIn("icon_id", (folder / "Ticket.txt").read_text(encoding="utf-8-sig"))
        self.assertEqual(ticket.create(fields), saved)
        self.assertEqual(icon_library.listing()["icons"][0]["folders"], [str(folder)])
        if sys.platform == "win32":
            hidden_system = stat.FILE_ATTRIBUTE_HIDDEN | stat.FILE_ATTRIBUTE_SYSTEM
            for file in (ini, icon_copy):
                self.assertEqual(file.stat().st_file_attributes & hidden_system, hidden_system)
            self.assertTrue(folder.stat().st_file_attributes & (stat.FILE_ATTRIBUTE_READONLY | stat.FILE_ATTRIBUTE_SYSTEM))

    def test_rejects_an_invalid_icon_before_creating_the_folder(self):
        with self.assertRaisesRegex(ValueError, "图标编号"):
            ticket.create({"title": "Task", "icon_id": "../icon.ico"})
        self.assertEqual(list(self.desktop.iterdir()), [])

    def test_reports_icon_failure_and_can_retry_without_duplicating_the_ticket(self):
        icon = self.add_icon()
        fields = {"title": "Task", "icon_id": icon["id"]}
        with patch.object(icon_library, "_apply_one", side_effect=OSError("icon write failed")):
            saved = ticket.create(fields)
        self.assertEqual(saved["icon_error"], "icon write failed")
        self.assertNotIn("icon_id", saved)
        self.assertTrue((Path(saved["folder"]) / "Ticket.txt").is_file())
        retried = ticket.create(fields)
        self.assertNotIn("icon_error", retried)
        self.assertEqual(retried["folder"], saved["folder"])
        self.assertEqual(len(list(self.desktop.iterdir())), 1)

    def test_open_reports_when_explorer_is_opened_without_foreground_focus(self):
        folder = Path(ticket.create({"title": "Task"})["folder"])
        with patch("os.startfile", create=True) as launch, patch.object(ticket, "_show_explorer_folder", return_value=False) as show:
            self.assertEqual(ticket.open_folder({"folder": str(folder)}), {"foreground": False})
        launch.assert_called_once_with(folder)
        show.assert_called_once_with(folder)


if __name__ == "__main__":
    unittest.main()
