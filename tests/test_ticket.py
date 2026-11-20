import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import ticket


class TicketTests(unittest.TestCase):
    def setUp(self):
        desktop = tempfile.TemporaryDirectory()
        self.addCleanup(desktop.cleanup)
        self.desktop = Path(desktop.name).resolve()
        desktop_lookup = patch.object(ticket, "_desktop", return_value=self.desktop)
        desktop_lookup.start()
        self.addCleanup(desktop_lookup.stop)

    def test_saves_confirmed_fields_as_readable_text(self):
        task = ticket.create({
            "ticket_id": "＃５５３４５", "title": "TBSNG0000_手書きサイン/?",
            "parent": "機能追加 - 案件 #55022", "progress": "0%", "estimated_hours": "9.00",
        })
        folder = Path(task["folder"])
        self.assertEqual(folder, self.desktop / "Ticket #55345 - TBSNG0000_手書きサイン__")
        self.assertEqual((folder / task["filename"]).read_text(encoding="utf-8-sig"),
                         "Ticket #55345\n\n任务标题：TBSNG0000_手書きサイン/?\n关联案件：機能追加 - 案件 #55022\n"
                         "进度：0%\n预计工时：9.00\n")

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


if __name__ == "__main__":
    unittest.main()
