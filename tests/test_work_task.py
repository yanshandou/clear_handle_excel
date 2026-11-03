import tempfile
import time
import unittest
from pathlib import Path, PureWindowsPath
from unittest.mock import patch

from core import distribute, storage, work_task, workflow as wf


def write(path: Path, text: str = "x") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


class WorkTaskTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.shared = root / "Users" / "s.go" / "OneDrive - 会社" / "介舟ファミリー" / "20基本設計" / "12.ビュー定義書"
        self.work_root = root / "clear_handle_excel"
        write(self.shared / "a.xlsx", "AAA")
        write(self.shared / "子" / "孫" / "b.xlsx", "b1")
        write(self.shared / "说明.pdf", "pdf")
        self.db = patch.object(storage, "DATABASE_PATH", root / "test.sqlite3")
        self.db.start()
        distribute._STATUS.clear()
        distribute._STATUS["running"] = False
        wf.STATE["busy"] = False

    def tearDown(self):
        self.db.stop()
        self.tmp.cleanup()

    def wait(self):
        distribute.start()
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def test_relative_part_after_onedrive(self):
        p = PureWindowsPath(r"D:\Users\s.go\OneDrive - 日本コンピュータコンサルタント\介舟ファミリー\20基本設計\10DB\12.ビュー定義書")
        self.assertEqual(work_task.relative_part(p), ["介舟ファミリー", "20基本設計", "10DB", "12.ビュー定義書"])
        self.assertEqual(work_task.relative_part(PureWindowsPath(r"E:\共有\仕様書")), ["仕様書"])
        self.assertEqual(work_task.default_work_root(), Path.home() / "Desktop" / "handle-excel")

    def test_derive_paths_and_rejects_onedrive_root(self):
        paths = work_task.derive(str(self.shared), str(self.work_root))
        self.assertEqual(paths["rel"], "介舟ファミリー/20基本設計/12.ビュー定義書")
        self.assertEqual(Path(paths["reference"]), self.work_root / "参考" / "介舟ファミリー" / "20基本設計" / "12.ビュー定義書")
        self.assertEqual(Path(paths["work"]), self.work_root / "处理" / "介舟ファミリー" / "20基本設計" / "12.ビュー定義書")
        with self.assertRaises(ValueError):
            work_task.derive(str(self.shared), str(self.shared.parent / "x"))  # 在 OneDrive 里
        with self.assertRaises(ValueError):
            work_task.derive(str(self.shared / "不存在"), str(self.work_root))

    def test_full_task_flow_with_subfolders(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        work, ref = Path(task["work"]), Path(task["reference"])
        prev = distribute.preview_fetch("", "", "", task["id"])
        self.assertEqual(prev["summary"]["要拿"], 2)
        self.wait()
        self.assertEqual((work / "子" / "孫" / "b.xlsx").read_text(encoding="utf-8"), "b1")
        self.assertEqual((ref / "a.xlsx").read_text(encoding="utf-8"), "AAA")
        self.assertFalse((work / "说明.pdf").exists())  # 只拿 Excel
        # 改 a 和 b；别人同时改了共享里的 a
        write(work / "a.xlsx", "BBB")
        write(work / "子" / "孫" / "b.xlsx", "b2")
        write(self.shared / "a.xlsx", "CCC")
        prev = distribute.preview("", "", "", task["id"])
        self.assertEqual(prev["summary"]["别人改过"], 1)
        result = self.wait()
        self.assertEqual((self.shared / "a.xlsx").read_text(encoding="utf-8"), "CCC")
        self.assertEqual((self.shared / "子" / "孫" / "b.xlsx").read_text(encoding="utf-8"), "b2")
        self.assertEqual(result["summary"]["别人改过没覆盖"], 1)
        # 再拿最新版：a 换成 CCC（BBB 备份到工作根目录），参考也更新成 CCC；b 是我们传上去的，参考不动
        distribute.preview_fetch("", "", "", task["id"])
        fetched = self.wait()
        self.assertEqual((work / "a.xlsx").read_text(encoding="utf-8"), "CCC")
        self.assertEqual((ref / "a.xlsx").read_text(encoding="utf-8"), "CCC")
        backup = Path(next(f for f in fetched["folders"][0]["files"] if f["rel"] == "a.xlsx")["backup_path"])
        self.assertEqual(backup.read_text(encoding="utf-8"), "BBB")
        self.assertIn(str(self.work_root / "拿最新版前的备份"), str(backup))

    def test_tasks_are_kept_and_switchable(self):
        first = work_task.create(str(self.shared), str(self.work_root))
        other = self.shared.parent / "13.テーブル定義書"
        other.mkdir()
        second = work_task.create(str(other), str(self.work_root))
        data = work_task.overview()
        self.assertEqual([t["id"] for t in data["tasks"]], [second["id"], first["id"]])
        self.assertEqual(data["current"], second["id"])
        work_task.switch(first["id"])
        self.assertEqual(work_task.overview()["current"], first["id"])
        again = work_task.create(str(self.shared), str(self.work_root))
        self.assertEqual(again["id"], first["id"])  # 同一个共享文件夹沿用以前的任务
        self.assertEqual(len(work_task.overview()["tasks"]), 2)

    def test_steps_only_accept_current_task_folders(self):
        """用户 2026-11-03：不走新任务不能直接填路径；OneDrive（共享）里的文件夹一律不让改"""
        with self.assertRaisesRegex(ValueError, "新任务"):
            work_task.require_folder(str(self.work_root))  # 还没有任务
        with self.assertRaisesRegex(ValueError, "OneDrive"):
            work_task.require_folder(str(self.shared))
        first = work_task.create(str(self.shared), str(self.work_root))
        sub = Path(first["work"]) / "子"
        sub.mkdir()
        self.assertEqual(work_task.require_folder(first["work"]), first["work"])
        self.assertEqual(work_task.require_folder(str(sub)), str(sub))  # 子文件夹可以
        self.assertEqual(work_task.require_folder(first["reference"], "reference"), first["reference"])
        for bad, role in ((first["reference"], "work"), (first["work"], "reference"), (str(self.work_root), "work"),
                          (str(self.shared), "work"), (str(self.shared / "子"), "reference")):
            with self.assertRaises(ValueError, msg=(bad, role)):
                work_task.require_folder(bad, role)
        other = self.shared.parent / "13.テーブル定義書"
        other.mkdir()
        second = work_task.create(str(other), str(self.work_root))
        with self.assertRaisesRegex(ValueError, "13.テーブル定義書"):
            work_task.require_folder(first["work"])  # 换了任务，旧任务的处理文件夹不能用
        work_task.require_folder(second["work"])
        with patch.object(work_task, "ENFORCE", False):
            work_task.require_folder(str(self.work_root))
            with self.assertRaises(ValueError):
                work_task.require_folder(str(self.shared))  # 关了任务检查，OneDrive 也照样拦

    def test_finish_check_blocks_until_everything_uploaded(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        work = Path(task["work"])
        distribute.preview_fetch("", "", "", task["id"])
        self.wait()
        write(work / "a.xlsx", "BBB")
        write(work / "新.xlsx", "new")
        result = work_task.finish(task["id"])
        self.assertFalse(result["done"])
        self.assertEqual(result["issues"]["没上传"], ["a.xlsx"])
        self.assertEqual(result["issues"]["共享里没有"], ["新.xlsx"])
        write(self.shared / "子" / "孫" / "b.xlsx", "别人改的")
        self.assertEqual(work_task.finish_check(task["id"])["issues"]["别人改过"], ["子/孫/b.xlsx"])
        distribute.preview_fetch("", "", "", task["id"])
        self.wait()
        distribute.preview("", "", "", task["id"])
        self.wait()
        result = work_task.finish(task["id"])
        self.assertTrue(result["done"], result["issues"])
        self.assertTrue(work_task.get(task["id"])["done"])
        self.assertFalse(work_task.create(str(self.shared), str(self.work_root))["done"])  # 再开一次重新开始


if __name__ == "__main__":
    unittest.main()
