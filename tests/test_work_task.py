import tempfile
import time
import unittest
from pathlib import Path, PureWindowsPath
from unittest.mock import patch

from core import distribute, excel_diff, storage, upload_prep, work_task, workflow as wf


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
        # 上传前的清标记 / 恢复颜色 / 视图标准化要开 Excel，这里只测上传本身（见 test_upload_prep.py）
        prep = patch.object(upload_prep, "prepare", return_value={"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0,
                                                                  "color_problems": [], "reference": True, "normalized": True})
        prep.start()
        self.addCleanup(prep.stop)
        distribute._STATUS.clear()
        distribute._STATUS["running"] = False
        wf.STATE["busy"] = False

    def tearDown(self):
        self.db.stop()
        self.tmp.cleanup()

    def wait(self):
        distribute.start()
        return self.wait_applied()

    def wait_applied(self):
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        self.assertFalse(distribute.status().get("running"))
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
        self.assertEqual(backup.parent.name, "处理")
        # 参考被覆盖前也备份（用户 2026-11-06 定）：旧参考 AAA 在同一次备份的「参考」子文件夹
        a_item = next(f for f in fetched["folders"][0]["files"] if f["rel"] == "a.xlsx")
        ref_backup = Path(a_item["ref_backup_path"])
        self.assertEqual(ref_backup.read_text(encoding="utf-8"), "AAA")
        self.assertEqual(ref_backup.parent, backup.parent.parent / "参考")
        self.assertEqual(fetched["summary"]["备份参考"], 2)  # b 共享已是我们传的 b2，参考从 b1 更新，旧的也备份

    def test_multiple_shared_paths_fetch_only_picked(self):
        """新任务填多行：共享文件夹取共同上级，只拿填了的文件夹 / 文件（用户 2026-11-12）"""
        base = self.shared.parent
        write(base / "13.テーブル定義書" / "t.xlsx", "T")
        write(base / "13.テーブル定義書" / "旧" / "t_old.xlsx", "O")
        write(base / "14.その他" / "skip.xlsx", "S")
        write(base / "15.単票" / "one.xlsx", "1")
        write(base / "15.単票" / "two.xlsx", "2")
        text = "\n".join([f'"{self.shared}"', str(base / "13.テーブル定義書"), "",
                          str(base / "15.単票" / "one.xlsx"), str(self.shared / "子")])  # 子 在 12 里面，只留 12
        paths = work_task.derive(text, str(self.work_root))
        self.assertEqual(Path(paths["shared"]), base.resolve())
        self.assertEqual(paths["picks"], ["12.ビュー定義書", "13.テーブル定義書", "15.単票/one.xlsx"])
        self.assertEqual(Path(paths["work"]), self.work_root / "处理" / "介舟ファミリー" / "20基本設計")
        task = work_task.create(text, str(self.work_root))
        prev = distribute.preview_fetch("", "", "", task["id"])
        rels = sorted(f["rel"] for f in prev["folders"][0]["files"])
        self.assertEqual(rels, ["12.ビュー定義書/a.xlsx", "12.ビュー定義書/子/孫/b.xlsx", "13.テーブル定義書/t.xlsx",
                                "13.テーブル定義書/旧/t_old.xlsx", "15.単票/one.xlsx"])
        self.wait()
        work = Path(task["work"])
        self.assertTrue((work / "15.単票" / "one.xlsx").exists())
        self.assertFalse((work / "15.単票" / "two.xlsx").exists())
        self.assertFalse((work / "14.その他").exists())
        self.assertEqual(work_task.finish_check(task["id"])["shared_files"], 5)

    def test_single_shared_path_unchanged_and_bad_multi_paths(self):
        self.assertEqual(work_task.derive(str(self.shared), str(self.work_root))["picks"], [])
        one = work_task.derive(str(self.shared / "a.xlsx"), str(self.work_root))  # 只填一个文件：拿这一份
        self.assertEqual((Path(one["shared"]), one["picks"]), (self.shared.resolve(), ["a.xlsx"]))
        whole = work_task.derive(f"{self.shared}\n{self.shared / 'a.xlsx'}", str(self.work_root))
        self.assertEqual(whole["picks"], [])  # 填了文件夹本身 = 整个都拿
        with self.assertRaisesRegex(ValueError, "不是 Excel"):
            work_task.derive(f"{self.shared}\n{self.shared / '说明.pdf'}", str(self.work_root))
        with self.assertRaisesRegex(ValueError, "找不到"):
            work_task.derive(f"{self.shared}\n{self.shared / '没有'}", str(self.work_root))
        other = Path(self.tmp.name) / "Users" / "s.go" / "OneDrive - 会社" / "別案件"
        other.mkdir(parents=True)
        with self.assertRaisesRegex(ValueError, "太靠上"):
            work_task.derive(f"{self.shared}\n{other}", str(self.work_root))

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

    def test_task_paths_are_defaults_and_all_local_folders_remain_editable(self):
        """不需要新建或切换任务，各功能可以直接选择本地路径。"""
        self.assertEqual(work_task.require_folder(str(self.work_root)), str(self.work_root))
        with self.assertRaisesRegex(ValueError, "OneDrive"):
            work_task.require_folder(str(self.shared))
        first = work_task.create(str(self.shared), str(self.work_root))
        sub = Path(first["work"]) / "子"
        sub.mkdir()
        self.assertEqual(work_task.require_folder(first["work"]), first["work"])
        self.assertEqual(work_task.require_folder(str(sub)), str(sub))  # 子文件夹可以
        self.assertEqual(work_task.require_folder(first["reference"], "reference"), first["reference"])
        for local, role in ((first["reference"], "work"), (first["work"], "reference"), (str(self.work_root), "work")):
            self.assertEqual(work_task.require_folder(local, role), local)
        for bad, role in ((str(self.shared), "work"), (str(self.shared / "子"), "reference")):
            with self.assertRaises(ValueError, msg=(bad, role)):
                work_task.require_folder(bad, role)
        other = self.shared.parent / "13.テーブル定義書"
        other.mkdir()
        second = work_task.create(str(other), str(self.work_root))
        self.assertEqual(work_task.require_folder(first["work"]), first["work"])
        work_task.require_folder(second["work"])
        for bad in ("", "  ", None, 123):
            with self.assertRaises(ValueError):
                work_task.require_folder(bad)

    def test_single_mode_allows_local_folders_but_not_onedrive(self):
        """旧路径偏好开关只影响默认值，不再限制可选路径。"""
        local = self.work_root / "随便一个文件夹"
        local.mkdir(parents=True)
        self.assertEqual(work_task.require_folder(str(local)), str(local))
        self.assertTrue(work_task.set_single(True)["single"])
        self.assertEqual(work_task.require_folder(str(local)), str(local))
        with self.assertRaisesRegex(ValueError, "OneDrive"):
            work_task.require_folder(str(self.shared))
        task = work_task.create(str(self.shared), str(self.work_root))
        self.assertFalse(work_task.overview()["single"])
        self.assertEqual(work_task.require_folder(str(local)), str(local))
        work_task.set_single(True)
        work_task.switch(task["id"])
        self.assertFalse(work_task.overview()["single"])
        self.assertFalse(work_task.set_single(False)["single"])

    def test_normalize_api_uses_custom_path_while_another_task_is_selected(self):
        import web_clip_fixer as server
        work_task.create(str(self.shared), str(self.work_root))
        chosen = self.work_root / "自选目录"
        write(chosen / "custom.xlsx")
        handler = object.__new__(server.WebHandler)
        handler.path = "/api/quick_task"
        handler._start_job = lambda kind, options: setattr(handler, "started", (kind, options))
        handler._handle_post({"kind": "normalize_view", "folder": str(chosen)})
        self.assertEqual(handler.started, ("normalize_view", {"folder": str(chosen.resolve())}))

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

    def wait_finish(self):
        for _ in range(200):
            status = work_task.finish_status()
            if not status.get("running"):
                return status
            time.sleep(0.05)
        self.fail("收尾检查没结束")

    def test_finish_page_runs_in_background_with_steps(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        work = Path(task["work"])
        distribute.preview_fetch("", "", "", task["id"])
        self.wait()
        write(work / "a.xlsx", "BBB")
        write(work / "新.xlsx", "new")
        work_task.start_finish(task["id"])
        status = self.wait_finish()
        self.assertEqual(status["task"]["id"], task["id"])
        self.assertEqual(status["early"], {"共享里没有": ["新.xlsx"], "待重新处理": []})
        self.assertEqual((status["done"], status["total"]), (2, 2))  # 只比对共享里有的两个
        result = status["result"]
        self.assertFalse(result["passed"])
        self.assertEqual(result["shared_files"], 2)
        states = {row["rel"]: row["state"] for row in result["file_states"]}
        self.assertEqual(states, {"a.xlsx": "没上传", "子/孫/b.xlsx": "ok", "新.xlsx": "共享里没有"})
        self.assertFalse(work_task.get(task["id"]).get("done"))
        (work / "新.xlsx").unlink()
        write(work / "a.xlsx", "AAA")
        work_task.start_finish(task["id"])
        status = self.wait_finish()
        self.assertTrue(status["result"]["passed"])
        self.assertTrue(status["done_task"])
        self.assertTrue(work_task.get(task["id"])["done"])

    def test_selected_upload_only_hashes_selected_shared_files(self):
        write(self.shared / "unselected.xlsx", "original")
        task = work_task.create(str(self.shared), str(self.work_root))
        work = Path(task["work"])
        task_id = task["id"]
        distribute.preview_fetch("", "", "", task_id)
        self.wait()
        write(work / "a.xlsx", "my changes")
        write(work / "子/孫/b.xlsx", "b2")
        write(self.shared / "a.xlsx", "shared v2")
        write(self.shared / "unselected.xlsx", "someone updated this too")
        digest = distribute._digest
        hashed = set()

        def record_hash(path):
            if path.is_relative_to(self.shared) and not path.name.endswith(distribute.TEMP_SUFFIX):
                hashed.add(path.relative_to(self.shared).as_posix())
            return digest(path)

        batch = ["a.xlsx", "子/孫/b.xlsx"]
        with (patch.object(distribute, "_digest", side_effect=record_hash),
              patch.object(excel_diff, "diff", return_value={"identical": False})):
            self.assertEqual(len(distribute.task_files(task_id)["files"]), 3)
            self.assertEqual(hashed, set())
            distribute.start_task_upload(task_id, batch, {})
            for _ in range(200):
                reviews = distribute.status().get("reviews")
                if reviews and (self.shared / batch[1]).read_text(encoding="utf-8") == "b2":
                    break
                time.sleep(0.05)
            self.assertEqual([r["rel"] for r in reviews], ["a.xlsx"])
            self.assertEqual((self.shared / "a.xlsx").read_text(encoding="utf-8"), "shared v2")
            self.assertEqual((work / "a.xlsx").read_text(encoding="utf-8"), "my changes")
            distribute.decide_upload_review(reviews[0]["token"], "upload")
            result = self.wait_applied()
            self.assertEqual(hashed, set(batch))
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual((self.shared / "a.xlsx").read_text(encoding="utf-8"), "my changes")
        self.assertEqual((work / "unselected.xlsx").read_text(encoding="utf-8"), "original")
        self.assertIn("a.xlsx | 已上传", distribute.markdown(distribute.get_run(result["id"])))
        uploaded = distribute.task_files(task_id)["files"][0]
        self.assertTrue(uploaded["uploaded_current"])
        self.assertTrue(uploaded["uploaded_at"])
        self.assertIsNone(uploaded["review"])
        write(work / "a.xlsx", "edited again")
        self.assertFalse(distribute.task_files(task_id)["files"][0]["uploaded_current"])

    def test_focus_lists_only_review_files_until_they_are_uploaded(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        work, ref, task_id = Path(task["work"]), Path(task["reference"]), task["id"]
        distribute.preview_fetch("", "", "", task_id)
        self.wait()
        write(work / "c.xlsx", "local only")
        _, all_books = wf.list_workbooks(str(work))
        self.assertEqual(len(all_books), 3)
        with self.assertRaises(ValueError):
            work_task.set_focus(task_id, ["不存在.xlsx"])
        work_task.set_focus(task_id, ["子/孫/b.xlsx", "a.xlsx", "子/孫/b.xlsx"])
        self.assertEqual(work_task.focus_files(work_task.get(task_id)), ["子/孫/b.xlsx", "a.xlsx"])
        # 处理、参考、子文件夹都只列清单里的；文件名前面的 ★ 不影响
        (work / "a.xlsx").rename(work / "★a.xlsx")
        _, books = wf.list_workbooks(str(work))
        self.assertEqual(sorted(p.name for p in books), ["b.xlsx", "★a.xlsx"])
        _, books = wf.list_workbooks(str(ref))
        self.assertEqual(sorted(p.name for p in books), ["a.xlsx", "b.xlsx"])
        _, books = wf.list_workbooks(str(work / "子"))
        self.assertEqual([p.name for p in books], ["b.xlsx"])
        (work / "★a.xlsx").rename(work / "a.xlsx")
        # 别的文件夹不受影响
        other = Path(self.tmp.name) / "other"
        write(other / "a.xlsx")
        write(other / "z.xlsx")
        self.assertEqual(len(wf.list_workbooks(str(other))[1]), 2)
        # 只传了 a：a 从清单划掉，b 还在
        write(work / "a.xlsx", "fixed a")
        distribute.start_task_upload(task_id, ["a.xlsx"], {})
        self.assertEqual(self.wait_applied()["summary"]["已复制"], 1)
        self.assertEqual(work_task.focus_files(work_task.get(task_id)), ["子/孫/b.xlsx"])
        self.assertEqual([p.name for p in wf.list_workbooks(str(work))[1]], ["b.xlsx"])
        distribute.start_task_upload(task_id, ["子/孫/b.xlsx"], {})
        self.assertEqual(self.wait_applied()["summary"]["已复制"], 1)
        self.assertEqual(work_task.focus_files(work_task.get(task_id)), [])
        self.assertEqual(len(wf.list_workbooks(str(work))[1]), 3)  # 自动回到整个文件夹

    def test_focus_exit_and_missing_files(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        work, task_id = Path(task["work"]), task["id"]
        distribute.preview_fetch("", "", "", task_id)
        self.wait()
        work_task.set_focus(task_id, ["a.xlsx"])
        (work / "a.xlsx").unlink()
        with self.assertRaisesRegex(ValueError, "退出，处理整个文件夹"):
            wf.list_workbooks(str(work))
        work_task.clear_focus(task_id)
        self.assertEqual(len(wf.list_workbooks(str(work))[1]), 1)
        # 同一个共享文件夹重新开任务 = 重新开始，清单也清掉
        write(work / "a.xlsx")
        work_task.set_focus(task_id, ["a.xlsx"])
        work_task.create(str(self.shared), str(self.work_root))
        self.assertEqual(work_task.focus_files(work_task.get(task_id)), [])

    def test_upload_rechecks_remote_immediately_before_replacing(self):
        task = work_task.create(str(self.shared), str(self.work_root))
        task_id, work, ref = task["id"], Path(task["work"]), Path(task["reference"])
        distribute.preview_fetch("", "", "", task_id)
        self.wait()
        write(work / "a.xlsx", "local edits")
        target = self.shared / "a.xlsx"
        original_copy = distribute.copy_one

        def copy_with_race(src, dst, **kwargs):
            if dst == target:
                verify = kwargs["before_replace"]

                def changed():
                    write(target, "arrived during staging")
                    verify()

                kwargs["before_replace"] = changed
            return original_copy(src, dst, **kwargs)

        with patch.object(distribute, "copy_one", side_effect=copy_with_race):
            distribute.start_task_upload(task_id, ["a.xlsx"], {})
            self.assertEqual(self.wait_applied()["summary"]["没复制成功"], 1)
        self.assertEqual(target.read_text(encoding="utf-8"), "arrived during staging")
        self.assertEqual((work / "a.xlsx").read_text(encoding="utf-8"), "local edits")
        self.assertEqual((ref / "a.xlsx").read_text(encoding="utf-8"), "AAA")
        self.assertIsNone(distribute.task_files(task_id)["files"][0]["review"])


if __name__ == "__main__":
    unittest.main()
