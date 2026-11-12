"""几个账号分着做（用户 2026-11-07）：张三上传到团队中转并写拿取版本清单，李四从中转「起点 → 终点」传到公司"""
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import distribute, excel_diff, storage, upload_prep, work_task, workflow as wf


def write(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


PREP = {"marks": 0, "tabs": 0, "colors": 0, "color_skipped": 0, "color_problems": [], "reference": False, "normalized": True}


class RelayUploadTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.company = root / "公司" / "仕様書"
        self.team = root / "团队" / "AAA"
        self.team.mkdir(parents=True)
        write(self.company / "a.xlsx", "A")
        write(self.company / "sub" / "b.xlsx", "B")
        for obj, name, value in ((storage, "DATABASE_PATH", root / "test.sqlite3"),
                                 (distribute, "_STATUS", {"running": False})):
            patcher = patch.object(obj, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        wf.STATE["busy"] = False
        self.task = work_task.create(str(self.company), str(root / "handle-excel"))
        distribute.preview_fetch("", "", "", self.task["id"])
        distribute.start()
        self.wait()
        self.work = Path(self.task["work"])
        write(self.work / "a.xlsx", "A fixed")
        write(self.work / "sub" / "b.xlsx", "B fixed")

    def wait(self):
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def patches(self, prepare=None):
        seen = []

        def default(get_excel, local, reference, where):
            seen.append((local, reference))
            local.write_text(local.read_text(encoding="utf-8") + " +标准化", encoding="utf-8")
            return dict(PREP)

        stack = [patch.object(upload_prep, "prepare", side_effect=prepare or default),
                 patch("core.excel_ops.start_background_excel", Mock()), patch("core.excel_ops.quit_background_excel", Mock()),
                 patch("core.file_problem.ask", Mock(return_value="skip"))]
        for p in stack:
            p.start()
            self.addCleanup(p.stop)
        return seen

    def relay_upload(self):
        work_task.set_upload_to(self.task["id"], str(self.team))
        self.patches(lambda get_excel, local, reference, where: dict(PREP))
        distribute.start_task_upload(self.task["id"], ["a.xlsx", "sub/b.xlsx"], {})
        return self.wait()

    def test_task_upload_to_relay_keeps_company_and_writes_manifest(self):
        result = self.relay_upload()
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual(read(self.team / "a.xlsx"), "A fixed")
        self.assertEqual(read(self.team / "sub" / "b.xlsx"), "B fixed")
        self.assertEqual(read(self.company / "a.xlsx"), "A")  # 公司那份不动
        manifest = json.loads(read(self.team / distribute.MANIFEST_NAME))
        self.assertEqual(Path(manifest["shared"]), self.company.resolve())
        self.assertEqual(manifest["files"]["a.xlsx"]["shared_sha256"], distribute._digest(self.company / "a.xlsx"))
        self.assertEqual(sorted(manifest["shared_files"]), ["a.xlsx", "sub/b.xlsx"])
        self.assertIn("上传到（中转）", distribute.markdown(result))
        files = {f["rel"]: f for f in distribute.task_files(self.task["id"])["files"]}
        self.assertTrue(files["a.xlsx"]["uploaded_current"])
        self.assertTrue(work_task._check(work_task.get(self.task["id"]))["passed"])

    def test_upload_to_rejects_folders_inside_the_task(self):
        with self.assertRaises(ValueError):
            work_task.set_upload_to(self.task["id"], str(self.work))
        with self.assertRaises(ValueError):
            work_task.set_upload_to(self.task["id"], str(self.company / "sub"))
        self.assertEqual(work_task.set_upload_to(self.task["id"], "")["upload_to"], "")

    def review(self, rel, decision):
        for _ in range(200):
            pending = next((r for r in distribute.status().get("reviews", []) if r["rel"] == rel), None)
            if pending:
                distribute.upload_review(pending["token"])
                distribute.decide_upload_review(pending["token"], decision)
                return
            time.sleep(0.05)
        self.fail(f"{rel} 没有进入逐文件确认")

    def test_free_upload_from_relay_uses_manifest_and_reviews_changed_remote(self):
        self.relay_upload()
        write(self.company / "sub" / "b.xlsx", "B 别人改了")
        write(self.company / "c.xlsx", "C 别人新加")
        preview = distribute.preview_free(str(self.team), str(self.company))
        actions = {f["rel"]: f["action"] for f in preview["files"]}
        self.assertEqual(actions, {"a.xlsx": "overwrite", "sub/b.xlsx": "review"})
        self.assertEqual(preview["added"], ["c.xlsx"])
        seen = self.patches()
        with patch.object(excel_diff, "diff", return_value={"identical": False}) as compare:
            distribute.start_free(preview["id"])
            self.review("sub/b.xlsx", "skip")
            result = self.wait()
        compare.assert_called_once_with(self.team / "sub" / "b.xlsx", self.company / "sub" / "b.xlsx")
        self.assertEqual(result["summary"]["已复制"], 1)
        self.assertEqual(read(self.company / "a.xlsx"), "A fixed +标准化")
        self.assertEqual(read(self.company / "sub" / "b.xlsx"), "B 别人改了")
        self.assertEqual(read(self.team / "a.xlsx"), "A fixed")
        self.assertNotEqual(seen[0][0].parent.parent, self.team)
        self.assertIsNone(seen[0][1])
        self.assertIn("期间别人新加", distribute.markdown(result))

    def test_free_upload_source_can_be_pasted_paths(self):
        """起点贴多行（用户 2026-11-12）：起点取共同上级，只传贴的文件夹 / 文件"""
        other = Path(self.tmp.name) / "别处" / "10画面"
        write(other / "01" / "a.xlsx", "A")
        write(other / "01" / "不传.xlsx", "X")
        write(other / "03" / "b.xlsx", "B")
        write(other / "05" / "c.xlsx", "C")
        dest = Path(self.tmp.name) / "终点"
        dest.mkdir()
        text = f'"{other / "01" / "a.xlsx"}"\n{other / "01" / "a.xlsx"}\n{other / "03"}\n'
        preview = distribute.preview_free(text, str(dest))
        self.assertEqual(Path(preview["source_root"]), other.resolve())
        self.assertEqual({f["rel"]: f["action"] for f in preview["files"]}, {"01/a.xlsx": "new", "03/b.xlsx": "new"})
        self.assertEqual(preview["picks"], ["01/a.xlsx", "03"])
        single = distribute.preview_free(str(other / "05" / "c.xlsx"), str(dest))  # 只贴一个文件：只传这一份
        self.assertEqual([f["rel"] for f in single["files"]], ["c.xlsx"])
        whole = distribute.preview_free(str(other), str(dest))  # 一个文件夹：和以前一样
        self.assertEqual((len(whole["files"]), whole["picks"]), (4, []))
        with self.assertRaisesRegex(ValueError, "起点里找不到"):
            distribute.preview_free(f"{other / '03'}\n{other / '没有.xlsx'}", str(dest))

    def test_free_upload_without_manifest_tracks_the_approved_source(self):
        other = Path(self.tmp.name) / "别处"
        write(other / "a.xlsx", "A 别处")
        write(other / "new.xlsx", "N")
        dest = Path(self.tmp.name) / "终点"
        write(dest / "a.xlsx", "A 终点原来的")
        preview = distribute.preview_free(str(other), str(dest))
        self.assertEqual({f["rel"]: f["action"] for f in preview["files"]}, {"a.xlsx": "review", "new.xlsx": "new"})
        self.patches()
        with patch.object(excel_diff, "diff", return_value={"identical": False}):
            distribute.start_free(preview["id"])
            self.review("a.xlsx", "upload")
            result = self.wait()
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual(read(dest / "a.xlsx"), "A 别处 +标准化")
        write(other / "a.xlsx", "A 第二版")
        again = distribute.preview_free(str(other), str(dest))
        self.assertEqual({f["rel"]: f["action"] for f in again["files"]}, {"a.xlsx": "overwrite", "new.xlsx": "overwrite"})
        unrelated = Path(self.tmp.name) / "另一个起点"
        write(unrelated / "a.xlsx", "和上次的起点无关")
        self.assertEqual(distribute.preview_free(str(unrelated), str(dest))["files"][0]["action"], "review")

    def test_failed_preparation_skips_only_that_file(self):
        other = Path(self.tmp.name) / "别处"
        write(other / "a.xlsx", "A")
        write(other / "b.xlsx", "B")
        dest = Path(self.tmp.name) / "终点"
        dest.mkdir()

        def prepare(get_excel, local, reference, where):
            if local.name == "b.xlsx":
                where["step"] = "视图标准化"
                raise ValueError("打不开")
            return dict(PREP)

        preview = distribute.preview_free(str(other), str(dest))
        self.patches(prepare)
        distribute.start_free(preview["id"])
        result = self.wait()
        # 用户 2026-11-10：没做成的那份不传，其余照传
        self.assertEqual(result["summary"]["已复制"], 1)
        self.assertTrue((dest / "a.xlsx").exists())
        self.assertFalse((dest / "b.xlsx").exists())
        self.assertIn("上传前视图标准化没做成", result["files"][1]["error"])

    def test_dest_changed_after_preview_is_not_overwritten(self):
        other = Path(self.tmp.name) / "别处"
        write(other / "a.xlsx", "A")
        dest = Path(self.tmp.name) / "终点"
        dest.mkdir()
        preview = distribute.preview_free(str(other), str(dest))
        write(dest / "a.xlsx", "刚被别人放进来")
        self.patches()
        distribute.start_free(preview["id"])
        result = self.wait()
        self.assertEqual(result["files"][0]["status"], "conflict")
        self.assertEqual(read(dest / "a.xlsx"), "刚被别人放进来")

    def test_free_upload_from_task_work_to_relay_writes_manifest(self):
        """用户 2026-11-12：起点 → 终点从任务处理文件夹传到团队中转，也要写拿取版本.json"""
        preview = distribute.preview_free(str(self.work), str(self.team))
        self.assertEqual(preview["relay_task"]["id"], self.task["id"])
        self.patches()
        distribute.start_free(preview["id"])
        result = self.wait()
        self.assertEqual(result["summary"]["已复制"], 2)
        manifest = json.loads(read(self.team / distribute.MANIFEST_NAME))
        self.assertEqual(Path(manifest["shared"]), self.company.resolve())
        self.assertEqual(manifest["files"]["a.xlsx"]["shared_sha256"], distribute._digest(self.company / "a.xlsx"))
        self.assertEqual(manifest["files"]["a.xlsx"]["sha256"], distribute._digest(self.team / "a.xlsx"))
        self.assertIn("拿取版本清单", distribute.markdown(result))
        # 再从中转传回公司：按清单核对，公司没变 → 直接覆盖
        back = distribute.preview_free(str(self.team), str(self.company))
        self.assertEqual({f["rel"]: f["guard"] for f in back["files"]}, {"a.xlsx": "manifest", "sub/b.xlsx": "manifest"})
        self.assertEqual({f["rel"]: f["action"] for f in back["files"]}, {"a.xlsx": "overwrite", "sub/b.xlsx": "overwrite"})

    def test_free_upload_from_task_subfolder_maps_to_shared_subfolder(self):
        preview = distribute.preview_free(str(self.work / "sub"), str(self.team))
        self.assertEqual(Path(preview["relay_task"]["shared"]), self.company.resolve() / "sub")

    def test_free_upload_back_to_shared_or_unrelated_source_writes_no_manifest(self):
        self.assertNotIn("relay_task", distribute.preview_free(str(self.work), str(self.company)))
        other = Path(self.tmp.name) / "别处"
        write(other / "a.xlsx", "A")
        self.assertNotIn("relay_task", distribute.preview_free(str(other), str(self.team)))

    def test_local_desktop_shared_is_not_a_real_share(self):
        self.assertTrue(distribute._is_local_shared(Path(r"D:/Users/s.go/Desktop/doublecheck/10画面")))
        self.assertFalse(distribute._is_local_shared(Path(r"D:/Users/k.haku/OneDrive - 会社/Desktop/x")))
        self.assertFalse(distribute._is_local_shared(Path(r"D:/Users/k.haku/会社/共用 - ファミリー/50Common")))

    def test_start_and_dest_cannot_contain_each_other(self):
        with self.assertRaises(ValueError):
            distribute.preview_free(str(self.company), str(self.company / "sub"))


if __name__ == "__main__":
    unittest.main()
