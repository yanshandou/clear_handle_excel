"""几个账号分着做（用户 2026-11-07）：张三上传到团队中转并写拿取版本清单，李四从中转「起点 → 终点」传到公司"""
import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

from core import distribute, storage, upload_prep, work_task, workflow as wf


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

    def test_free_upload_from_relay_overwrites_unchanged_and_blocks_changed(self):
        self.relay_upload()
        write(self.company / "sub" / "b.xlsx", "B 别人改了")
        write(self.company / "c.xlsx", "C 别人新加")
        preview = distribute.preview_free(str(self.team), str(self.company))
        actions = {f["rel"]: f["action"] for f in preview["files"]}
        self.assertEqual(actions, {"a.xlsx": "overwrite", "sub/b.xlsx": "conflict"})
        self.assertEqual(preview["added"], ["c.xlsx"])
        seen = self.patches()
        distribute.start_free(preview["id"], False)
        result = self.wait()
        self.assertEqual(result["summary"]["已复制"], 1)
        self.assertEqual(read(self.company / "a.xlsx"), "A fixed +标准化")
        self.assertEqual(read(self.company / "sub" / "b.xlsx"), "B 别人改了")
        self.assertEqual(read(self.team / "a.xlsx"), "A fixed")  # 起点不动，标准化做在临时副本上
        self.assertNotEqual(seen[0][0].parent.parent, self.team)
        self.assertIsNone(seen[0][1])  # 没有参考，不恢复颜色
        self.assertIn("期间别人新加", distribute.markdown(result))

    def test_free_upload_without_manifest_needs_confirm(self):
        other = Path(self.tmp.name) / "别处"
        write(other / "a.xlsx", "A 别处")
        write(other / "new.xlsx", "N")
        dest = Path(self.tmp.name) / "终点"
        write(dest / "a.xlsx", "A 终点原来的")
        preview = distribute.preview_free(str(other), str(dest))
        self.assertEqual({f["rel"]: f["action"] for f in preview["files"]}, {"a.xlsx": "confirm", "new.xlsx": "new"})
        with self.assertRaises(ValueError):
            distribute.start_free(preview["id"], False)
        self.patches()
        distribute.start_free(preview["id"], True)
        result = self.wait()
        self.assertEqual(result["summary"]["已复制"], 2)
        self.assertEqual(read(dest / "a.xlsx"), "A 别处 +标准化")
        # 传过一次就有记录，下次不用再确认
        write(other / "a.xlsx", "A 第二版")
        again = distribute.preview_free(str(other), str(dest))
        self.assertEqual({f["rel"]: f["action"] for f in again["files"]}["a.xlsx"], "overwrite")

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
        distribute.start_free(preview["id"], False)
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
        distribute.start_free(preview["id"], False)
        result = self.wait()
        self.assertEqual(result["files"][0]["status"], "conflict")
        self.assertEqual(read(dest / "a.xlsx"), "刚被别人放进来")

    def test_start_and_dest_cannot_contain_each_other(self):
        with self.assertRaises(ValueError):
            distribute.preview_free(str(self.company), str(self.company / "sub"))


if __name__ == "__main__":
    unittest.main()
