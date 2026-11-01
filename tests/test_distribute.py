import json
import tempfile
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from core import distribute, storage, workflow as wf


def plan(folders, **extra):
    return json.dumps({"label": "测试上传", "people": [{"name": "呉", "folders": folders}], **extra}, ensure_ascii=False)


def write(path: Path, text: str = "x") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


class DistributeTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.src, self.dst = root / "src", root / "dst"
        self.src.mkdir()
        self.dst.mkdir()
        self.db = patch.object(storage, "DATABASE_PATH", root / "test.sqlite3")
        self.db.start()
        distribute._STATUS.clear()
        distribute._STATUS["running"] = False
        wf.STATE["busy"] = False

    def tearDown(self):
        self.db.stop()
        self.tmp.cleanup()

    def run_copy(self):
        distribute.start()
        for _ in range(200):
            if not distribute.status().get("running"):
                break
            time.sleep(0.05)
        return distribute.status()["applied"]

    def test_parse_accepts_user_format_and_checks_totals(self):
        text = json.dumps({"totalFiles": 10, "folderCount": 2, "people": [
            {"name": "张晓静", "totalFiles": 5, "folders": [
                {"source": ["10要件定義", "10画面"], "destination": ["张晓静", "10要件定義"], "files": 4, "owner": "呉"}]}]},
            ensure_ascii=False)
        result = distribute.parse_plan(text)
        folder = result["folders"][0]
        self.assertEqual(folder["source"], ["10要件定義", "10画面"])
        self.assertEqual(folder["owner"], "呉")
        self.assertEqual(folder["person"], "张晓静")
        self.assertEqual(len(result["notes"]), 3)  # 个人合计、总合计、文件夹数都对不上

    def test_parse_rejects_parent_paths(self):
        with self.assertRaises(ValueError):
            distribute.parse_plan(plan([{"source": ["..", "x"], "destination": ["a"], "files": 1}]))
        with self.assertRaises(ValueError):
            distribute.parse_plan("{not json")

    def test_only_excel_recursive_and_overwrite(self):
        write(self.src / "A" / "a.xlsx", "new-a")
        write(self.src / "A" / "sub" / "b.XLS", "b")
        write(self.src / "A" / "c.xlsb", "c")
        write(self.src / "A" / "note.pdf")
        write(self.src / "A" / "~$old.xlsx")
        write(self.dst / "呉" / "A" / "a.xlsx", "old-a")
        write(self.dst / "呉" / "A" / "keep.xlsm", "keep")
        prev = distribute.preview(plan([{"source": ["A"], "destination": ["呉", "A"], "files": 3}]), str(self.src), str(self.dst))
        folder = prev["folders"][0]
        self.assertEqual(folder["source_count"], 3)
        self.assertEqual({f["rel"]: f["action"] for f in folder["files"]},
                         {"a.xlsx": "overwrite", "sub/b.XLS": "new", "c.xlsb": "new"})
        self.assertEqual(folder["dest_extra"], ["keep.xlsm"])
        self.assertEqual(folder["other"], {".pdf": 1})
        self.assertEqual(len(folder["temp"]), 1)
        result = self.run_copy()
        self.assertEqual((self.dst / "呉" / "A" / "a.xlsx").read_text(encoding="utf-8"), "new-a")
        self.assertTrue((self.dst / "呉" / "A" / "sub" / "b.XLS").exists())
        self.assertFalse((self.dst / "呉" / "A" / "note.pdf").exists())
        self.assertTrue((self.src / "A" / "a.xlsx").exists())  # 只复制不移动
        folder = result["folders"][0]
        self.assertEqual(folder["dest_count"], 4)
        bad = [c for c in folder["checks"] if not c["ok"]]
        self.assertEqual([c["pair"] for c in bad], ["来源 ↔ 目的地", "JSON ↔ 目的地"])
        self.assertIn("目的地多 1 个", bad[0]["text"])
        self.assertEqual(folder["dest_extra"], ["keep.xlsm"])
        md = distribute.markdown(result)
        self.assertIn("🔁 覆盖", md)
        self.assertIn("keep.xlsm", md)
        self.assertNotRegex(md, r"\d{4}-\d{2}-\d{2} \d{2}:\d{2}")  # 记录文档不写日期时间
        self.assertEqual(distribute.history()[0]["id"], result["id"])

    def test_json_count_mismatch_hint_and_missing_source(self):
        write(self.src / "A" / "a.xlsx")
        write(self.src / "A" / "a.pdf")
        prev = distribute.preview(plan([
            {"source": ["A"], "destination": ["X"], "files": 2},
            {"source": ["Missing"], "destination": ["Y"], "files": 5},
        ]), str(self.src), str(self.dst))
        a, missing = prev["folders"]
        self.assertIn("JSON 2 个，来源 1 个，JSON 多 1 个", a["checks"][0]["text"])
        self.assertIn("非 Excel", a["hint"])
        self.assertEqual(missing["error"], "找不到来源文件夹")
        result = self.run_copy()
        self.assertTrue((self.dst / "X" / "a.xlsx").exists())
        self.assertTrue(distribute.needs_attention(result))

    def test_source_changed_after_preview_copies_nothing(self):
        write(self.src / "A" / "a.xlsx")
        distribute.preview(plan([{"source": ["A"], "destination": ["X"], "files": 1}]), str(self.src), str(self.dst))
        write(self.src / "A" / "new.xlsx")
        result = self.run_copy()
        self.assertTrue(result["stopped"])
        self.assertFalse((self.dst / "X").exists())

    def test_copy_failure_is_recorded(self):
        write(self.src / "A" / "a.xlsx")
        write(self.src / "A" / "b.xlsx")
        distribute.preview(plan([{"source": ["A"], "destination": ["X"], "files": 2}]), str(self.src), str(self.dst))
        real = distribute.copy_one

        def flaky(src, dst):
            if src.name == "b.xlsx":
                raise PermissionError(13, "denied", str(dst))
            return real(src, dst)

        with patch.object(distribute, "copy_one", flaky):
            result = self.run_copy()
        folder = result["folders"][0]
        self.assertEqual(folder["missing_in_dest"], ["b.xlsx"])
        self.assertEqual(result["summary"]["没复制成功"], 1)
        self.assertIn("来源多 1 个", [c for c in folder["checks"] if c["pair"] == "来源 ↔ 目的地"][0]["text"])

    def test_open_source_asks_and_is_skipped(self):
        write(self.src / "A" / "a.xlsx")
        write(self.src / "A" / "~$a.xlsx")
        prev = distribute.preview(plan([{"source": ["A"], "destination": ["X"], "files": 1}]), str(self.src), str(self.dst))
        self.assertTrue(prev["folders"][0]["files"][0]["source_open"])
        result = self.run_copy()
        self.assertEqual(result["folders"][0]["files"][0]["status"], "failed")
        self.assertIn("Excel 中打开", result["folders"][0]["files"][0]["error"])
        self.assertTrue((self.src / "A" / "~$a.xlsx").exists())  # 来源一点不动

    def test_same_root_rejected(self):
        with self.assertRaises(ValueError):
            distribute.preview(plan([{"source": ["A"], "destination": ["A"], "files": 1}]), str(self.src), str(self.src))


if __name__ == "__main__":
    unittest.main()
