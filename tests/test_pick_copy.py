import tempfile
import unittest
from pathlib import Path

from core import pick_copy


def write(path: Path, text: str = "x") -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


class PickCopyTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        root = Path(self.tmp.name)
        self.src = root / "doublecheck" / "10画面"
        self.dest = root / "发给同事"
        write(self.src / "01_CM" / "a.xlsx", "A")
        write(self.src / "01_CM" / "不要.xlsx", "N")
        write(self.src / "03_KS" / "b.xlsx", "B")
        write(self.src / "11_SY" / "マクロ" / "m.xlsm", "M")
        write(self.src / "11_SY" / "マクロ" / "~$m.xlsm", "T")
        write(self.src / "11_SY" / "说明.txt", "t")

    def text(self, *paths) -> str:
        return "\n".join(str(p) for p in paths)

    def test_keeps_structure_from_common_parent_and_only_listed_files(self):
        a = self.src / "01_CM" / "a.xlsx"
        text = self.text(a, f'"{a}"', a, self.src / "03_KS" / "b.xlsx", "", self.src / "11_SY" / "マクロ",
                         self.src / "03_KS" / "没有.xlsx")
        plan = pick_copy.preview(text, str(self.dest))
        self.assertEqual([f["rel"] for f in plan["files"]],
                         ["10画面/01_CM/a.xlsx", "10画面/03_KS/b.xlsx", "10画面/11_SY/マクロ/m.xlsm"])
        self.assertEqual(plan["duplicates"], 2)
        self.assertEqual(len(plan["missing"]), 1)
        self.assertFalse(self.dest.exists())  # 只看清单不建文件夹
        done = pick_copy.run(text, str(self.dest))
        self.assertEqual((done["copied"], done["failed"]), (3, 0))
        got = sorted(p.relative_to(self.dest).as_posix() for p in self.dest.rglob("*") if p.is_file())
        self.assertEqual(got, ["10画面/01_CM/a.xlsx", "10画面/03_KS/b.xlsx", "10画面/11_SY/マクロ/m.xlsm"])
        self.assertEqual((self.src / "01_CM" / "a.xlsx").read_text(encoding="utf-8"), "A")

    def test_existing_targets_need_confirmation(self):
        text = self.text(self.src / "01_CM" / "a.xlsx", self.src / "03_KS" / "b.xlsx")
        pick_copy.run(text, str(self.dest))
        write(self.src / "01_CM" / "a.xlsx", "A2")
        self.assertEqual(pick_copy.preview(text, str(self.dest))["exists"], 2)
        with self.assertRaisesRegex(ValueError, "同名"):
            pick_copy.run(text, str(self.dest))
        pick_copy.run(text, str(self.dest), overwrite=True)
        self.assertEqual((self.dest / "10画面" / "01_CM" / "a.xlsx").read_text(encoding="utf-8"), "A2")

    def test_bad_input(self):
        with self.assertRaisesRegex(ValueError, "一个也找不到"):
            pick_copy.preview(str(self.src / "没有.xlsx"), str(self.dest))
        with self.assertRaisesRegex(ValueError, "目标文件夹"):
            pick_copy.preview(str(self.src / "01_CM" / "a.xlsx"), "")
        # 目标正好是来源所在的地方：复制过去等于盖住自己
        with self.assertRaisesRegex(ValueError, "盖住自己"):
            pick_copy.preview(self.text(self.src / "01_CM" / "a.xlsx", self.src / "03_KS" / "b.xlsx"),
                              str(self.src.parent))
        # 一个文件、非 Excel 也行（发给同事什么都可能带）
        plan = pick_copy.preview(str(self.src / "11_SY" / "说明.txt"), str(self.dest))
        self.assertEqual([f["rel"] for f in plan["files"]], ["11_SY/说明.txt"])


if __name__ == "__main__":
    unittest.main()
