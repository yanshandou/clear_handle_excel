import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from core import structure_library


class StructureLibraryTests(unittest.TestCase):
    def setUp(self):
        library = tempfile.TemporaryDirectory()
        self.addCleanup(library.cleanup)
        self.lib = Path(library.name)
        library_path = patch.object(structure_library, "LIB_DIR", self.lib)
        library_path.start()
        self.addCleanup(library_path.stop)
        target = tempfile.TemporaryDirectory()
        self.addCleanup(target.cleanup)
        self.target = Path(target.name).resolve()
        self.source_dir = self.lib / "来源"
        self.source_dir.mkdir()

    def make_source(self, name="模板.xlsx", body=b"excel bytes"):
        path = self.source_dir / name
        path.write_bytes(body)
        return str(path)

    def make_template(self, name="基本設計", tree=None):
        template = structure_library.create(name)["template"]
        if tree is not None:
            template = structure_library.save_tree(template["id"], tree)["template"]
        return template

    def test_create_rename_delete_roundtrip(self):
        created = structure_library.create("基本設計")["template"]
        self.assertEqual(structure_library.listing()["templates"][0]["name"], "基本設計")
        renamed = structure_library.rename(created["id"], "詳細設計")["template"]
        self.assertEqual(renamed["name"], "詳細設計")
        structure_library.delete(created["id"])
        self.assertEqual(structure_library.listing()["templates"], [])
        with self.assertRaisesRegex(ValueError, "没有这个目录结构模板"):
            structure_library.rename(created["id"], "随便")

    def test_rejects_duplicate_and_invalid_names(self):
        self.make_template("基本設計")
        with self.assertRaisesRegex(ValueError, "已经有叫"):
            structure_library.create("基本設計")
        for bad in ("", "  ", "..", "."):
            with self.assertRaises(ValueError):
                structure_library.create(bad)

    def test_save_tree_assigns_and_keeps_ids(self):
        template = self.make_template(tree=[
            {"kind": "dir", "name": "01_要件", "children": [
                {"kind": "file", "name": "ビュー定義書.xlsx", "source": self.make_source()},
            ]},
        ])
        dir_id = template["tree"][0]["id"]
        file_id = template["tree"][0]["children"][0]["id"]
        self.assertTrue(dir_id and file_id)
        saved = structure_library.save_tree(template["id"], [
            {"id": dir_id, "kind": "dir", "name": "01_要件", "children": [
                {"id": file_id, "kind": "file", "name": "ビュー定義書.xlsx", "source": self.make_source()},
                {"kind": "file", "name": "新增.txt", "source": ""},
            ]},
        ])["template"]
        self.assertEqual(saved["tree"][0]["id"], dir_id)
        self.assertEqual(saved["tree"][0]["children"][0]["id"], file_id)
        self.assertTrue(saved["tree"][0]["children"][1]["id"])

    def test_save_tree_rejects_bad_nodes(self):
        template = self.make_template()
        cases = [
            [{"kind": "dir", "name": "同名"}, {"kind": "file", "name": "同名"}],
            [{"kind": "file", "name": "a.txt", "source": "相对路径.txt"}],
            [{"kind": "dir", "name": ".."}],
            [{"kind": "link", "name": "x"}],
            [{"kind": "file", "name": "a.txt", "source": "x" * 600}],
        ]
        for tree in cases:
            with self.assertRaises(ValueError):
                structure_library.save_tree(template["id"], tree)
        deep = tree = []
        for _ in range(structure_library.MAX_DEPTH + 1):
            node = {"kind": "dir", "name": "层", "children": []}
            tree.append(node)
            tree = node["children"]
        with self.assertRaisesRegex(ValueError, "层级太深"):
            structure_library.save_tree(template["id"], deep)

    def test_listing_warns_about_missing_sources(self):
        template = self.make_template(tree=[
            {"kind": "file", "name": "有的.txt", "source": self.make_source("有的.txt")},
            {"kind": "file", "name": "没的.txt", "source": str(self.source_dir / "没的.txt")},
            {"kind": "file", "name": "空来源.txt", "source": ""},
        ])
        warnings = structure_library.listing()["templates"][0]["warnings"]
        self.assertEqual({w["name"] for w in warnings}, {"没的.txt", "空来源.txt"})
        self.assertEqual(structure_library.save_tree(template["id"], template["tree"])["warnings"], warnings)

    def test_materialize_builds_dirs_and_copies_files(self):
        source = self.make_source(body=b"design v1")
        template = self.make_template(tree=[
            {"kind": "dir", "name": "01_要件", "children": [
                {"kind": "file", "name": "ビュー定義書.xlsx", "source": source},
            ]},
            {"kind": "dir", "name": "02_設計", "children": []},
        ])
        result = structure_library.materialize(template["id"], self.target)
        self.assertEqual((result["dirs"], result["files"], result["problems"]), (2, 1, []))
        copied = self.target / "01_要件" / "ビュー定義書.xlsx"
        self.assertEqual(copied.read_bytes(), b"design v1")
        self.assertTrue((self.target / "02_設計").is_dir())
        self.assertEqual(result["template"], "基本設計")

    def test_materialize_applies_exclude_and_rename(self):
        template = self.make_template(tree=[
            {"kind": "dir", "name": "01_要件", "children": [
                {"kind": "file", "name": "a.txt", "source": self.make_source("a.txt")},
            ]},
            {"kind": "dir", "name": "02_設計", "children": [
                {"kind": "file", "name": "b.txt", "source": self.make_source("b.txt")},
            ]},
        ])
        drop_dir, keep_dir = template["tree"]
        overrides = structure_library.parse_overrides({"structure_overrides": {
            "exclude": [drop_dir["id"]],
            "rename": {keep_dir["id"]: "02_詳細設計"},
        }})
        result = structure_library.materialize(template["id"], self.target, overrides)
        self.assertEqual((result["dirs"], result["files"]), (1, 1))
        self.assertFalse((self.target / "01_要件").exists())
        self.assertTrue((self.target / "02_詳細設計" / "b.txt").is_file())

    def test_materialize_reports_problems_without_stopping(self):
        template = self.make_template(tree=[
            {"kind": "file", "name": "缺失.txt", "source": str(self.source_dir / "不存在.txt")},
            {"kind": "file", "name": "空来源.txt", "source": ""},
            {"kind": "file", "name": "正常.txt", "source": self.make_source("正常.txt")},
        ])
        result = structure_library.materialize(template["id"], self.target)
        self.assertEqual(result["files"], 1)
        self.assertEqual(len(result["problems"]), 2)
        self.assertTrue((self.target / "正常.txt").is_file())

    def test_materialize_never_overwrites_existing_files(self):
        source = self.make_source(body=b"new")
        template = self.make_template(tree=[
            {"kind": "file", "name": "手放.txt", "source": source},
        ])
        (self.target / "手放.txt").write_bytes(b"manual")
        result = structure_library.materialize(template["id"], self.target)
        self.assertEqual(result["files"], 0)
        self.assertEqual((self.target / "手放.txt").read_bytes(), b"manual")
        self.assertTrue(result["items"][0]["exists"])
        again = structure_library.materialize(template["id"], self.target)
        self.assertEqual(again["problems"], [])

    def test_dir_icon_is_saved_and_applied_when_building(self):
        template = self.make_template(tree=[
            {"kind": "dir", "name": "01_要件", "icon": "abc123", "children": [
                {"kind": "dir", "name": "資料", "icon": "abc123", "children": []},
            ]},
            {"kind": "dir", "name": "02_設計", "children": []},
        ])
        self.assertEqual(template["tree"][0]["icon"], "abc123")
        self.assertNotIn("icon", template["tree"][1])
        calls = []

        def fake_apply(icon_id, folders):
            calls.append((icon_id, folders))
            return {"results": [{"folder": f, "ok": True} for f in folders]}

        with patch("core.icon_library.apply", side_effect=fake_apply):
            result = structure_library.materialize(template["id"], self.target)
        self.assertEqual(result["icons"], 2)
        self.assertEqual(calls, [("abc123", [str(self.target / "01_要件"), str(self.target / "01_要件" / "資料")])])
        self.assertEqual(result["problems"], [])

    def test_dir_icon_failure_is_only_a_problem(self):
        template = self.make_template(tree=[{"kind": "dir", "name": "01_要件", "icon": "abc123", "children": []}])
        with patch("core.icon_library.apply", side_effect=ValueError("图标库里没有这个图标")):
            result = structure_library.materialize(template["id"], self.target)
        self.assertTrue((self.target / "01_要件").is_dir())
        self.assertEqual(result["icons"], 0)
        self.assertEqual(result["problems"], [{"path": "01_要件", "reason": "套图标失败：图标库里没有这个图标"}])

    def test_dir_icon_rejects_bad_id_and_warns_when_missing(self):
        template = self.make_template()
        with self.assertRaisesRegex(ValueError, "图标编号不对"):
            structure_library.save_tree(template["id"], [{"kind": "dir", "name": "a", "icon": "../x", "children": []}])
        with patch("core.icon_library.listing", return_value={"icons": [{"id": "fff000"}]}):
            saved = structure_library.save_tree(template["id"], [{"kind": "dir", "name": "a", "icon": "abc123", "children": []}])
        self.assertEqual(len(saved["warnings"]), 1)
        self.assertIn("图标已不在图标库", saved["warnings"][0]["reason"])

    def test_parse_overrides_rejects_bad_input(self):
        for raw in ({"exclude": "not-a-list"}, {"rename": {"id": ".."}},
                    {"rename": {"id": ""}}, {"rename": []}):
            with self.assertRaises(ValueError):
                structure_library.parse_overrides({"structure_overrides": raw})
        self.assertEqual(structure_library.parse_overrides({}), {"exclude": set(), "rename": {}})
        cleaned = structure_library.parse_overrides({"structure_overrides": {"rename": {"id": "../evil"}}})
        self.assertEqual(cleaned["rename"], {"id": ".._evil"})

    def test_names_cannot_escape_the_ticket_folder(self):
        template = self.make_template(tree=[{"kind": "dir", "name": "a/b", "children": []}])
        self.assertEqual(template["tree"][0]["name"], "a_b")
        structure_library.materialize(template["id"], self.target)
        self.assertTrue((self.target / "a_b").is_dir())


if __name__ == "__main__":
    unittest.main()