import struct
import tempfile
import unittest
from pathlib import Path

from core import icon_library as L
from core.pdf_detector import fitz


def png(width=300, height=200, box=(100, 50, 200, 150), color=(30, 120, 250), alpha=True) -> bytes:
    """透明底上画一块方块（不用 Pillow，只用 PyMuPDF）"""
    pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, width, height), 1 if alpha else 0)
    if alpha:
        pix.clear_with()  # 全透明
    else:
        pix.clear_with(255)
    x0, y0, x1, y1 = box
    for y in range(y0, y1):
        for x in range(x0, x1):
            pix.set_pixel(x, y, color + ((255,) if alpha else ()))
    return pix.tobytes("png")


def ico_sizes(body: bytes) -> list[tuple[int, bytes]]:
    _, kind, count = struct.unpack_from("<HHH", body, 0)
    assert kind == 1
    out = []
    for n in range(count):
        w, _h, _c, _r, _p, bpp, size, offset = struct.unpack_from("<BBBBHHII", body, 6 + 16 * n)
        assert bpp == 32
        out.append((w or 256, body[offset:offset + size]))
    return out


class IconLibraryTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)
        self.old = L.LIB_DIR
        L.LIB_DIR = self.root / "图标库"
        self.addCleanup(setattr, L, "LIB_DIR", self.old)

    def test_png_becomes_multi_size_ico_with_transparency(self):
        result = L.add("蓝色方块.png", png(), "文件夹")
        icon = result["icon"]
        self.assertEqual(icon["name"], "蓝色方块")
        self.assertEqual(icon["category"], "文件夹")
        self.assertTrue(result["info"]["transparent"])
        body = (L.LIB_DIR / "icons" / f"{icon['id']}.ico").read_bytes()
        entries = ico_sizes(body)
        self.assertEqual(sorted(w for w, _ in entries), sorted(L.SIZES))
        for size, data in entries:
            if size == 256:
                self.assertTrue(data.startswith(b"\x89PNG"))
            else:
                header = struct.unpack_from("<IiiHH", data, 0)
                self.assertEqual(header[:3], (40, size, size * 2))
                # 裁掉透明边后方块铺满，左下角像素是方块颜色、完全不透明（BGRA）
                b, g, r, a = data[40:44]
                self.assertEqual((r, g, b, a), (30, 120, 250, 255))

    def test_without_trim_keeps_transparent_margin(self):
        icon = L.add("a.png", png(), trim=False)["icon"]
        entries = dict(ico_sizes((L.LIB_DIR / "icons" / f"{icon['id']}.ico").read_bytes()))
        self.assertEqual(entries[16][40:44], b"\x00\x00\x00\x00")  # 左下角在透明边里

    def test_not_transparent_and_small_are_reported(self):
        info = L.add("photo.png", png(60, 40, (0, 0, 60, 40), alpha=False))["info"]
        self.assertFalse(info["transparent"])
        self.assertTrue(info["small"])

    def test_bad_file_is_rejected(self):
        with self.assertRaises(ValueError):
            L.add("x.png", b"not a picture")
        with self.assertRaises(ValueError):
            L.add("x.png", b"")

    def test_update_category_search_list_and_rename_category(self):
        a = L.add("a.png", png(), "旧分类")["icon"]
        L.add("b.png", png(), "")
        L.update(a["id"], name="  新  名字 ", category="工作")
        data = L.listing()
        self.assertEqual({c["name"]: c["count"] for c in data["categories"]}, {"工作": 1, "": 1})
        self.assertEqual(next(i for i in data["icons"] if i["id"] == a["id"])["name"], "新 名字")
        self.assertEqual(L.rename_category("工作", "项目")["count"], 1)
        self.assertEqual(L.listing()["categories"][0]["name"], "项目")

    def test_apply_writes_desktop_ini_and_keeps_other_settings(self):
        icon = L.add("a.png", png())["icon"]
        folder = self.root / "目标 文件夹"
        folder.mkdir()
        (folder / "desktop.ini").write_bytes("[.ShellClassInfo]\r\nIconFile=old.ico\r\nIconIndex=0\r\nLocalizedResourceName=名前\r\n[ViewState]\r\nMode=4\r\n".encode("utf-16"))
        result = L.apply(icon["id"], f'"{folder}"\n{folder}\n')
        self.assertEqual(result["ok_count"], 1)
        self.assertEqual(len(result["results"]), 1)  # 重复的一行只算一次
        text = (folder / "desktop.ini").read_bytes().decode("utf-16")
        self.assertIn(f"IconResource={L.ICON_PREFIX}{icon['id']}_1.ico,0", text)
        self.assertNotIn("IconFile=old.ico", text)
        self.assertIn("LocalizedResourceName=名前", text)
        self.assertIn("Mode=4", text)
        self.assertTrue((folder / f"{L.ICON_PREFIX}{icon['id']}_1.ico").is_file())
        self.assertEqual(L.listing()["icons"][0]["folders"], [str(folder)])

        # 恢复默认：只去掉本工具加的，别的设置留着
        L.restore(str(folder))
        text = (folder / "desktop.ini").read_bytes().decode("utf-16")
        self.assertNotIn("IconResource", text)
        self.assertIn("LocalizedResourceName=名前", text)
        self.assertEqual(list(folder.glob("*.ico")), [])
        self.assertEqual(L.listing()["icons"][0]["folders"], [])

    def test_replace_updates_applied_folders_and_restore_removes_ini(self):
        icon = L.add("a.png", png())["icon"]
        folder = self.root / "f"
        folder.mkdir()
        L.apply(icon["id"], str(folder))
        result = L.replace(icon["id"], "b.png", png(color=(200, 10, 10)))
        self.assertEqual(result["icon"]["version"], 2)
        self.assertTrue(result["folders"][0]["ok"])
        self.assertEqual(sorted(p.name for p in folder.iterdir()), sorted(["desktop.ini", f"{L.ICON_PREFIX}{icon['id']}_2.ico"]))
        L.restore(str(folder))
        self.assertEqual(list(folder.iterdir()), [])  # desktop.ini 只有本工具的内容 → 整个删掉

    def test_restore_leaves_icon_set_by_someone_else(self):
        folder = self.root / "f"
        folder.mkdir()
        (folder / "desktop.ini").write_text("[.ShellClassInfo]\r\nIconResource=C:\\x.ico,0\r\n")
        result = L.restore(str(folder))
        self.assertFalse(result["results"][0]["ok"])
        self.assertIn("C:\\x.ico", (folder / "desktop.ini").read_text())

    def test_apply_reports_missing_folder_and_moves_record_between_icons(self):
        a = L.add("a.png", png())["icon"]
        b = L.add("b.png", png())["icon"]
        folder = self.root / "f"
        folder.mkdir()
        r = L.apply(a["id"], f"{folder}\n{self.root / '没有'}\nrelative\\path")
        self.assertEqual(r["ok_count"], 1)
        self.assertEqual([x["ok"] for x in r["results"]], [True, False, False])
        L.apply(b["id"], str(folder))
        folders = {i["id"]: i["folders"] for i in L.listing()["icons"]}
        self.assertEqual(folders[a["id"]], [])
        self.assertEqual(folders[b["id"]], [str(folder)])
        self.assertEqual(len(list(folder.glob("*.ico"))), 1)  # 换图标时旧的那份删掉

    def test_tool_icon_and_delete(self):
        icon = L.add("a.png", png())["icon"]
        L.set_tool_icon("clean", icon["id"])
        self.assertEqual(L.listing()["tools"], {"clean": icon["id"]})
        with self.assertRaises(ValueError):
            L.set_tool_icon("../x", icon["id"])
        deleted = L.delete(icon["id"])
        self.assertEqual(deleted["tools"], ["clean"])
        data = L.listing()
        self.assertEqual(data["icons"], [])
        self.assertEqual(data["tools"], {})
        self.assertEqual(list((L.LIB_DIR / "icons").iterdir()), [])
        with self.assertRaises(ValueError):
            L.icon_file(icon["id"], "png")

    def test_broken_index_is_backed_up(self):
        L.LIB_DIR.mkdir(parents=True)
        (L.LIB_DIR / "index.json").write_text("{broken", encoding="utf-8")
        self.assertEqual(L.listing()["icons"], [])
        self.assertTrue(list(L.LIB_DIR.glob("index.broken-*.json")))

    def test_ini_line_editing(self):
        self.assertEqual(L._set_icon_lines([], "x.ico,0"), ["[.ShellClassInfo]", "IconResource=x.ico,0"])
        self.assertEqual(L._set_icon_lines(["[ViewState]", "Mode=4"], "x.ico,0"),
                         ["[.ShellClassInfo]", "IconResource=x.ico,0", "", "[ViewState]", "Mode=4"])
        self.assertEqual(L._set_icon_lines(["[.shellclassinfo]", "iconresource=a,0", "InfoTip=t"], None),
                         ["[.shellclassinfo]", "InfoTip=t"])


if __name__ == "__main__":
    unittest.main()
