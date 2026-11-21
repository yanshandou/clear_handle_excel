import hashlib
import struct
import tempfile
from concurrent.futures import ThreadPoolExecutor
import unittest
from pathlib import Path
from unittest.mock import patch

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

    def test_trim_ignores_nearly_invisible_stray_pixels(self):
        pix = L._load_png(png(20, 20, (5, 6, 15, 14)))
        pix.set_pixel(0, 0, (1, 1, 1, 1))
        pix.set_pixel(19, 19, (2, 2, 2, 2))
        pix.set_pixel(5, 6, (4, 4, 4, 8))
        self.assertEqual(L._content_box(pix), (5, 6, 14, 13))
        cropped = L._crop(pix, L._content_box(pix))
        self.assertEqual(cropped.pixel(0, 0)[3], 8)

    def test_trim_preserves_entirely_faint_image(self):
        pix = fitz.Pixmap(fitz.csRGB, fitz.IRect(0, 0, 10, 10), 1)
        pix.clear_with()
        pix.set_pixel(3, 4, (1, 1, 1, 1))
        self.assertEqual(L._content_box(pix), (3, 4, 3, 4))
        pix.clear_with()
        self.assertEqual(L._content_box(pix), (0, 0, 9, 9))

    def test_retrim_removes_source_margin_and_keeps_rectangular_shape(self):
        pix = L._load_png(png(300, 300, (50, 100, 250, 200)))
        pix.set_pixel(1, 1, (1, 1, 1, 1))
        original = pix.tobytes("png")
        added = L.add("wide.png", original, trim=False)
        self.assertFalse(added["info"]["trimmed"])
        icon_id = added["icon"]["id"]

        result = L.retrim(icon_id)
        self.assertTrue(result["info"]["trimmed"])
        icon = L.listing()["icons"][0]
        self.assertEqual((icon["width"], icon["height"]), (300, 300))
        self.assertEqual((icon["render_width"], icon["render_height"]), (200, 100))
        preview, _ = L.icon_file(icon_id, "png")
        ico, _ = L.icon_file(icon_id, "ico")
        for data in (preview, dict(ico_sizes(ico))[256]):
            self.assertEqual(L._content_box(L._load_png(data)), (0, 64, 255, 191))
        self.assertEqual(L.icon_file(icon_id, "orig")[0], original)

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
        L.add("b.png", png(color=(50, 120, 250)), "")
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
        b = L.add("b.png", png(color=(50, 120, 250)))["icon"]
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

    def test_subfolders_lists_dirs_sorted_with_current_icon(self):
        icon = L.add("a.png", png())["icon"]
        parent = self.root / "parent"
        parent.mkdir()
        for name in ("b_dir", "A_dir"):
            (parent / name).mkdir()
        (parent / "file.txt").write_text("x")
        L.apply(icon["id"], str(parent / "A_dir"))
        result = L.subfolders(f'"{parent}"')
        self.assertEqual(result["dir"], str(parent))
        self.assertEqual(result["total"], 2)
        self.assertEqual(result["folders"], [
            {"path": str(parent / "A_dir"), "icon_id": icon["id"]},
            {"path": str(parent / "b_dir"), "icon_id": ""},
        ])

    def test_subfolders_rejects_bad_parent_and_caps_at_200(self):
        with self.assertRaisesRegex(ValueError, "完整路径"):
            L.subfolders("relative\\path")
        with self.assertRaisesRegex(ValueError, "找不到"):
            L.subfolders(str(self.root / "没有"))
        file = self.root / "f.txt"
        file.write_text("x")
        with self.assertRaisesRegex(ValueError, "不是文件夹"):
            L.subfolders(str(file))
        parent = self.root / "many"
        parent.mkdir()
        for n in range(201):
            (parent / f"d{n:03d}").mkdir()
        result = L.subfolders(str(parent))
        self.assertEqual(result["total"], 201)
        self.assertEqual(len(result["folders"]), 200)
        self.assertEqual(result["folders"][0]["path"], str(parent / "d000"))

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

    def test_duplicate_pixels_are_rejected_on_add_and_replace(self):
        a = L.add("a.png", png())['icon']
        b = L.add("b.png", png(color=(50, 120, 250)))['icon']
        # 不同画布和文件名，但裁去透明边后像素相同。
        same = png(400, 300, (150, 100, 250, 200))
        with patch.dict(L._fingerprint_cache, clear=True), patch.object(L, '_load_png', wraps=L._load_png) as decode:
            with self.assertRaisesRegex(L.DuplicateIconError, '重复'):
                L.add('another.png', same, '其他')
            decode.assert_called_once_with(same)
        with self.assertRaisesRegex(ValueError, '重复'):
            L.replace(b['id'], 'another.png', same)
        self.assertEqual(L.listing()['icons'][0]['id'], b['id'])
        self.assertEqual(L.listing()['icons'][0]['version'], 1)
        # 用自己的原图替换仍允许，可重新裁剪。
        self.assertEqual(L.replace(a['id'], 'a.png', same)['icon']['version'], 2)

    def test_original_fingerprint_is_saved_and_short_circuits_decode(self):
        original = png()
        L.add('first.png', original)
        fingerprints = L._load()['icons'][0]['fingerprints']
        self.assertEqual(fingerprints['original'], hashlib.sha256(original).hexdigest())
        self.assertIn(L._PIXEL_FINGERPRINT, fingerprints)

        with patch.dict(L._fingerprint_cache, clear=True), patch.object(L, '_load_png') as decode, \
                patch.object(Path, 'read_bytes') as read_original:
            with self.assertRaises(L.DuplicateIconError):
                L.add('renamed.png', original, trim=False)
            decode.assert_not_called()
            read_original.assert_not_called()
        self.assertEqual(len(L.listing()['icons']), 1)

    def test_legacy_original_short_circuits_decode(self):
        original = png()
        L.add('first.png', original)
        data = L._load()
        del data['icons'][0]['fingerprints']
        L._save(data)

        with patch.dict(L._fingerprint_cache, clear=True), patch.object(L, '_load_png') as decode:
            with self.assertRaises(L.DuplicateIconError):
                L.add('renamed.png', original)
            decode.assert_not_called()

    def test_legacy_duplicates_and_missing_original_are_reported(self):
        a = L.add('a.png', png())['icon']
        b = L.add('b.png', png(color=(50, 120, 250)))['icon']
        c = L.add('c.png', png(color=(60, 120, 250)))['icon']
        folder = L.LIB_DIR / 'icons'
        (folder / f"{b['id']}_orig.png").write_bytes((folder / f"{a['id']}_orig.png").read_bytes())
        (folder / f"{c['id']}_orig.png").unlink()
        result = L.listing()
        self.assertEqual(set(result['duplicates'][0]), {a['id'], b['id']})
        self.assertEqual(result['duplicate_unchecked'], [c['id']])

    def test_concurrent_duplicate_import_only_adds_one(self):
        image = png()
        def add_one(n):
            try:
                L.add(f'{n}.png', image)
                return True
            except ValueError:
                return False
        with ThreadPoolExecutor(max_workers=2) as pool:
            results = list(pool.map(add_one, range(2)))
        self.assertEqual(sorted(results), [False, True])
        self.assertEqual(len(L.listing()['icons']), 1)

    def test_retrim_works_with_preexisting_duplicate(self):
        a = L.add('a.png', png(), trim=False)['icon']
        b = L.add('b.png', png(color=(50, 120, 250)))['icon']
        folder = L.LIB_DIR / 'icons'
        original = (folder / f"{a['id']}_orig.png").read_bytes()
        (folder / f"{b['id']}_orig.png").write_bytes(original)
        before = L._load_png((folder / f"{a['id']}.png").read_bytes())
        self.assertLess(L._content_box(before)[2] - L._content_box(before)[0], 200)
        result = L.retrim(a['id'])
        after = L._load_png((folder / f"{a['id']}.png").read_bytes())
        self.assertEqual(L._content_box(after), (0, 0, 255, 255))
        self.assertEqual(result['icon']['version'], 2)
        self.assertEqual((folder / f"{a['id']}_orig.png").read_bytes(), original)

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
