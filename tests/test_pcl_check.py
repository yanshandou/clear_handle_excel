import base64
import http.client
import json
import tempfile
import threading
import unittest
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

from core import pcl_check, storage

PNG = b"\x89PNG\r\n\x1a\n" + b"\x00" * 16
PNG_URL = "data:image/png;base64," + base64.b64encode(PNG).decode()

CLOVER = """<?xml version="1.0" encoding="UTF-8"?>
<coverage generated="1762900000">
  <project timestamp="1762900000">
    <file name="/app/src/Order/Save.php">
      <class name="Save"><metrics methods="2" coveredmethods="1" statements="6" coveredstatements="3"/></class>
      <line num="5" type="method" name="run" count="2"/>
      <line num="6" type="stmt" count="2"/>
      <line num="7" type="cond" count="2" truecount="2" falsecount="0"/>
      <line num="8" type="stmt" count="0"/>
      <line num="9" type="stmt" count="0"/>
      <line num="12" type="method" name="unused" count="0"/>
      <line num="13" type="stmt" count="0"/>
      <metrics loc="20" ncloc="18" classes="1" methods="2" coveredmethods="1" conditionals="0" coveredconditionals="0"
               statements="6" coveredstatements="3" elements="8" coveredelements="4"/>
    </file>
    <file name="/app/src/Order/Load.php">
      <line num="3" type="stmt" count="1"/>
    </file>
    <metrics files="2"/>
  </project>
</coverage>"""


class PclCheckTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        patcher = patch.object(storage, "DATABASE_PATH", self.root / "execution.sqlite3")
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_verdict_changes_set_time_and_history_but_note_edits_do_not(self):
        target = pcl_check.create_target("画面 TB001")["id"]
        ok = pcl_check.save_result(target, "T01", {"verdict": "ok", "tester": "张三"})
        self.assertEqual((ok["verdict"], ok["tester"]), ("ok", "张三"))
        self.assertTrue(ok["tested_at"])
        noted = pcl_check.save_result(target, "T01", {"note": "回显一致"})
        self.assertEqual((noted["verdict"], noted["note"], len(noted["history"])), ("ok", "回显一致", 1))
        undone = pcl_check.save_result(target, "T01", {"verdict": ""})
        self.assertEqual((undone["tested_at"], [h["verdict"] for h in undone["history"]]), ("", ["ok", ""]))
        pcl_check.save_result(target, "T02", {"verdict": "ng"})
        self.assertEqual(pcl_check.list_targets()[0]["counts"], {"ok": 0, "ng": 1, "na": 0})
        with self.assertRaises(ValueError):
            pcl_check.save_result(target, "T02", {"verdict": "done"})
        with self.assertRaises(ValueError):
            pcl_check.save_result(target, "X1", {"verdict": "ok"})
        pcl_check.delete_target(target)
        self.assertEqual(pcl_check.list_targets(), [])
        with self.assertRaises(ValueError):
            pcl_check.load_target(target)

    def test_images_must_really_be_images(self):
        target = pcl_check.create_target("画面")["id"]
        image = pcl_check.add_image(target, "T07", PNG_URL, "debug.png")
        self.assertEqual(pcl_check.get_image(str(image["id"])), ("image/png", PNG))
        self.assertEqual(pcl_check.load_target(target)["results"]["T07"]["images"][0]["name"], "debug.png")
        fake = "data:image/png;base64," + base64.b64encode(b"<script>").decode()
        for bad in (fake, "data:text/html;base64,PGI+", "data:image/png;base64,***"):
            with self.assertRaises(ValueError):
                pcl_check.add_image(target, "T07", bad)
        pcl_check.delete_image(image["id"])
        with self.assertRaises(ValueError):
            pcl_check.get_image(str(image["id"]))

    def test_read_sources_walks_folders_and_decodes_shift_jis(self):
        project = self.root / "project"
        (project / "templates").mkdir(parents=True)
        (project / "vendor" / "lib").mkdir(parents=True)
        (project / "templates" / "edit.twig").write_bytes(b'{ view: "text" }\r\n')
        (project / "Save.php").write_bytes("<?php echo '登録';".encode("cp932"))
        (project / "vendor" / "lib" / "x.php").write_text("<?php", encoding="utf-8")
        (project / "readme.txt").write_text("skip", encoding="utf-8")
        (project / "huge.js").write_text("x" * (pcl_check.SOURCE_LIMIT + 1), encoding="utf-8")
        data = pcl_check.read_sources(f'"{project}"\n{self.root / "missing.php"}\n{project / "Save.php"}')
        files = {f["name"]: f["code"] for f in data["files"]}
        self.assertEqual(files, {"project/Save.php": "<?php echo '登録';", "project/templates/edit.twig": '{ view: "text" }\n'})
        self.assertEqual([m["path"] for m in data["missing"]], [str(self.root / "missing.php")])
        self.assertEqual([Path(s["path"]).name for s in data["skipped"]], ["huge.js"])
        with self.assertRaises(ValueError):
            pcl_check.read_sources(str(self.root / "missing.php"))

    def test_clover_coverage_lists_missed_lines_and_branches(self):
        coverage = pcl_check.parse_clover(CLOVER)
        self.assertEqual(coverage["prefix"], "/app/src/Order/")
        self.assertEqual(coverage["totals"]["lines"], {"covered": 4, "total": 7})
        save, load = coverage["files"][1], coverage["files"][0]
        self.assertEqual((save["short"], save["uncovered"], save["partial"]), ("Save.php", [[8, 9], [13, 13]], [[7, 7]]))
        self.assertEqual(save["uncovered_methods"], ["unused"])
        self.assertEqual((load["short"], load["lines"]), ("Load.php", {"covered": 1, "total": 1}))
        with self.assertRaises(ValueError):
            pcl_check.parse_clover('<!DOCTYPE x [<!ENTITY a "b">]><coverage><project/></coverage>')
        with self.assertRaises(ValueError):
            pcl_check.parse_clover("<phpunit/>")
        target = pcl_check.create_target("画面")["id"]
        path = self.root / "clover.xml"
        path.write_text(CLOVER, encoding="utf-8")
        pcl_check.import_coverage(target, path=str(path))
        self.assertEqual(pcl_check.load_target(target)["coverage"]["source"], str(path))

    def test_old_v1_export_imports_and_v2_export_round_trips(self):
        first = pcl_check.create_target("旧记录")["id"]
        pcl_check.import_results(first, {"format": "pcl-qa-checklist", "version": 1, "items": [
            {"id": "T01", "title": "x", "done": True, "note": "已确认"}, {"id": "T02", "done": False, "note": ""},
            {"id": "bad", "done": True}]})
        results = pcl_check.load_target(first)["results"]
        self.assertEqual((results["T01"]["verdict"], results["T01"]["note"], results["T02"]["verdict"]), ("ok", "已确认", ""))
        pcl_check.save_result(first, "T03", {"verdict": "ng", "tester": "李四"})
        pcl_check.add_image(first, "T03", PNG_URL, "ng.png")
        pcl_check.save_sources(first, [{"name": "a.twig", "type": "twig", "code": "x", "path": ""}])
        exported = json.loads(json.dumps(pcl_check.export_target(first)))
        second = pcl_check.create_target("复制")["id"]
        self.assertEqual(pcl_check.import_results(second, exported), {"items": 3, "images": 1})
        copy = pcl_check.load_target(second)
        self.assertEqual((copy["results"]["T03"]["verdict"], copy["results"]["T03"]["tester"]), ("ng", "李四"))
        self.assertEqual(len(copy["results"]["T03"]["history"]), 1)
        self.assertEqual(pcl_check.get_image(str(copy["results"]["T03"]["images"][0]["id"]))[1], PNG)
        self.assertEqual(copy["sources"][0]["name"], "a.twig")
        with self.assertRaises(ValueError):
            pcl_check.import_results(second, {"format": "other", "version": 1, "items": []})

    def test_selected_items_and_saved_test_runs(self):
        target = pcl_check.create_target("画面")["id"]
        self.assertIsNone(pcl_check.load_target(target)["target"]["selection"])  # 没存过 = 全选
        pcl_check.save_result(target, "T04", {"verdict": "ok"})
        pcl_check.save_result(target, "T09", {"verdict": "ng"})
        pcl_check.save_selection(target, ["T11", "T04", "T04"])
        info = pcl_check.load_target(target)["target"]
        self.assertEqual((info["selection"], info["counts"]), (["T04", "T11"], {"ok": 1, "ng": 0, "na": 0}))
        finding = {"status": "warn", "message": "拖拽表格的列开启了排序。", "junk": 1,
                   "files": [{"name": "a.twig", "count": 99, "lines": [[n, "x"] for n in range(1, 60)]}]}
        saved = pcl_check.save_runs(target, {"T04": {"status": "warn", "sig": "1:abc", "findings": [finding]},
                                             "T11": {"status": "none", "sig": "1:abc", "findings": []}}, "张三")
        self.assertEqual((saved["T04"]["tester"], bool(saved["T04"]["at"])), ("张三", True))
        runs = pcl_check.load_target(target)["runs"]
        file = runs["T04"]["findings"][0]["files"][0]
        self.assertEqual((runs["T11"]["status"], file["count"], len(file["lines"])), ("none", 99, pcl_check.MAX_RUN_LINES))
        self.assertNotIn("junk", runs["T04"]["findings"][0])
        for bad in ({"T04": {"status": "passed"}}, {"X": {"status": "none"}}, {}):
            with self.assertRaises(ValueError):
                pcl_check.save_runs(target, bad)
        copy = pcl_check.create_target("复制")["id"]
        pcl_check.import_results(copy, json.loads(json.dumps(pcl_check.export_target(target))))
        copied = pcl_check.load_target(copy)
        self.assertEqual((copied["target"]["selection"], copied["runs"]["T04"]["tester"]), (["T04", "T11"], "张三"))
        pcl_check.reset_results(target)
        cleared = pcl_check.load_target(target)
        self.assertEqual((cleared["runs"], cleared["target"]["selection"]), ({}, ["T04", "T11"]))

    def test_page_and_api_routes(self):
        from web_clip_fixer import WebHandler
        server = ThreadingHTTPServer(("127.0.0.1", 0), WebHandler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        self.addCleanup(lambda: (server.shutdown(), thread.join(), server.server_close()))

        def request(method, path, payload=None):
            connection = http.client.HTTPConnection(*server.server_address, timeout=10)
            body = json.dumps(payload).encode() if payload is not None else None
            try:
                connection.request(method, path, body, {"Content-Type": "application/json"} if body else {})
                response = connection.getresponse()
                data = response.read()
                is_json = "application/json" in response.getheader("Content-Type", "")
                return response.status, json.loads(data) if is_json else data
            finally:
                connection.close()

        status, page = request("GET", "/pcl")
        self.assertEqual(status, 200)
        self.assertIn(b'src="/static/pcl_check.js"', page)
        target = request("POST", "/api/pcl/targets/create", {"name": "画面"})[1]["target"]["id"]
        self.assertEqual(request("POST", "/api/pcl/result", {"target": target, "item": "T05", "verdict": "na"})[1]["result"]["verdict"], "na")
        image = request("POST", "/api/pcl/image", {"target": target, "item": "T05", "data": PNG_URL})[1]["image"]["id"]
        self.assertEqual(request("GET", f"/api/pcl/image?id={image}"), (200, PNG))
        self.assertEqual(request("GET", f"/api/pcl/target?id={target}")[1]["results"]["T05"]["verdict"], "na")
        status, error = request("POST", "/api/pcl/result", {"target": 999, "item": "T05", "verdict": "ok"})
        self.assertEqual((status, error["ok"]), (400, False))


if __name__ == "__main__":
    unittest.main()
