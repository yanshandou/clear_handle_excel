"""图标上传的真实 HTTP 路径；使用临时图标库。"""
import json
import tempfile
import threading
import unittest
from unittest.mock import patch
from http.server import ThreadingHTTPServer
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen

from core import icon_library
from tests.test_icon_library import png
from web_clip_fixer import WebHandler


class IconHttpTests(unittest.TestCase):
    def setUp(self):
        system = patch.object(icon_library, "SYSTEM_ICONS", ())   # 系统图标另有测试，这里只看用户上传的
        system.start()
        self.addCleanup(system.stop)

    def test_duplicate_upload_and_retrim(self):
        with tempfile.TemporaryDirectory() as temp:
            previous = icon_library.LIB_DIR
            icon_library.LIB_DIR = Path(temp)
            server = ThreadingHTTPServer(('127.0.0.1', 0), WebHandler)
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                url = f'http://127.0.0.1:{server.server_port}/api/icons/'
                image = png()
                def upload(name):
                    request = Request(url + 'upload', data=image, headers={'X-File-Name': name, 'X-Icon-Trim': '0'}, method='POST')
                    return json.load(urlopen(request, timeout=10))
                first = upload('first.png')
                self.assertTrue(first['ok'])
                with self.assertRaises(HTTPError) as caught:
                    upload('other.png')
                self.assertEqual(caught.exception.code, 400)
                duplicate = json.load(caught.exception)
                self.assertEqual(duplicate['code'], 'duplicate_icon')
                self.assertIn('重复', duplicate['error'])
                caught.exception.close()
                self.assertEqual(len(icon_library.listing()['icons']), 1)
                request = Request(url + 'retrim', data=json.dumps({'id': first['icon']['id']}).encode(),
                                  headers={'Content-Type': 'application/json'}, method='POST')
                result = json.load(urlopen(request, timeout=10))
                self.assertEqual(result['icon']['version'], 2)
                self.assertTrue(result['info']['trimmed'])
                self.assertEqual((result['icon']['render_width'], result['icon']['render_height']), (100, 100))
                output = icon_library._load_png((Path(temp) / 'icons' / (first['icon']['id'] + '.png')).read_bytes())
                self.assertEqual(icon_library._content_box(output), (0, 0, 255, 255))
            finally:
                server.shutdown()
                server.server_close()
                thread.join()
                icon_library.LIB_DIR = previous


if __name__ == '__main__':
    unittest.main()
