"""Exercise the real toolbox handler and saved-plan to copy integration."""
import http.client
import json
import tempfile
import threading
import time
import unittest
from http.server import ThreadingHTTPServer
from pathlib import Path
from unittest.mock import patch

from core import storage, upload_manager, distribute, workflow
from web_clip_fixer import WebHandler


class UploadManagerIntegrationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        for module, key, value in [
            (upload_manager, 'DATABASE', self.root / 'manager.sqlite3'),
            (storage, 'DATABASE_PATH', self.root / 'execution.sqlite3'),
            (distribute, '_STATUS', {'running': False}),
        ]:
            patcher = patch.object(module, key, value)
            patcher.start()
            self.addCleanup(patcher.stop)
        self.server = ThreadingHTTPServer(('127.0.0.1', 0), WebHandler)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()
        self.addCleanup(self.close_server)

    def close_server(self):
        self.server.shutdown()
        self.thread.join()
        self.server.server_close()

    def request(self, method, path, payload=None, origin=None):
        connection = http.client.HTTPConnection(*self.server.server_address, timeout=10)
        headers = {}
        if payload is not None:
            payload = json.dumps(payload, ensure_ascii=False).encode('utf-8')
            headers['Content-Type'] = 'application/json'
        if origin:
            headers['Origin'] = origin
        try:
            connection.request(method, path, payload, headers)
            response = connection.getresponse()
            data = response.read()
            return response.status, json.loads(data) if 'application/json' in response.getheader('Content-Type', '') else data
        finally:
            connection.close()

    def test_manager_routes_and_local_resources(self):
        status, page = self.request('GET', '/distribute/manage')
        self.assertEqual(status, 200)
        self.assertNotIn(b'node_modules', page)
        self.assertNotIn(b'https://', page)
        self.assertIn(b'type="date"', page)
        self.assertIn(b'id="manager-task-upload" class="button primary" href="/distribute"', page)
        status, home = self.request('GET', '/')
        self.assertEqual(status, 200)
        self.assertIn(b'href="/distribute" id="workTaskUploadLink"', home)
        self.assertEqual(home.count(b'class="feature-title"'), 9)
        for filename in ('app.js', 'planner.js', 'icons.js', 'integration.css'):
            self.assertEqual(self.request('GET', '/static/upload_manager/' + filename)[0], 200)
        self.assertEqual(self.request('GET', '/api/upload-manager/items'), (200, []))
        self.assertEqual(self.request('GET', '/api/upload-manager/items', origin='https://example.invalid')[0], 403)

    def test_saved_snapshot_can_be_copied_and_stays_independent_of_items(self):
        status, item = self.request('POST', '/api/upload-manager/items', {
            'group': '父', 'name': '子', 'total': 1, 'processed': 1,
            'date': '2026-11-01', 'owner': '原负责人', 'complete': True,
        })
        self.assertEqual(status, 201)
        manifest = {'version': 2, 'label': '集成测试', 'people': [{'name': '甲', 'folders': [
            {'source': ['父', '子'], 'destination': ['甲', '父', '子'], 'files': 1, 'owner': '原负责人'},
        ]}]}
        payload = {'label': '集成测试', 'manifest': manifest, 'assignments': [
            {'id': item['id'], 'version': item['version'], 'uploader': '甲', 'uploadDate': '2026-11-04'},
        ]}
        status, saved = self.request('POST', '/api/upload-manager/allocations/confirm', payload)
        self.assertEqual(status, 200)
        self.assertEqual(saved['items'][0]['uploaded'], 0)
        self.assertEqual(saved['items'][0]['status'], 'uploading')
        self.assertEqual(self.request('POST', '/api/upload-manager/allocations/confirm', payload)[0], 409)
        path = '/api/upload-manager/allocation-history/' + saved['historyId'] + '/manifest'
        self.assertEqual(self.request('GET', path), (200, manifest))
        self.assertEqual(self.request('DELETE', '/api/upload-manager/items/' + item['id'], {'version': saved['items'][0]['version']})[0], 200)
        self.assertEqual(self.request('GET', path), (200, manifest))
        self.assertEqual(self.request('GET', '/api/upload-manager/allocation-history/missing/manifest')[0], 404)

        source, dest = self.root / 'source', self.root / 'dest'
        folder = source / '父' / '子'
        folder.mkdir(parents=True)
        dest.mkdir()
        (folder / 'test.xlsx').write_bytes(b'test Excel content')
        (folder / 'notes.txt').write_text('do not copy', encoding='utf-8')
        with patch.dict(workflow.STATE, {'busy': False, 'maintenance': ''}):
            status, preview = self.request('POST', '/api/distribute/preview', {
                'source_root': str(source), 'dest_root': str(dest), 'plan': json.dumps(manifest),
            })
            self.assertEqual(status, 200)
            self.assertEqual(preview['preview']['folders'][0]['source_count'], 1)
            self.assertFalse((dest / '甲').exists())
            self.assertEqual(self.request('POST', '/api/distribute/start', {})[0], 200)
            deadline = time.monotonic() + 10
            while distribute.status().get('running') and time.monotonic() < deadline:
                time.sleep(0.01)
            self.assertFalse(distribute.status().get('running'))
        self.assertEqual((dest / '甲' / '父' / '子' / 'test.xlsx').read_bytes(), b'test Excel content')
        self.assertFalse((dest / '甲' / '父' / '子' / 'notes.txt').exists())
        self.assertTrue((folder / 'test.xlsx').exists())

    def test_current_task_upload_routes_record_only_selected_files(self):
        shared = self.root / 'mock-shared'
        shared.mkdir()
        status, result = self.request('POST', '/api/tasks/new', {
            'shared': str(shared), 'work_root': str(self.root / 'local'),
        })
        self.assertEqual(status, 200)
        task = result['task']
        for name in ['selected.xlsx', 'other.xlsx']:
            (Path(task['work']) / name).write_bytes(b'local Excel')
        payload = {'task_id': task['id']}
        status, result = self.request('POST', '/api/distribute/task_files', payload)
        self.assertEqual(status, 200)
        self.assertEqual([f['rel'] for f in result['files']], ['other.xlsx', 'selected.xlsx'])
        self.assertEqual(self.request('POST', '/api/distribute/task_upload', {**payload, 'files': ['../outside.xlsx']})[0], 400)
        with patch.dict(workflow.STATE, {'busy': False, 'maintenance': ''}):
            self.assertEqual(self.request('POST', '/api/distribute/task_upload', {
                **payload, 'files': ['selected.xlsx'],
            })[0], 200)
            deadline = time.monotonic() + 10
            while distribute.status().get('running') and time.monotonic() < deadline:
                time.sleep(0.01)
            self.assertFalse(distribute.status().get('running'))
        self.assertEqual((shared / 'selected.xlsx').read_bytes(), b'local Excel')
        self.assertFalse((shared / 'other.xlsx').exists())
        status, result = self.request('POST', '/api/distribute/task_files', payload)
        self.assertEqual(status, 200)
        self.assertTrue(result['files'][1]['uploaded_current'])
        self.assertFalse(result['files'][0]['uploaded'])


if __name__ == '__main__':
    unittest.main()
