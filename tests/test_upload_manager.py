import io
import json
import sqlite3
import tempfile
import threading
import unittest
from datetime import date, datetime
from pathlib import Path
from urllib.error import HTTPError
from urllib.request import Request, urlopen
from unittest.mock import patch

from openpyxl import load_workbook

from core import upload_manager as server
from core.upload_manager import export_workbook, parse_workbook

WORKBOOK = Path(__file__).resolve().parent.parent / 'tests' / 'fixtures' / 'upload_manager_source.xlsx'


MANIFEST = {'version': 1, 'label': '方案 1', 'people': [{'name': '甲', 'totalFiles': 10, 'folders': [{'source': ['父', '一'], 'destination': ['甲', '父', '一'], 'files': 10}]}]}


class AllocationDateTests(unittest.TestCase):
    def test_list_source_files_is_relative_sorted_and_safe(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory) / '父' / '子'
            (folder / '旧版').mkdir(parents=True)
            (folder / 'b.xlsx').write_text('x')
            (folder / '旧版' / 'a.xls').write_text('x')
            with patch.dict('os.environ', {'SOURCE_ROOT': directory}):
                self.assertEqual(server.list_source_files(['父', '子']), ['b.xlsx', '旧版/a.xls'])
                self.assertEqual(server.list_source_files(['父', '..']), [])
                self.assertEqual(server.list_source_files(['不存在']), [])
            with patch.dict('os.environ', {'SOURCE_ROOT': ''}):
                self.assertEqual(server.list_source_files(['父', '子']), [])

    def test_invalid_history_rolls_back_all_assignments(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [
                    {'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True},
                    {'group':'父', 'name':'二', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True},
                ])
                first, second = server.all_items(connection)
            assignments = [
                {'id':first['id'], 'version':first['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'},
                {'id':second['id'], 'version':second['version'], 'uploader':'乙', 'uploadDate':'2026-11-03'},
            ]
            payload = {'label':'方案 1', 'manifest':MANIFEST, 'assignments':assignments}
            invalid_payloads = [
                payload,
                {**payload, 'manifest':'not-json'},
                {**payload, 'manifest':{'version':1}},
                {**payload, 'manifest':None, 'assignments':[{**assignments[0]}, {**assignments[1], 'uploadDate':'2026-11-02'}]},
            ]
            for invalid in invalid_payloads:
                with self.assertRaises(ValueError):
                    with server.connect() as connection:
                        server.save_allocation_history(connection, invalid)
                with server.connect() as connection:
                    self.assertEqual([item['status'] for item in server.all_items(connection)], ['ready', 'ready'])
                    self.assertEqual(connection.execute('SELECT count(*) FROM allocation_history').fetchone()[0], 0)

    def test_history_http_returns_manifest_and_keeps_legacy_archive(self):
        legacy_archive = b'PK legacy zip bytes'
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [{'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True}])
                item = server.all_items(connection)[0]
                connection.execute("INSERT INTO allocation_history (id,created_at,planned_date,label,assignments,archive) VALUES ('legacy','2026-10-01T00:00:00','2026-10-02','旧方案','[]',?)", (legacy_archive,))
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        payload = {'label':'方案 1', 'manifest':MANIFEST,
                                   'assignments':[{'id':item['id'], 'version':item['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'}]}
                        request = Request(f'{url}/api/upload-manager/allocations/confirm', data=json.dumps(payload).encode(), headers={'Content-Type':'application/json'}, method='POST')
                        with urlopen(request) as response:
                            saved = json.load(response)
                        with urlopen(f'{url}/api/upload-manager/allocation-history') as response:
                            history = {entry['id']: entry for entry in json.load(response)}
                        self.assertEqual(len(history), 2)
                        current = history[saved['historyId']]
                        self.assertEqual(current['assignments'][0]['uploader'], '甲')
                        self.assertEqual(current['plannedDate'], '2026-11-02')
                        self.assertEqual(current['manifest'], MANIFEST)
                        self.assertFalse(current['hasArchive'])
                        self.assertIsNone(history['legacy']['manifest'])
                        self.assertTrue(history['legacy']['hasArchive'])
                        with urlopen(f'{url}/api/upload-manager/allocation-history/legacy/archive') as response:
                            self.assertEqual(response.headers.get_content_type(), 'application/zip')
                            self.assertEqual(response.read(), legacy_archive)
                        with self.assertRaises(HTTPError) as error:
                            urlopen(f"{url}/api/upload-manager/allocation-history/{saved['historyId']}/archive")
                        self.assertEqual(error.exception.code, 404)
                        error.exception.close()
                        with server.connect() as connection:
                            self.assertEqual(server.all_items(connection), saved['items'])
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_history_stores_snapshot_and_manifest_atomically(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [{'group':'父', 'name':'一', 'total':10, 'processed':10, 'date':'2026-10-30', 'complete':True}])
                item = server.all_items(connection)[0]
            assignment = {'id':item['id'], 'version':item['version'], 'uploader':'甲', 'uploadDate':'2026-11-02'}
            payload = {'label':'方案 1', 'assignments':[assignment], 'manifest':MANIFEST}
            with server.connect() as connection:
                saved, history_id = server.save_allocation_history(connection, payload)
                self.assertEqual(saved[0]['status'], 'uploading')
                self.assertEqual(saved[0]['uploaded'], 0)
                history = connection.execute('SELECT * FROM allocation_history WHERE id=?', (history_id,)).fetchone()
                self.assertEqual(json.loads(history['manifest']), MANIFEST)
                self.assertEqual(history['archive'], b'')
                self.assertEqual(json.loads(history['assignments'])[0]['name'], '一')
                connection.execute('DELETE FROM items WHERE id=?', (item['id'],))
            with self.assertRaises(server.AllocationConflict):
                with server.connect() as connection:
                    server.save_allocation_history(connection, payload)
            with server.connect() as connection:
                self.assertEqual(connection.execute('SELECT count(*) FROM allocation_history').fetchone()[0], 1)
                history = connection.execute('SELECT * FROM allocation_history WHERE id=?', (history_id,)).fetchone()
                self.assertEqual(json.loads(history['assignments'])[0]['name'], '一')
                self.assertEqual(json.loads(history['manifest']), MANIFEST)

    def test_folder_status_rules_and_manual_override(self):
        item = {'total': 10, 'processed': 0, 'date': '', 'uploaded': 0}
        self.assertEqual(server.folder_status(item), 'unstarted')
        self.assertEqual(server.folder_status({**item, 'date': '2026-10-30'}), 'processing')
        self.assertEqual(server.folder_status({**item, 'date': '2026-10-30', 'processed': 10}), 'ready')
        self.assertEqual(server.folder_status({**item, 'plannedDate': '2026-11-02'}), 'uploading')
        self.assertEqual(server.folder_status({**item, 'uploaded': 10}), 'uploaded')
        self.assertEqual(server.folder_status({**item, 'statusOverride': 'ready'}), 'ready')

    def test_allocation_date_and_legacy_default(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [
                    {'group': '父', 'name': '一', 'total': 10, 'processed': 10, 'date': '2026-10-30', 'complete': True},
                    {'group': '父', 'name': '二', 'total': 10, 'processed': 10, 'date': '2026-10-30', 'complete': True},
                ])
                first, second = server.all_items(connection)
            assignments = [
                {'id': first['id'], 'version': first['version'], 'uploader': '甲', 'uploadDate': '2026-11-02'},
                {'id': second['id'], 'version': second['version'], 'uploader': '乙', 'uploadDate': 'invalid'},
            ]
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.mark_allocations(connection, assignments)
            with server.connect() as connection:
                self.assertEqual([item['uploaded'] for item in server.all_items(connection)], [0, 0])
                self.assertEqual([item['plannedDate'] for item in server.all_items(connection)], ['', ''])
                del assignments[1]['uploadDate']
                saved = server.mark_allocations(connection, assignments)
                self.assertEqual([item['plannedDate'] for item in saved], ['2026-11-02', date.today().isoformat()])
                self.assertEqual([item['uploaded'] for item in saved], [0, 0])
                self.assertEqual([item['uploadDate'] for item in saved], ['', ''])
                self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])
                exported = server.export_workbook(saved)
            parsed, _ = server.parse_workbook(exported)
            self.assertEqual([(item['plannedDate'], item['statusOverride'], item['uploader']) for item in parsed],
                             [('2026-11-02', 'uploading', '甲'), (date.today().isoformat(), 'uploading', '乙')])
            with server.connect() as connection:
                server.merge_items(connection, parsed, preserve_upload=True)
                self.assertEqual([item['status'] for item in server.all_items(connection)], ['uploading', 'uploading'])


class ScriptResourceTests(unittest.TestCase):
    def test_script_resource_downloads(self):
        from web_clip_fixer import WebHandler
        with server.LocalHTTPServer(('127.0.0.1', 0), WebHandler) as http_server:
            with patch.object(server, 'PORT', http_server.server_address[1]):
                thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                thread.start()
                try:
                    for name in ('step1_copy_files.py', 'step2_copy_onedrive.py'):
                        with urlopen(f'http://127.0.0.1:{server.PORT}/static/upload_manager/{name}') as response:
                            self.assertEqual(response.read(), (server.ROOT / 'static' / 'upload_manager' / name).read_bytes())
                            self.assertEqual(response.headers.get_content_type(), 'text/plain')
                finally:
                    http_server.shutdown()
                    thread.join()


class WorkbookTests(unittest.TestCase):
    def test_excel_import_api_keeps_upload_and_rolls_back_conflicts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        original = {**records[0], 'uploaded': records[0]['total'], 'uploader': '原上传者'}
        incoming = {**records[0], 'uploaded': 0, 'uploader': '', 'owner': '新负责人'}
        new_item = {**records[2], 'uploaded': records[2]['total'], 'uploader': '新上传者'}
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [original])
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}/api/upload-manager/import'
                        def import_excel(items):
                            request = Request(url, data=export_workbook(items),
                                              headers={'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'}, method='POST')
                            return urlopen(request)
                        with import_excel([incoming, new_item]) as response:
                            self.assertEqual(json.load(response), {'inserted': 1, 'updated': 1})
                        with server.connect() as connection:
                            first, second = server.all_items(connection)
                        self.assertEqual((first['uploaded'], first['uploader'], first['owner']), (original['uploaded'], '原上传者', '新负责人'))
                        self.assertEqual((second['uploaded'], second['uploader']), (new_item['uploaded'], '新上传者'))
                        conflict = {**incoming, 'complete': False}
                        with self.assertRaises(HTTPError) as error:
                            import_excel([{**new_item, 'owner': '不应保存'}, conflict])
                        self.assertEqual(error.exception.code, 400)
                        self.assertIn('未导入任何记录', json.load(error.exception)['error'])
                        error.exception.close()
                        with server.connect() as connection:
                            first, second = server.all_items(connection)
                        self.assertEqual((first['uploaded'], first['uploader'], first['complete']), (original['uploaded'], '原上传者', True))
                        self.assertNotEqual(second['owner'], '不应保存')
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_excel_merge_preserves_uploaded_values_even_when_columns_exist(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        original = {**records[0], 'uploaded': records[0]['total'], 'uploader': '原上传者'}
        incoming = {**records[0], 'uploaded': 0, 'uploader': '', 'owner': '新负责人'}
        parsed, _ = parse_workbook(export_workbook([incoming]))
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [original])
            with server.connect() as connection:
                server.merge_items(connection, parsed, preserve_upload=True)
                saved = server.all_items(connection)[0]
            self.assertEqual((saved['uploaded'], saved['uploader'], saved['owner']), (original['uploaded'], '原上传者', '新负责人'))
            parsed[0]['complete'] = False
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.merge_items(connection, parsed, preserve_upload=True)
            with server.connect() as connection:
                saved = server.all_items(connection)[0]
                self.assertEqual((saved['uploaded'], saved['uploader'], saved['complete']), (original['uploaded'], '原上传者', True))

    def test_allocation_is_atomic_and_records_uploader(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[0], records[2]])
                first, second = server.all_items(connection)
            assignments = [{'id': first['id'], 'version': first['version'], 'uploader': '甲'},
                           {'id': second['id'], 'version': second['version'] + 1, 'uploader': '乙'}]
            with self.assertRaises(server.AllocationConflict):
                with server.connect() as connection:
                    server.mark_allocations(connection, assignments)
            with server.connect() as connection:
                self.assertEqual([item['uploaded'] for item in server.all_items(connection)], [0, 0])
                assignments[1]['version'] = second['version']
                saved = server.mark_allocations(connection, assignments)
                self.assertEqual([(item['uploaded'], item['uploader']) for item in saved],
                                 [(0, '甲'), (0, '乙')])
                self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])

    def test_customer_export_and_backup_uploader_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0].update(uploaded=records[0]['total'], uploader='甲', uploadDate='2026-10-29')
        records[1].update(total=145, processed=145, complete=True, uploaded=10, uploader='张晓静', uploadDate='2026-10-30')
        result = load_workbook(io.BytesIO(server.export_customer_workbook(records)))
        template = load_workbook(server.CUSTOMER_TEMPLATE)
        customer = result.active
        self.assertEqual(result.sheetnames, template.sheetnames)
        self.assertEqual(customer.max_row, 87)
        self.assertEqual([cell.value for cell in customer[2]], [cell.value for cell in template.active[2]])
        self.assertEqual(str(customer.merged_cells), str(template.active.merged_cells))
        self.assertEqual(customer.auto_filter.ref, template.active.auto_filter.ref)
        self.assertEqual(customer.column_dimensions['B'].width, template.active.column_dimensions['B'].width)
        self.assertEqual((customer['E3'].value, customer['F3'].value, customer['G3'].value), (datetime(2026, 10, 29), '甲', 2))
        self.assertEqual((customer['E4'].value, customer['F4'].value, customer['G4'].value), (datetime(2026, 10, 30), '静', 10))
        self.assertEqual([customer.cell(5, col).value for col in (5, 6, 7)], [None, None, None])
        self.assertEqual(customer['H4'].value, '=IF(AND(D4<>"",G4<>"",D4=G4),"〇","")')
        self.assertEqual(customer['D87'].value, '=SUM(D3:D86)')
        self.assertEqual(customer['G87'].value, '=SUM(G3:G86)')
        self.assertEqual(customer['E3'].number_format, template.active['E5'].number_format)
        other = result['変更対象']
        expected_other = template['変更対象']
        self.assertEqual([[cell.value for cell in row] for row in other], [[cell.value for cell in row] for row in expected_other])
        reloaded, _ = parse_workbook(export_workbook(records))
        self.assertEqual([item['uploader'] for item in reloaded[:2]], ['甲', '张晓静'])
        self.assertEqual([item['uploadDate'] for item in reloaded[:2]], ['2026-10-29', '2026-10-30'])
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, reloaded)
                saved = server.all_items(connection)
                self.assertEqual([item['uploader'] for item in saved[:2]], ['甲', '张晓静'])
                self.assertEqual(saved[0]['uploadDate'], '2026-10-29')
            originals, _ = parse_workbook(WORKBOOK.read_bytes())
            with server.connect() as connection:
                server.merge_items(connection, originals[:1])
                self.assertEqual(server.all_items(connection)[0]['uploader'], '甲')
                self.assertEqual(server.all_items(connection)[0]['uploadDate'], '2026-10-29')
            self.assertEqual(server.validate_item({**saved[0], 'uploaded': 0})['uploader'], '')
            self.assertEqual(server.validate_item({**saved[0], 'uploaded': 0})['uploadDate'], '')

    def test_customer_export_added_deleted_and_special_folders(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        first = records[0]
        extra = {**first, 'name': '=新文件夹', 'total': 4, 'uploaded': 2, 'uploader': '新担当', 'uploadDate': '2026-11-01'}
        new_group = {**extra, 'group': '新增父文件夹', 'name': '子文件夹'}
        customer = load_workbook(io.BytesIO(server.export_customer_workbook([first, extra, new_group]))).active
        self.assertEqual(customer.max_row, 6)
        self.assertEqual(str(customer.merged_cells), 'B3:B4')
        self.assertEqual(customer['C4'].value, '=新文件夹')
        self.assertEqual(customer['C4'].data_type, 's')
        self.assertEqual((customer['E4'].value, customer['F4'].value, customer['G4'].value), (datetime(2026, 11, 1), '新担当', 2))
        self.assertEqual(customer['B5'].value, '新增父文件夹')
        self.assertEqual(customer['D6'].value, '=SUM(D3:D5)')
        empty = load_workbook(io.BytesIO(server.export_customer_workbook([]))).active
        self.assertEqual(empty['D3'].value, '=0')

    def test_customer_export_matches_complete_example(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        group = '10要件定義\\10画面'
        uploads = {
            '02_SG総合メニュー': (6, '呉'), '04_KH介護保険業務': (63, '呉'),
            '05_IR医療保険業務': (22, '张晓静'), '06_SS障害者支援業務': (57, '张晓静'),
            '12_KE契約管理業務': (2, '呉')
        }
        for item in records:
            if item['group'] == group and item['name'] in uploads:
                item['uploaded'], item['uploader'] = uploads[item['name']]
                item['uploadDate'] = '2026-10-29'
        expected = load_workbook(server.CUSTOMER_TEMPLATE)
        actual = load_workbook(io.BytesIO(server.export_customer_workbook(records)))
        self.assertEqual(actual.sheetnames, expected.sheetnames)
        for expected_sheet, actual_sheet in zip(expected, actual):
            self.assertEqual(str(actual_sheet.merged_cells), str(expected_sheet.merged_cells))
            self.assertEqual((actual_sheet.max_row, actual_sheet.max_column), (expected_sheet.max_row, expected_sheet.max_column))
            for row in expected_sheet:
                for cell in row:
                    with self.subTest(sheet=expected_sheet.title, cell=cell.coordinate):
                        self.assertEqual(actual_sheet[cell.coordinate].value, cell.value)
                        self.assertEqual(actual_sheet[cell.coordinate]._style, cell._style)

    def test_allocation_and_customer_http_endpoints(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[0], records[2], records[1]])
                first, second, _ = server.all_items(connection)
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        assignments = [{'id': first['id'], 'version': first['version'], 'uploader': '甲', 'uploadDate': '2026-11-02'},
                                       {'id': second['id'], 'version': second['version'], 'uploader': '乙'}]
                        def confirm():
                            request = Request(f'{url}/api/upload-manager/allocations/confirm', data=json.dumps(assignments).encode(),
                                              headers={'Content-Type': 'application/json'}, method='POST')
                            return urlopen(request)
                        with confirm() as response:
                            saved = json.load(response)
                        self.assertEqual([item['uploader'] for item in saved], ['甲', '乙'])
                        self.assertEqual([item['plannedDate'] for item in saved], ['2026-11-02', date.today().isoformat()])
                        self.assertEqual([item['uploadDate'] for item in saved], ['', ''])
                        self.assertEqual([item['status'] for item in saved], ['uploading', 'uploading'])
                        with self.assertRaises(HTTPError) as error:
                            confirm()
                        self.assertEqual(error.exception.code, 409)
                        error.exception.close()
                        with urlopen(f'{url}/api/upload-manager/export/customer') as response:
                            customer = load_workbook(io.BytesIO(response.read())).active
                        self.assertEqual(customer.max_row, 6)
                        self.assertEqual([customer.cell(row, 5).value for row in (3, 5)], [None, None])
                        for item in saved:
                            payload = {**item, 'statusOverride': 'uploaded', 'uploaded': item['total']}
                            request = Request(f"{url}/api/upload-manager/items/{item['id']}", data=json.dumps(payload).encode(),
                                              headers={'Content-Type': 'application/json'}, method='PUT')
                            with urlopen(request) as response:
                                updated = json.load(response)
                            self.assertEqual(updated['status'], 'uploaded')
                            self.assertEqual(updated['plannedDate'], item['plannedDate'])
                        with urlopen(f'{url}/api/upload-manager/export/customer') as response:
                            customer = load_workbook(io.BytesIO(response.read())).active
                        self.assertEqual([customer.cell(row, 6).value for row in (3, 5)], ['甲', '乙'])
                        self.assertEqual(customer.cell(3, 5).value.date(), date.today())
                        self.assertEqual(customer.cell(5, 5).value.date(), date.today())
                    finally:
                        http_server.shutdown()
                        thread.join()

    def test_original_workbook(self):
        records, sheet = parse_workbook(WORKBOOK.read_bytes())
        self.assertEqual(sheet, 'Sheet1')
        self.assertEqual(len(records), 84)
        self.assertEqual(sum(item['total'] for item in records), 3490)
        self.assertEqual(records[1]['processed'], 59)
        self.assertFalse(records[1]['complete'])
        self.assertEqual(records[1]['uploaded'], 0)
        self.assertTrue(records[0]['complete'])
        self.assertTrue(records[40]['parentOnly'])

    def test_export_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0]['uploaded'] = records[0]['total']
        data = export_workbook(records)
        reloaded, _ = parse_workbook(data)
        self.assertEqual(len(reloaded), len(records))
        self.assertEqual(reloaded[0]['uploaded'], records[0]['total'])
        self.assertEqual(reloaded[1]['processed'], 59)
        self.assertEqual(reloaded[40]['parentOnly'], records[40]['parentOnly'])
        self.assertEqual(load_workbook(io.BytesIO(data)).active['A1'].value, '親フォルダ')

    def test_explicit_incomplete_preserved(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        records[0]['complete'] = False
        reloaded, _ = parse_workbook(export_workbook(records))
        self.assertFalse(reloaded[0]['complete'])

    def test_edited_values_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        edited = {**records[1], 'total': 145, 'totalMissing': False, 'processed': 145,
                  'complete': True, 'uploaded': 10, 'uploadedProvided': True,
                  'date': '2026-11-01', 'owner': '测试'}
        parent = {**records[40], 'total': 0, 'totalMissing': True, 'processed': None,
                  'complete': False, 'uploaded': 0, 'uploadedProvided': True}
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [edited, parent])
                saved = server.all_items(connection)
            reloaded, _ = parse_workbook(export_workbook(saved))
        for field in ('group', 'name', 'total', 'totalMissing', 'processed', 'complete', 'uploaded', 'date', 'owner', 'parentOnly'):
            self.assertEqual(reloaded[0][field], edited[field])
            self.assertEqual(reloaded[1][field], parent[field])

    def test_edit_api_exports_updated_counts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [records[1]])
                item = server.all_items(connection)[0]
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        item.update(processed=item['total'], complete=True, uploaded=10)
                        url = f'http://127.0.0.1:{server.PORT}'
                        request = Request(f'{url}/api/upload-manager/items/{item["id"]}', data=json.dumps(item).encode(),
                                          headers={'Content-Type': 'application/json'}, method='PUT')
                        with urlopen(request) as response:
                            saved = json.load(response)
                        with urlopen(f'{url}/api/upload-manager/export') as response:
                            exported, _ = parse_workbook(response.read())
                        reset_payload = {**saved, 'uploaded': 0}
                        request = Request(f'{url}/api/upload-manager/items/{saved["id"]}', data=json.dumps(reset_payload).encode(),
                                          headers={'Content-Type': 'application/json'}, method='PUT')
                        with urlopen(request) as response:
                            reset = json.load(response)
                    finally:
                        http_server.shutdown()
                        thread.join()
            self.assertEqual((saved['processed'], saved['complete'], saved['uploaded']), (145, True, 10))
            self.assertEqual((exported[0]['processed'], exported[0]['complete'], exported[0]['uploaded']), (145, True, 10))
            self.assertEqual(saved['uploadDate'], date.today().isoformat())
            self.assertEqual(exported[0]['uploadDate'], saved['uploadDate'])
            self.assertEqual(reset['uploadDate'], '')
            self.assertEqual(reset['uploader'], '')

    def test_import_preserves_upload_and_rolls_back_conflicts(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                self.assertEqual(server.merge_items(connection, records), (84, 0))
                connection.execute('UPDATE items SET uploaded=total WHERE rowid=1')
            with server.connect() as connection:
                self.assertEqual(server.merge_items(connection, records), (0, 84))
                self.assertEqual(server.all_items(connection)[0]['uploaded'], 2)
            records[1]['owner'] = 'changed'
            records[0]['complete'] = False
            with self.assertRaises(ValueError):
                with server.connect() as connection:
                    server.merge_items(connection, [records[1], records[0], *records[2:]])
            with server.connect() as connection:
                self.assertEqual(server.all_items(connection)[1]['owner'], '剛')

    def test_json_export_and_import_round_trip(self):
        records, _ = parse_workbook(WORKBOOK.read_bytes())
        first = {**records[0], 'uploaded': records[0]['total'], 'uploader': '甲', 'uploadDate': '2026-10-29', 'statusOverride': 'uploaded'}
        second = {**records[1], 'processed': 145, 'complete': True, 'plannedDate': '2026-11-02', 'uploader': '乙', 'statusOverride': 'uploading'}
        third = records[40]
        with tempfile.TemporaryDirectory() as directory, patch.object(server, 'DATABASE', Path(directory) / 'test.sqlite3'):
            server.init_db()
            with server.connect() as connection:
                server.merge_items(connection, [first, second, third])
            with server.LocalHTTPServer(('127.0.0.1', 0), server.Handler) as http_server:
                with patch.object(server, 'PORT', http_server.server_address[1]):
                    thread = threading.Thread(target=http_server.serve_forever, daemon=True)
                    thread.start()
                    try:
                        url = f'http://127.0.0.1:{server.PORT}'
                        with urlopen(f'{url}/api/upload-manager/export/json') as response:
                            self.assertEqual(response.headers.get_content_type(), 'application/json')
                            self.assertIn('upload-desk-backup.json', response.headers.get('Content-Disposition', ''))
                            exported = json.load(response)
                        self.assertEqual(len(exported), 3)
                        with server.connect() as connection:
                            connection.execute('DELETE FROM items WHERE name=?', (second['name'],))
                            connection.execute("UPDATE items SET owner='旧负责人', uploaded=0, uploader='', upload_date='', status_override='' WHERE name=?", (first['name'],))
                        exported[0]['owner'] = '新负责人'
                        request = Request(f'{url}/api/upload-manager/migrate', data=json.dumps(exported).encode('utf-8'),
                                          headers={'Content-Type': 'application/json'}, method='POST')
                        with urlopen(request) as response:
                            self.assertEqual(json.load(response), {'inserted': 1, 'updated': 2})
                        with urlopen(f'{url}/api/upload-manager/export/json') as response:
                            restored = {item['name']: item for item in json.load(response)}
                    finally:
                        http_server.shutdown()
                        thread.join()
            self.assertEqual(restored[first['name']]['owner'], '新负责人')
            self.assertEqual(restored[first['name']]['uploaded'], first['total'])
            self.assertEqual(restored[first['name']]['uploader'], '甲')
            self.assertEqual(restored[first['name']]['uploadDate'], '2026-10-29')
            self.assertEqual(restored[first['name']]['status'], 'uploaded')
            self.assertEqual(restored[second['name']]['plannedDate'], '2026-11-02')
            self.assertEqual(restored[second['name']]['uploader'], '乙')
            self.assertEqual(restored[second['name']]['status'], 'uploading')
            self.assertTrue(restored[third['name']]['parentOnly'])


if __name__ == '__main__':
    unittest.main()
