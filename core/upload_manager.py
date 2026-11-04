"""Upload Desk integrated into step 9; Python/SQLite, no Node.js runtime.

Business rules and workbook handling recovered from rebuild.md with verified hashes.
HTTP routes are isolated under /api/upload-manager/. No separate server is started.
"""
import io
import json
import os
import socket
import sqlite3
import uuid
import threading

from core.common import PROJECT_ROOT
from core.data_dir import DATA_DIR
from contextlib import contextmanager
from copy import copy
from datetime import date, datetime
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

ROOT = PROJECT_ROOT
DATABASE = Path(os.environ.get('UPLOAD_DESK_DB', DATA_DIR / 'upload_manager.sqlite3'))
CUSTOMER_TEMPLATE = ROOT / 'data' / 'upload_manager' / 'customer-template.xlsx'
PORT = int(os.environ.get('PORT', '8000'))
MAX_BODY = 20 * 1024 * 1024
STATIC = {}


def list_source_files(parts):
    root = os.environ.get('SOURCE_ROOT', '').strip()
    if not root or not isinstance(parts, list) or not parts:
        return []
    if any(not isinstance(part, str) or part in ('', '.', '..') or any(char in '<>:"/\\|?*' for char in part) for part in parts):
        return []
    folder = Path(root).joinpath(*parts)
    if not folder.is_dir():
        return []
    return sorted(path.relative_to(folder).as_posix() for path in folder.rglob('*') if path.is_file())


class LocalHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = os.name != 'nt'

    def server_bind(self):
        if os.name == 'nt':
            self.socket.setsockopt(socket.SOL_SOCKET, socket.SO_EXCLUSIVEADDRUSE, 1)
        super().server_bind()


@contextmanager
def connect():
    DATABASE.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(DATABASE, timeout=10)
    connection.row_factory = sqlite3.Row
    try:
        with connection:
            yield connection
    finally:
        connection.close()


def init_db():
    with connect() as connection:
        connection.execute('''CREATE TABLE IF NOT EXISTS items (
            id TEXT PRIMARY KEY, parent TEXT NOT NULL, name TEXT NOT NULL,
            total INTEGER NOT NULL CHECK(total >= 0), total_missing INTEGER NOT NULL DEFAULT 0,
            start_date TEXT NOT NULL DEFAULT '', owner TEXT NOT NULL DEFAULT '',
            processed INTEGER, complete INTEGER NOT NULL DEFAULT 0,
            uploaded INTEGER NOT NULL DEFAULT 0, parent_only INTEGER NOT NULL DEFAULT 0,
            version INTEGER NOT NULL DEFAULT 1,
            UNIQUE(parent, name)
        )''')
        columns = {row['name'] for row in connection.execute('PRAGMA table_info(items)')}
        if 'uploader' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN uploader TEXT NOT NULL DEFAULT ''")
        if 'upload_date' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN upload_date TEXT NOT NULL DEFAULT ''")
        if 'planned_date' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN planned_date TEXT NOT NULL DEFAULT ''")
        if 'status_override' not in columns:
            connection.execute("ALTER TABLE items ADD COLUMN status_override TEXT NOT NULL DEFAULT ''")
        connection.execute('''CREATE TABLE IF NOT EXISTS allocation_history (
            id TEXT PRIMARY KEY, created_at TEXT NOT NULL, planned_date TEXT NOT NULL,
            label TEXT NOT NULL, assignments TEXT NOT NULL, archive BLOB NOT NULL,
            manifest TEXT NOT NULL DEFAULT ''
        )''')
        if 'manifest' not in {row['name'] for row in connection.execute('PRAGMA table_info(allocation_history)')}:
            connection.execute("ALTER TABLE allocation_history ADD COLUMN manifest TEXT NOT NULL DEFAULT ''")


def as_item(row):
    item = {'id': row['id'], 'group': row['parent'], 'name': row['name'],
            'total': row['total'], 'totalMissing': bool(row['total_missing']),
            'date': row['start_date'], 'owner': row['owner'], 'processed': row['processed'],
            'complete': bool(row['complete']), 'uploaded': row['uploaded'],
            'uploader': row['uploader'], 'uploadDate': row['upload_date'], 'plannedDate': row['planned_date'],
            'statusOverride': row['status_override'], 'parentOnly': bool(row['parent_only']), 'version': row['version']}
    item['status'] = folder_status(item)
    return item


def folder_status(item):
    if item.get('statusOverride'):
        return item['statusOverride']
    if item['uploaded'] and item['uploaded'] == item['total']:
        return 'uploaded'
    if item['uploaded'] or item.get('plannedDate'):
        return 'uploading'
    if not item.get('date'):
        return 'unstarted'
    if not item.get('totalMissing') and item.get('processed') is not None and item['processed'] == item['total']:
        return 'ready'
    return 'processing'


def all_items(connection):
    return [as_item(row) for row in connection.execute('SELECT * FROM items ORDER BY rowid')]


class AllocationConflict(ValueError):
    pass


def mark_allocations(connection, assignments):
    if not isinstance(assignments, list) or not assignments:
        raise ValueError('分配记录不能为空')
    seen = set()
    connection.execute('BEGIN IMMEDIATE')
    for assignment in assignments:
        if not isinstance(assignment, dict):
            raise ValueError('分配记录格式无效')
        item_id = assignment.get('id')
        uploader = assignment.get('uploader')
        if not isinstance(item_id, str) or not item_id or item_id in seen:
            raise ValueError('分配记录重复或标识无效')
        if not isinstance(uploader, str) or not uploader.strip() or len(uploader.strip()) > 100:
            raise ValueError('名义上传者无效')
        upload_date = assignment.get('uploadDate', date.today().isoformat())
        if not isinstance(upload_date, str) or not upload_date.strip():
            raise ValueError('上传日期不能为空')
        upload_date = parse_date(upload_date)
        version = integer(assignment.get('version'), '版本')
        seen.add(item_id)
        row = connection.execute('SELECT * FROM items WHERE id=?', (item_id,)).fetchone()
        if not row or folder_status(as_item(row)) != 'ready':
            raise AllocationConflict('分配中的文件夹已被修改或不再符合条件，请刷新并重新生成方案')
        result = connection.execute('''UPDATE items SET uploader=?,planned_date=?,complete=1,status_override='uploading',version=version+1
            WHERE id=? AND version=? AND uploaded=0 AND planned_date='' AND total>0''',
            (uploader.strip(), upload_date, item_id, version))
        if result.rowcount != 1:
            raise AllocationConflict('分配中的文件夹已被修改或不再符合条件，请刷新并重新生成方案')
    return [as_item(connection.execute('SELECT * FROM items WHERE id=?', (assignment['id'],)).fetchone()) for assignment in assignments]


def save_allocation_history(connection, payload):
    label = payload.get('label')
    if not isinstance(label, str) or not label.strip() or len(label) > 200:
        raise ValueError('方案名称无效')
    manifest = payload.get('manifest')
    if not isinstance(manifest, dict) or not isinstance(manifest.get('people'), list):
        raise ValueError('方案 JSON 格式无效')
    saved = mark_allocations(connection, payload.get('assignments'))
    dates = {item['plannedDate'] for item in saved}
    if len(dates) != 1:
        raise ValueError('同一方案的计划日期必须一致')
    history_id = str(uuid.uuid4())
    connection.execute('INSERT INTO allocation_history (id,created_at,planned_date,label,assignments,archive,manifest) VALUES (?,?,?,?,?,?,?)',
                       (history_id, datetime.now().isoformat(timespec='seconds'), dates.pop(), label.strip(),
                        json.dumps(saved, ensure_ascii=False), b'', json.dumps(manifest, ensure_ascii=False)))
    return saved, history_id


def integer(value, label, allow_empty=False):
    if value is None or value == '':
        if allow_empty:
            return None
        raise ValueError(f'{label}不能为空')
    if isinstance(value, bool) or not isinstance(value, (int, float)) or int(value) != value or not 0 <= value <= 1000000:
        raise ValueError(f'{label}必须是 0 到 1000000 的整数')
    return int(value)


def text(value):
    return str(value).strip() if value is not None else ''


def parse_date(value):
    if value is None or value == '':
        return ''
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    try:
        return date.fromisoformat(str(value).strip().replace('/', '-')).isoformat()
    except ValueError as error:
        raise ValueError(f'日期格式无效：{value}') from error


def validate_item(item):
    group, name = text(item.get('group')), text(item.get('name'))
    if not group or not name or len(group) > 500 or len(name) > 200:
        raise ValueError('父文件夹和子文件夹名称不能为空或过长')
    total = integer(item.get('total'), '文件数量')
    uploaded = integer(item.get('uploaded', 0), '已上传数量')
    processed = integer(item.get('processed'), '消化件数', allow_empty=True)
    complete = uploaded > 0 if item.get('complete') is None else item.get('complete')
    if not isinstance(complete, bool):
        raise ValueError('完成状态无效')
    if uploaded > total or (processed is not None and processed > total):
        raise ValueError('已上传数量或消化件数不能大于文件数量')
    if uploaded and not complete:
        raise ValueError('已有上传数量的文件夹必须为已完成')
    owner = text(item.get('owner'))
    if len(owner) > 100:
        raise ValueError('负责人名称过长')
    planned_date = parse_date(item.get('plannedDate'))
    status_override = item.get('statusOverride', '') or ''
    if status_override not in ('', 'unstarted', 'processing', 'ready', 'uploading', 'uploaded'):
        raise ValueError('文件夹状态无效')
    if status_override == 'uploaded' and (not total or uploaded != total):
        raise ValueError('已上传状态要求全部文件已上传')
    if status_override in ('unstarted', 'processing', 'ready') and (uploaded or planned_date):
        raise ValueError('请先取消上传计划或上传状态')
    if status_override == 'uploading' and total and uploaded == total:
        raise ValueError('上传中状态不能是全部文件已上传')
    if (planned_date or status_override in ('ready', 'uploading', 'uploaded')) and not complete:
        raise ValueError('已完成或上传中的文件夹必须为已完成')
    uploader = text(item.get('uploader')) if uploaded or planned_date or status_override == 'uploading' else ''
    if len(uploader) > 100:
        raise ValueError('名义上传者名称过长')
    return {'group': group, 'name': name, 'total': total, 'totalMissing': bool(item.get('totalMissing', False)),
            'date': parse_date(item.get('date')), 'owner': owner, 'processed': processed,
            'complete': complete, 'uploaded': uploaded, 'uploader': uploader,
            'uploadDate': parse_date(item.get('uploadDate')) if uploaded else '', 'plannedDate': planned_date,
            'statusOverride': status_override,
            'parentOnly': bool(item.get('parentOnly', False))}


def parse_workbook(data):
    try:
        workbook = load_workbook(io.BytesIO(data), data_only=False)
    except Exception as error:
        raise ValueError('无法读取 Excel 文件，请选择有效的 .xlsx 文件') from error
    required = {'親フォルダ', '子フォルダ', 'ファイル数', '開始日', '担当者', '消化件数', '完了'}
    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30)):
            headers = {text(cell.value): cell.column for cell in row if isinstance(cell.value, str)}
            if required.issubset(headers):
                break
        else:
            continue
        break
    else:
        raise ValueError('未找到包含親フォルダ、子フォルダ、ファイル数等列的工作表')
    parent_col = headers['親フォルダ']
    merged_parents = {}
    for merged in sheet.merged_cells.ranges:
        if merged.min_col == merged.max_col == parent_col:
            parent = sheet.cell(merged.min_row, parent_col).value
            for row_number in range(merged.min_row, merged.max_row + 1):
                merged_parents[row_number] = parent
    results = []
    paths = set()
    for row_number in range(row[0].row + 1, sheet.max_row + 1):
        def cell(column):
            return sheet.cell(row_number, headers[column]).value
        group = text(merged_parents.get(row_number, cell('親フォルダ')))
        child = text(cell('子フォルダ'))
        if not group:
            if child and not any(word in child for word in ('総件数', '合計')):
                raise ValueError(f'第 {row_number} 行缺少父文件夹')
            continue
        try:
            total_value = integer(cell('ファイル数'), '文件数量', allow_empty=True)
            total = total_value if total_value is not None else 0
            processed = integer(cell('消化件数'), '消化件数', allow_empty=True)
            marker = cell('完了')
            complete = (text(marker) in ('〇', '○', '已完成', '完成', '✓', '✔', 'TRUE', 'True') or marker is True)
            if marker is None or (isinstance(marker, str) and marker.startswith('=')):
                complete = total_value is not None and processed is not None and total == processed
            upload_column = next((name for name in ('已上传数量', 'アップロード数') if name in headers), None)
            uploaded = integer(cell(upload_column), '已上传数量', allow_empty=True) if upload_column else None
            parent_only = not child
            name = child or group.replace('¥', '\\').replace('/', '\\').split('\\')[-1]
            parsed = validate_item({'group': group, 'name': name, 'total': total, 'totalMissing': total_value is None,
                                    'date': parse_date(cell('開始日')), 'owner': text(cell('担当者')),
                                    'processed': processed, 'complete': complete, 'uploaded': uploaded or 0,
                                    'uploader': text(cell('名义上传者')) if '名义上传者' in headers else '',
                                    'uploadDate': parse_date(cell('上传日期')) if '上传日期' in headers else '',
                                    'plannedDate': parse_date(cell('计划上传日期')) if '计划上传日期' in headers else '',
                                    'statusOverride': text(cell('手动状态')) if '手动状态' in headers else '',
                                    'parentOnly': parent_only})
            parsed['uploadedProvided'] = upload_column is not None
            parsed['uploaderProvided'] = '名义上传者' in headers
            parsed['uploadDateProvided'] = '上传日期' in headers
            parsed['plannedDateProvided'] = '计划上传日期' in headers
            parsed['statusOverrideProvided'] = '手动状态' in headers
        except ValueError as error:
            raise ValueError(f'第 {row_number} 行：{error}') from error
        key = (parsed['group'], parsed['name'])
        if key in paths:
            raise ValueError(f'第 {row_number} 行的文件夹重复：{group} / {name}')
        paths.add(key)
        results.append(parsed)
    if not results:
        raise ValueError('工作表没有可导入的文件夹数据')
    return results, sheet.title


def merge_items(connection, records, preserve_upload=False):
    inserted = updated = 0
    for record in records:
        if not isinstance(record, dict):
            raise ValueError('备份数据格式不正确')
        existing = connection.execute('SELECT * FROM items WHERE parent=? AND name=?', (text(record.get('group')), text(record.get('name')))).fetchone()
        if existing and (preserve_upload or not record.get('uploadedProvided', True)):
            record = {**record, 'uploaded': existing['uploaded']}
        if existing and (preserve_upload or not record.get('uploaderProvided', 'uploader' in record)):
            record = {**record, 'uploader': existing['uploader']}
        if existing and (preserve_upload or not record.get('uploadDateProvided', 'uploadDate' in record)):
            record = {**record, 'uploadDate': existing['upload_date']}
        if existing and (preserve_upload or not record.get('plannedDateProvided', 'plannedDate' in record)):
            record = {**record, 'plannedDate': existing['planned_date']}
        if existing and (preserve_upload or not record.get('statusOverrideProvided', 'statusOverride' in record)):
            record = {**record, 'statusOverride': existing['status_override']}
        if preserve_upload and existing and existing['uploaded'] and (record['total'] < existing['uploaded'] or not record['complete']):
            raise ValueError(f'{record["group"]} / {record["name"]}：表格中的文件数量或完成状态与已有上传记录冲突，未导入任何记录')
        try:
            record = validate_item(record)
        except ValueError as error:
            raise ValueError(f'{record.get("group", "")} / {record.get("name", "")}：{error}') from error
        values = (record['group'], record['name'], record['total'], int(record['totalMissing']), record['date'], record['owner'],
                record['processed'], int(record['complete']), record['uploaded'], int(record['parentOnly']), record['uploader'], record['uploadDate'], record['plannedDate'], record['statusOverride'])
        if existing:
            connection.execute('UPDATE items SET parent=?,name=?,total=?,total_missing=?,start_date=?,owner=?,processed=?,complete=?,uploaded=?,parent_only=?,uploader=?,upload_date=?,planned_date=?,status_override=?,version=version+1 WHERE id=?', (*values, existing['id']))
            updated += 1
        else:
            connection.execute('INSERT INTO items (parent,name,total,total_missing,start_date,owner,processed,complete,uploaded,parent_only,uploader,upload_date,planned_date,status_override,id) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)', (*values, str(uuid.uuid4())))
            inserted += 1
    return inserted, updated


def export_workbook(records):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = '文件夹明细'
    sheet.append(['親フォルダ', '子フォルダ', 'ファイル数', '開始日', '担当者', '消化件数', '完了', '已上传数量', '上传状态', '名义上传者', '上传日期', '计划上传日期', '文件夹状态', '手动状态'])
    for item in records:
        sheet.append([item['group'], '' if item['parentOnly'] else item['name'], None if item['totalMissing'] else item['total'],
                      item['date'], item['owner'], item['processed'], '〇' if item['complete'] else '未完成', item['uploaded'],
                      '已上传' if item['total'] > 0 and item['uploaded'] == item['total'] else '部分上传' if item['uploaded'] else '未上传', item.get('uploader', ''), item.get('uploadDate', ''),
                      item.get('plannedDate', ''), {'unstarted':'未完成', 'processing':'完成中', 'ready':'已完成/未上传', 'uploading':'上传中', 'uploaded':'已上传'}[folder_status(item)], item.get('statusOverride', '')])
        for cell in sheet[sheet.max_row]:
            if isinstance(cell.value, str):
                cell.data_type = 's'
    for cell in sheet[1]:
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='28556A')
    for column, width in {'A': 35, 'B': 38, 'C': 13, 'D': 16, 'E': 16, 'F': 14, 'G': 12, 'H': 16, 'I': 14, 'J': 18, 'K': 16, 'L': 18, 'M': 20, 'N': 18}.items():
        sheet.column_dimensions[column].width = width
    sheet.freeze_panes = 'A2'
    sheet.auto_filter.ref = sheet.dimensions
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


def export_customer_workbook(records):
    if not CUSTOMER_TEMPLATE.is_file():
        raise FileNotFoundError(f'客户表模板不存在：{CUSTOMER_TEMPLATE}')
    workbook = load_workbook(CUSTOMER_TEMPLATE)
    sheet = workbook.active
    if [sheet.cell(2, column).value for column in range(2, 9)] != ['親フォルダ', '子フォルダ', 'ファイル数', '消化日', '担当者', '消化件数', '完了']:
        raise ValueError('客户表模板的表头不符合预期')
    footer = sheet.max_row
    if sheet.cell(footer, 3).value != '総件数：' or sheet.cell(footer, 6).value != '完了件数：':
        raise ValueError('客户表模板的汇总行不符合预期')

    merged_parents = {}
    for merged in list(sheet.merged_cells.ranges):
        if merged.min_col != 2 or merged.max_col != 2 or merged.min_row < 3 or merged.max_row >= footer:
            raise ValueError('客户表模板含有无法处理的合并单元格')
        for row in range(merged.min_row, merged.max_row + 1):
            merged_parents[row] = sheet.cell(merged.min_row, 2).value
    row_styles = {}
    parent_styles = {}
    group_order = []
    for row in range(3, footer):
        group = text(merged_parents.get(row, sheet.cell(row, 2).value))
        child = text(sheet.cell(row, 3).value)
        if not group:
            raise ValueError(f'客户表模板第 {row} 行缺少父文件夹')
        name = child or group.replace('¥', '\\').replace('/', '\\').split('\\')[-1]
        if group not in parent_styles:
            parent_styles[group] = copy(sheet.cell(row, 2)._style)
            group_order.append(group)
        row_styles[(group, name)] = ([copy(sheet.cell(row, col)._style) for col in range(1, 9)],
                                     sheet.row_dimensions[row].height, bool(sheet.cell(row, 8).value))
    default_styles = [copy(sheet.cell(5, col)._style) for col in range(1, 9)]
    default_parent_style = copy(sheet.cell(43, 2)._style)
    date_style = copy(sheet.cell(5, 5)._style)
    footer_styles = [copy(sheet.cell(footer, col)._style) for col in range(1, 9)]
    footer_height = sheet.row_dimensions[footer].height

    grouped = {group: [] for group in group_order}
    for item in records:
        grouped.setdefault(item['group'], []).append(item)
    for merged in list(sheet.merged_cells.ranges):
        sheet.unmerge_cells(str(merged))
    sheet.delete_rows(3, footer - 2)

    for group, items_in_group in grouped.items():
        if not items_in_group:
            continue
        by_name = {item['name']: item for item in items_in_group}
        ordered = [by_name.pop(name) for template_group, name in row_styles if template_group == group and name in by_name]
        ordered.extend(item for item in items_in_group if item['name'] in by_name)
        first_row = sheet.max_row + 1
        for item in ordered:
            row = sheet.max_row + 1
            styles, height, had_formula = row_styles.get((group, item['name']), (default_styles, None, True))
            for col, style in enumerate(styles, 1):
                sheet.cell(row, col)._style = copy(style)
            sheet.cell(row, 2)._style = copy(parent_styles.get(group, default_parent_style))
            sheet.row_dimensions[row].height = height
            values = [None, group if row == first_row else None, '' if item['parentOnly'] else item['name'],
                      None if item['totalMissing'] else item['total'],
                      date.fromisoformat(item['uploadDate']) if item['uploaded'] and item.get('uploadDate') else None,
                      ('静' if item.get('uploader') == '张晓静' else item.get('uploader', '')) if item['uploaded'] else None,
                      item['uploaded'] if item['uploaded'] else None]
            for col, value in enumerate(values, 1):
                cell = sheet.cell(row, col, value)
                if isinstance(value, str):
                    cell.data_type = 's'
            if values[4] is not None:
                sheet.cell(row, 5)._style = copy(date_style)
            if had_formula or item['uploaded']:
                sheet.cell(row, 8, f'=IF(AND(D{row}<>"",G{row}<>"",D{row}=G{row}),"〇","")')
        last_row = sheet.max_row
        if last_row > first_row:
            sheet.merge_cells(start_row=first_row, start_column=2, end_row=last_row, end_column=2)

    footer = sheet.max_row + 1
    for col, style in enumerate(footer_styles, 1):
        sheet.cell(footer, col)._style = copy(style)
    sheet.row_dimensions[footer].height = footer_height
    sheet.cell(footer, 3, '総件数：')
    sheet.cell(footer, 4, f'=SUM(D3:D{footer - 1})' if footer > 3 else '=0')
    sheet.cell(footer, 6, '完了件数：')
    sheet.cell(footer, 7, f'=SUM(G3:G{footer - 1})' if footer > 3 else '=0')
    sheet.auto_filter.ref = f'A2:K{footer}'
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True
    output = io.BytesIO()
    workbook.save(output)
    return output.getvalue()


class Handler(BaseHTTPRequestHandler):
    def respond(self, code, body, content_type='application/json; charset=utf-8', headers=None):
        if isinstance(body, (dict, list)):
            body = json.dumps(body, ensure_ascii=False).encode('utf-8')
        self.send_response(code)
        self.send_header('Content-Type', content_type)
        self.send_header('Content-Length', str(len(body)))
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        for name, value in (headers or {}).items():
            self.send_header(name, value)
        self.end_headers()
        self.wfile.write(body)

    def body(self):
        length = int(self.headers.get('Content-Length', '0'))
        if length <= 0 or length > MAX_BODY:
            raise ValueError('文件为空或超过 20 MB')
        return self.rfile.read(length)

    def trusted_origin(self):
        origin = self.headers.get('Origin')
        host = self.headers.get('Host', '')
        port = self.server.server_address[1]
        allowed = (f'127.0.0.1:{port}', f'localhost:{port}')
        return host in allowed and (not origin or origin in tuple(f'http://{host}' for host in allowed))

    def do_GET(self):
        if not self.trusted_origin():
            return self.respond(403, {'error': '请求来源不被允许'})
        path = urlparse(self.path).path
        if not path.startswith('/api/upload-manager/'):
            return self.respond(404, {'error': '未找到接口'})
        ensure_db()
        if path == '/api/upload-manager/health':
            return self.respond(200, {'service': 'upload-desk'})
        if path == '/api/upload-manager/items':
            with connect() as connection:
                return self.respond(200, all_items(connection))
        if path == '/api/upload-manager/allocation-history':
            with connect() as connection:
                rows = connection.execute('SELECT id,created_at,planned_date,label,assignments,manifest,length(archive) AS archive_size FROM allocation_history ORDER BY rowid DESC').fetchall()
                return self.respond(200, [{'id':row['id'], 'createdAt':row['created_at'], 'plannedDate':row['planned_date'],
                                           'label':row['label'], 'assignments':json.loads(row['assignments']),
                                           'manifest':json.loads(row['manifest']) if row['manifest'] else None,
                                           'hasArchive':bool(row['archive_size'])} for row in rows])
        if path.startswith('/api/upload-manager/allocation-history/') and path.endswith('/archive'):
            history_id = path.removeprefix('/api/upload-manager/allocation-history/').removesuffix('/archive')
            with connect() as connection:
                row = connection.execute('SELECT id,archive FROM allocation_history WHERE id=?', (history_id,)).fetchone()
            if row is None or not row['archive']:
                return self.respond(404, {'error': '历史方案不存在或没有 ZIP'})
            return self.respond(200, row['archive'], 'application/zip', {'Content-Disposition': f'attachment; filename="upload-allocation-{row["id"]}.zip"'})
        if path.startswith('/api/upload-manager/allocation-history/') and path.endswith('/manifest'):
            history_id = path.removeprefix('/api/upload-manager/allocation-history/').removesuffix('/manifest')
            with connect() as connection:
                row = connection.execute('SELECT manifest FROM allocation_history WHERE id=?', (history_id,)).fetchone()
            if row is None or not row['manifest']:
                return self.respond(404, {'error': '历史方案不存在或没有 JSON'})
            return self.respond(200, json.loads(row['manifest']))
        if path == '/api/upload-manager/export':
            with connect() as connection:
                data = export_workbook(all_items(connection))
            return self.respond(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', {'Content-Disposition': 'attachment; filename="upload-desk.xlsx"'})
        if path == '/api/upload-manager/export/json':
            with connect() as connection:
                data = json.dumps(all_items(connection), ensure_ascii=False, indent=2).encode('utf-8')
            return self.respond(200, data, 'application/json; charset=utf-8', {'Content-Disposition': 'attachment; filename="upload-desk-backup.json"'})
        if path == '/api/upload-manager/export/customer':
            try:
                with connect() as connection:
                    data = export_customer_workbook(all_items(connection))
            except (OSError, ValueError) as error:
                return self.respond(500, {'error': f'客户表导出失败：{error}'})
            return self.respond(200, data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', {'Content-Disposition': 'attachment; filename="uploaded-folders.xlsx"'})
        if path in STATIC:
            filename, content_type = STATIC[path]
            try:
                return self.respond(200, (ROOT / filename).read_bytes(), content_type)
            except FileNotFoundError:
                return self.respond(404, {'error': '资源不存在，请先安装依赖'})
        self.respond(404, {'error': '未找到资源'})

    def do_POST(self):
        self.write_request('POST')

    def do_PUT(self):
        self.write_request('PUT')

    def do_DELETE(self):
        self.write_request('DELETE')

    def write_request(self, method):
        if not self.trusted_origin():
            return self.respond(403, {'error': '请求来源不被允许'})
        path = urlparse(self.path).path
        if not path.startswith('/api/upload-manager/'):
            return self.respond(404, {'error': '未找到接口'})
        try:
            ensure_db()
            if method == 'POST' and path == '/api/upload-manager/file-lists':
                sources = json.loads(self.body()).get('sources')
                if not isinstance(sources, list):
                    raise ValueError('sources 格式无效')
                return self.respond(200, {'fileLists': [list_source_files(parts) for parts in sources]})
            if method == 'POST' and path == '/api/upload-manager/allocations/confirm':
                payload = json.loads(self.body())
                with connect() as connection:
                    if isinstance(payload, dict):
                        saved, history_id = save_allocation_history(connection, payload)
                        result = {'items':saved, 'historyId':history_id}
                    else:
                        result = mark_allocations(connection, payload)
                return self.respond(200, result)
            if method == 'POST' and path in ('/api/upload-manager/import/preview', '/api/upload-manager/import'):
                records, sheet = parse_workbook(self.body())
                with connect() as connection:
                    existing = {(row['parent'], row['name']) for row in connection.execute('SELECT parent,name FROM items')}
                    matched = sum((item['group'], item['name']) in existing for item in records)
                    if path.endswith('/preview'):
                        return self.respond(200, {'sheet': sheet, 'total': len(records), 'inserted': len(records) - matched, 'updated': matched})
                    inserted, updated = merge_items(connection, records, preserve_upload=True)
                return self.respond(200, {'inserted': inserted, 'updated': updated})
            if method == 'POST' and path == '/api/upload-manager/migrate':
                records = json.loads(self.body())
                if not isinstance(records, list) or not records:
                    raise ValueError('备份数据为空或格式不正确')
                with connect() as connection:
                    inserted, updated = merge_items(connection, records)
                return self.respond(200, {'inserted': inserted, 'updated': updated})
            if method == 'POST' and path == '/api/upload-manager/items':
                item = validate_item(json.loads(self.body()))
                with connect() as connection:
                    connection.execute('INSERT INTO items (id,parent,name,total,total_missing,start_date,owner,processed,complete,uploaded,parent_only,uploader,upload_date,planned_date,status_override) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)',
                                       (str(uuid.uuid4()),item['group'],item['name'],item['total'],int(item['totalMissing']),item['date'],item['owner'],item['processed'],int(item['complete']),item['uploaded'],int(item['parentOnly']),item['uploader'],date.today().isoformat() if item['uploaded'] else '',item['plannedDate'],item['statusOverride']))
                    row = connection.execute('SELECT * FROM items WHERE rowid=last_insert_rowid()').fetchone()
                    connection.commit()
                    return self.respond(201, as_item(row))
            if path.startswith('/api/upload-manager/items/') and method in ('PUT', 'DELETE'):
                item_id = path.removeprefix('/api/upload-manager/items/')
                payload = json.loads(self.body())
                version = integer(payload.get('version'), '版本')
                with connect() as connection:
                    if method == 'DELETE':
                        result = connection.execute('DELETE FROM items WHERE id=? AND version=?', (item_id, version))
                    else:
                        item = validate_item(payload)
                        result = connection.execute('''UPDATE items SET parent=?,name=?,total=?,total_missing=?,start_date=?,owner=?,processed=?,complete=?,
                            upload_date=CASE WHEN ?=0 THEN '' WHEN uploaded=0 THEN ? ELSE upload_date END,
                            uploaded=?,parent_only=?,uploader=?,planned_date=?,status_override=?,version=version+1 WHERE id=? AND version=?''',
                            (item['group'],item['name'],item['total'],int(item['totalMissing']),item['date'],item['owner'],item['processed'],int(item['complete']),
                             item['uploaded'],date.today().isoformat(),item['uploaded'],int(item['parentOnly']),item['uploader'],item['plannedDate'],item['statusOverride'],item_id,version))
                    if result.rowcount == 0:
                        return self.respond(409, {'error': '记录已被修改或删除，请刷新后重试'})
                    connection.commit()
                    if method == 'DELETE':
                        return self.respond(200, {'ok': True})
                    return self.respond(200, as_item(connection.execute('SELECT * FROM items WHERE id=?', (item_id,)).fetchone()))
            self.respond(404, {'error': '未找到接口'})
        except AllocationConflict as error:
            self.respond(409, {'error': str(error)})
        except (ValueError, KeyError, json.JSONDecodeError, sqlite3.IntegrityError) as error:
            self.respond(400, {'error': str(error) or '数据格式无效'})
        except Exception:
            self.respond(500, {'error': '服务器处理失败，请检查服务器日志'})
            raise



_init_lock = threading.RLock()
_initialized = set()


def ensure_db():
    with _init_lock:
        key = str(DATABASE.resolve())
        if key not in _initialized or not DATABASE.is_file():
            init_db()
            _initialized.add(key)
