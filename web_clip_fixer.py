import json
import mimetypes
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import urlparse

from core.common import STATIC_DIR, is_backup_sheet, sanitize_filename
from core.excel_ops import extract_sheet_snapshot, match_excel_row, normalize_workbook_view
from core.pdf_detector import (
    extract_page_lines,
    find_clipped_spans_on_page,
    is_region_still_clipped,
    render_crop_base64,
    render_page_overview,
)
from core.reporter import get_execution_report_snapshot, render_execution_markdown
from core.workflow import (
    STATE,
    STATE_LOCK,
    fix_selected_workflow,
    get_state_snapshot,
    log_msg,
    scan_folder_workflow,
    try_acquire_busy,
    update_item_row,
)

__all__ = [
    "STATE",
    "STATE_LOCK",
    "WebHandler",
    "extract_page_lines",
    "extract_sheet_snapshot",
    "find_clipped_spans_on_page",
    "fix_selected_workflow",
    "is_backup_sheet",
    "is_region_still_clipped",
    "log_msg",
    "main",
    "match_excel_row",
    "normalize_workbook_view",
    "render_crop_base64",
    "render_page_overview",
    "sanitize_filename",
    "scan_folder_workflow",
]

MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".svg": "image/svg+xml",
    ".png": "image/png",
}


class WebHandler(BaseHTTPRequestHandler):
    def log_message(self, format: str, *args: object) -> None:
        del format, args  # 保持终端清爽

    def _send_bytes(self, body: bytes, content_type: str, status: int = 200) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _send_json(self, data: dict, status: int = 200) -> None:
        body = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self._send_bytes(body, "application/json; charset=utf-8", status=status)

    def _serve_static_file(self, rel_path: str) -> None:
        static_root = STATIC_DIR.resolve()
        target = (static_root / rel_path).resolve()
        if not target.is_file() or static_root not in target.parents:
            self.send_error(404)
            return

        content_type = MIME_TYPES.get(
            target.suffix.lower(),
            mimetypes.guess_type(str(target))[0] or "application/octet-stream",
        )
        self._send_bytes(target.read_bytes(), content_type)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        if parsed.path in {"/", "/index.html"}:
            self._serve_static_file("index.html")
        elif parsed.path in {"/records", "/records.html"}:
            self._serve_static_file("records.html")
        elif parsed.path.startswith("/static/"):
            rel = parsed.path.removeprefix("/static/")
            self._serve_static_file(rel)
        elif parsed.path == "/api/state":
            self._send_json(get_state_snapshot())
        elif parsed.path == "/api/records":
            self._send_json(get_execution_report_snapshot())
        elif parsed.path == "/api/records/markdown":
            md_bytes = render_execution_markdown().encode("utf-8")
            self.send_response(200)
            self.send_header("Content-Type", "text/markdown; charset=utf-8")
            self.send_header("Content-Disposition", 'attachment; filename="execution_report.md"')
            self.send_header("Content-Length", str(len(md_bytes)))
            self.end_headers()
            self.wfile.write(md_bytes)
        else:
            self.send_error(404)

    def do_POST(self) -> None:
        length = int(self.headers.get("Content-Length", 0))
        payload = json.loads(self.rfile.read(length).decode("utf-8")) if length else {}

        if self.path == "/api/scan":
            if try_acquire_busy():
                threading.Thread(
                    target=scan_folder_workflow,
                    args=(payload.get("folder", ""), bool(payload.get("delete_backups", True))),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/fix":
            if try_acquire_busy():
                try:
                    step_pt = max(0.5, float(payload.get("step_pt", 3.0)))
                except (TypeError, ValueError):
                    step_pt = 3.0
                threading.Thread(
                    target=fix_selected_workflow,
                    args=(payload.get("selected_ids", []), step_pt),
                    daemon=True,
                ).start()
            self._send_json({"ok": True})

        elif self.path == "/api/update_row":
            iid = str(payload.get("id", ""))
            try:
                new_row = max(1, int(payload.get("excel_row", 1)))
            except (TypeError, ValueError):
                self._send_json({"ok": False, "error": "invalid excel_row"}, status=400)
                return

            updated = update_item_row(iid, new_row)
            self._send_json({"ok": updated})
        else:
            self.send_error(404)


def main() -> None:
    port = 8765
    server = ThreadingHTTPServer(("127.0.0.1", port), WebHandler)
    url = f"http://127.0.0.1:{port}"
    print(f"🚀 Web 界面已启动: {url}")
    print("正在自动打开浏览器... (按 Ctrl+C 可退出程序)")
    webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已退出。")


if __name__ == "__main__":
    main()

