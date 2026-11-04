"""Local page that collects choices and writes the PowerPoint file."""

import json
import traceback
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

from power_report.assemble import assemble
from power_report.model import Report
from power_report.pptx_build import write_pptx
from power_report.selection import (
    choices_from_payload,
    load_saved_groups,
    page_model,
    save_choices,
    selection_path,
)

PAGE = (Path(__file__).with_name("picker.html")).read_text(encoding="utf-8")


def serve(report: Report, workbook: Path, open_browser: bool = True) -> int:
    server = make_server(report, workbook)
    url = f"http://127.0.0.1:{server.server_address[1]}/"
    print(f"已讀取：{report.source_name}")
    print(f"勾選頁：{url}")
    print("按 Ctrl+C 結束。")
    if open_browser:
        try:
            webbrowser.open(url)
        except OSError:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n已結束。")
    finally:
        server.server_close()
    return 0


def make_server(report: Report, workbook: Path) -> ThreadingHTTPServer:
    handler = _handler_class(report, workbook.resolve())
    return ThreadingHTTPServer(("127.0.0.1", 0), handler)


def _handler_class(report: Report, workbook: Path):
    class Handler(BaseHTTPRequestHandler):
        def do_GET(self):
            if self.path.split("?", 1)[0] not in {"/", "/index.html"}:
                self.send_error(404)
                return
            model = page_model(report, load_saved_groups(selection_path(workbook)))
            body = PAGE.replace("__REPORT_JSON__", _json_for_html(model)).encode("utf-8")
            self._send(200, "text/html; charset=utf-8", body)

        def do_POST(self):
            if self.path.split("?", 1)[0] != "/api/generate":
                self.send_error(404)
                return
            length = int(self.headers.get("Content-Length", "0"))
            if length > 2_000_000:
                self._send_json(413, {"ok": False, "error": "勾選內容過大"})
                return
            raw = self.rfile.read(length)
            try:
                payload = json.loads(raw.decode("utf-8"))
                choices = choices_from_payload(report, payload)
                deck = assemble(report, choices)
                pptx_path = workbook.with_suffix(".pptx")
                saved_path = selection_path(workbook)
                write_pptx(deck, pptx_path)
                save_choices(saved_path, report, choices)
            except Exception as exc:
                traceback.print_exc()
                self._send_json(500, {"ok": False, "error": str(exc)})
                return
            self._send_json(
                200,
                {"ok": True, "path": str(pptx_path), "selectionPath": str(saved_path)},
            )

        def log_message(self, fmt, *args):
            print("[%s] %s" % (self.log_date_time_string(), fmt % args))

        def _send(self, status: int, content_type: str, body: bytes) -> None:
            self.send_response(status)
            self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def _send_json(self, status: int, payload: dict) -> None:
            self._send(status, "application/json; charset=utf-8", json.dumps(payload).encode("utf-8"))

    return Handler


def _json_for_html(payload: dict) -> str:
    return json.dumps(payload, ensure_ascii=False).replace("<", "\\u003c")
