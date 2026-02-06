import argparse
import json
import re
import sys
import tempfile
import webbrowser
from datetime import datetime
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
SRC = ROOT / "src"
if SRC.exists() and str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from md2word import convert
from md2word.config import Config, DEFAULT_CONFIG

HERE = Path(__file__).resolve().parent
if hasattr(sys, "_MEIPASS"):
    INDEX_PATH = Path(sys._MEIPASS) / "index.html"
else:
    INDEX_PATH = HERE / "index.html"
DEFAULT_CONFIG_FULL = Config.from_dict(DEFAULT_CONFIG).to_dict()
LOG_PATH = Path(tempfile.gettempdir()) / "md2word_demo.log"


def _safe_filename(name: str) -> str:
    name = name.strip().replace("\\", "/")
    name = name.split("/")[-1]
    name = re.sub(r"[^A-Za-z0-9._-]+", "_", name)
    return name or "output.docx"


class Handler(BaseHTTPRequestHandler):
    server_version = "md2word-demo/0.1"

    def do_GET(self):
        try:
            _log(f"GET {self.path}")
            if self.path in ("/", "/index.html"):
                self._serve_index()
                return
            if self.path == "/health":
                self._send_text("ok", status=HTTPStatus.OK)
                return
            if self.path == "/default-config":
                self._send_json(DEFAULT_CONFIG_FULL, status=HTTPStatus.OK)
                return
            self._send_text("Not Found", status=HTTPStatus.NOT_FOUND)
        except Exception as exc:
            _log(f"GET error: {exc!r}")
            self._send_text("Internal Server Error", status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def do_POST(self):
        try:
            _log(f"POST {self.path}")
            if self.path != "/convert":
                self._send_text("Not Found", status=HTTPStatus.NOT_FOUND)
                return

            content_length = int(self.headers.get("Content-Length", "0"))
            if content_length <= 0:
                self._send_text("Missing request body", status=HTTPStatus.BAD_REQUEST)
                return

            raw = self.rfile.read(content_length)
            try:
                payload = json.loads(raw.decode("utf-8"))
            except json.JSONDecodeError:
                self._send_text("Invalid JSON", status=HTTPStatus.BAD_REQUEST)
                return

            markdown = (payload.get("markdown") or "").strip()
            if not markdown:
                self._send_text("Markdown content is empty", status=HTTPStatus.BAD_REQUEST)
                return

            toc = bool(payload.get("toc"))
            toc_title = (payload.get("toc_title") or "目录").strip() or "目录"
            try:
                toc_level = int(payload.get("toc_level", 3))
            except (TypeError, ValueError):
                toc_level = 3
            toc_level = max(1, min(9, toc_level))

            filename = _safe_filename(payload.get("filename") or "output.docx")
            if not filename.lower().endswith(".docx"):
                filename = f"{filename}.docx"

            config = None
            config_payload = payload.get("config")
            if isinstance(config_payload, dict):
                try:
                    config = Config.from_dict(config_payload)
                except Exception as exc:
                    self._send_text(f"Invalid config: {exc}", status=HTTPStatus.BAD_REQUEST)
                    return
            elif isinstance(config_payload, str) and config_payload.strip():
                try:
                    config_dict = json.loads(config_payload)
                    if not isinstance(config_dict, dict):
                        raise ValueError("Config JSON must be an object")
                    config = Config.from_dict(config_dict)
                except Exception as exc:
                    self._send_text(f"Invalid config: {exc}", status=HTTPStatus.BAD_REQUEST)
                    return

            try:
                with tempfile.TemporaryDirectory() as tmp_dir:
                    output_path = Path(tmp_dir) / filename
                    convert(
                        markdown,
                        output_path,
                        config=config,
                        toc=toc,
                        toc_title=toc_title,
                        toc_max_level=toc_level,
                    )
                    data = output_path.read_bytes()
            except Exception as exc:
                self._send_text(f"Conversion failed: {exc}", status=HTTPStatus.INTERNAL_SERVER_ERROR)
                return

            self.send_response(HTTPStatus.OK)
            self.send_header(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(len(data)))
            self.end_headers()
            self.wfile.write(data)
        except Exception as exc:
            _log(f"POST error: {exc!r}")
            self._send_text("Internal Server Error", status=HTTPStatus.INTERNAL_SERVER_ERROR)

    def _serve_index(self):
        if not INDEX_PATH.exists():
            self._send_text("index.html not found", status=HTTPStatus.NOT_FOUND)
            return
        data = INDEX_PATH.read_bytes()
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_text(self, text: str, status: HTTPStatus) -> None:
        data = text.encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def _send_json(self, payload: dict, status: HTTPStatus) -> None:
        data = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def log_message(self, fmt, *args):
        line = "%s - - [%s] %s\n" % (self.address_string(), self.log_date_time_string(), fmt % args)
        if sys.stderr:
            sys.stderr.write(line)
        else:
            _log(line.strip())


def main() -> int:
    parser = argparse.ArgumentParser(description="Minimal md2word desktop demo (local web UI)")
    parser.add_argument("--host", default="127.0.0.1", help="Bind host (default: 127.0.0.1)")
    parser.add_argument("--port", type=int, default=7860, help="Bind port (default: 7860)")
    parser.add_argument("--no-browser", action="store_true", help="Do not open browser automatically")
    args = parser.parse_args()

    try:
        server = ThreadingHTTPServer((args.host, args.port), Handler)
    except OSError as exc:
        _show_error(f"无法启动服务：端口 {args.port} 可能被占用。\n{exc}")
        return 1

    url = f"http://{args.host}:{args.port}"
    print(f"[INFO] Server running: {url}")
    if not args.no_browser:
        try:
            webbrowser.open(url)
        except Exception:
            pass
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n[INFO] Shutting down")
    finally:
        server.server_close()

    return 0


def _show_error(message: str) -> None:
    try:
        import ctypes

        ctypes.windll.user32.MessageBoxW(0, message, "md2word demo", 0x10)
    except Exception:
        print(f"[ERROR] {message}")


def _log(message: str) -> None:
    try:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        LOG_PATH.parent.mkdir(parents=True, exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{timestamp}] {message}\n")
    except Exception:
        pass


if __name__ == "__main__":
    raise SystemExit(main())
