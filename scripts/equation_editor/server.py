#!/usr/bin/env python3
"""
PPT Master - Reduction-Project Equation Web Editor (local server)

Serves the offline web editor for `monitoring_plan.py`: the user types formulas
as rough text, sees them as rendered equations with the auto-matched
"기호 | 정의 | 단위" tables and parameter grid, and downloads a Word or
PowerPoint file in which every formula and symbol is native, editable Office
Math (OMML). All parsing and file writing is done by the Python modules, so
the browser preview and the exported files come from one source.

Endpoints:
    GET  /                      editor page (static/)
    GET  /api/example           sample plan
    GET  /api/initial           plan given with --plan (or null)
    POST /api/derive            {text, parameters, title, font, options} -> {plan, derived}
    POST /api/latex             {text} -> {latex, mathml} | {error}
    POST /api/export/docx       {plan} -> .docx download (400 + issues on errors)
    POST /api/export/pptx       {plan} -> .pptx download
    POST /api/import            raw .docx/.pptx bytes (header X-Filename) -> {plan, text}

Usage:
    python3 scripts/equation_editor/server.py [--port 8095] [--plan plan.json] [--no-browser]

Examples:
    python3 scripts/equation_editor/server.py --plan projects/demo/notes/plan.json

Dependencies:
    python-pptx (PowerPoint export only)
"""

from __future__ import annotations

import argparse
import json
import mimetypes
import sys
import tempfile
import threading
import webbrowser
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any, Optional
from urllib.parse import quote, unquote

_HERE = Path(__file__).resolve().parent
_SCRIPTS_DIR = _HERE.parent
if str(_SCRIPTS_DIR) not in sys.path:
    sys.path.insert(0, str(_SCRIPTS_DIR))

from console_encoding import configure_utf8_stdio  # noqa: E402
from math_shorthand import to_latex  # noqa: E402
from monitoring_plan import (  # noqa: E402
    EXAMPLE_TEXT,
    derive,
    import_office,
    new_plan,
    plan_from_text,
    plan_to_text,
    write_docx,
    write_pptx,
)
from omml_math import EquationError, mathml  # noqa: E402
from server_common import find_free_port  # noqa: E402

STATIC = _HERE / "static"
MAX_UPLOAD = 60 * 1024 * 1024
EXPORT_NAMES = {"docx": "수식_모니터링계획.docx", "pptx": "수식_모니터링계획.pptx"}
EXPORT_TYPES = {
    "docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
}
INITIAL: dict[str, Any] = {"plan": None}


def _plan_from_request(body: dict[str, Any]) -> dict[str, Any]:
    plan = new_plan()
    for key in ("title", "font", "font_size_pt"):
        if body.get(key) not in (None, ""):
            plan[key] = body[key]
    plan["options"] = {**plan["options"], **(body.get("options") or {})}
    plan["parameters"] = [dict(p) for p in body.get("parameters") or [] if p.get("symbol")]
    return plan_from_text(body.get("text") or "", plan)


def _public_derived(derived: dict[str, Any]) -> dict[str, Any]:
    return {
        "sections": derived["sections"],
        "parameters": derived["parameters"],
        "issues": derived["issues"],
        "options": derived["options"],
    }


def _export_plan(body: dict[str, Any]) -> dict[str, Any]:
    plan = body.get("plan") or body
    if "sections" not in plan and "text" in plan:
        plan = _plan_from_request(plan)
    return plan


class Handler(BaseHTTPRequestHandler):
    server_version = "EquationEditor/1.0"

    def log_message(self, fmt: str, *args: Any) -> None:  # quiet console
        pass

    # -- helpers ---------------------------------------------------------
    def _send(self, status: int, data: bytes, ctype: str, extra: Optional[dict[str, str]] = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        for key, value in (extra or {}).items():
            self.send_header(key, value)
        self.end_headers()
        self.wfile.write(data)

    def _json(self, payload: Any, status: int = 200) -> None:
        self._send(status, json.dumps(payload, ensure_ascii=False).encode("utf-8"),
                   "application/json; charset=utf-8")

    def _body(self) -> bytes:
        length = int(self.headers.get("Content-Length") or 0)
        if length > MAX_UPLOAD:
            raise ValueError("파일이 너무 큽니다 (60MB 초과)")
        return self.rfile.read(length) if length else b""

    def _json_body(self) -> dict[str, Any]:
        raw = self._body()
        return json.loads(raw.decode("utf-8")) if raw else {}

    # -- GET -------------------------------------------------------------
    def do_GET(self) -> None:  # noqa: N802
        path = unquote(self.path.split("?", 1)[0])
        if path == "/api/example":
            self._json({"plan": plan_from_text(EXAMPLE_TEXT), "text": EXAMPLE_TEXT})
            return
        if path == "/api/initial":
            plan = INITIAL["plan"]
            self._json({"plan": plan, "text": plan_to_text(plan) if plan else None})
            return
        rel = "index.html" if path in ("/", "") else path.lstrip("/")
        target = (STATIC / rel).resolve()
        if STATIC.resolve() not in target.parents or not target.is_file():
            self._send(404, b"not found", "text/plain")
            return
        ctype = mimetypes.guess_type(target.name)[0] or "application/octet-stream"
        if ctype.startswith("text/") or ctype.endswith("javascript"):
            ctype += "; charset=utf-8"
        self._send(200, target.read_bytes(), ctype)

    # -- POST ------------------------------------------------------------
    def do_POST(self) -> None:  # noqa: N802
        path = self.path.split("?", 1)[0]
        try:
            if path == "/api/derive":
                plan = _plan_from_request(self._json_body())
                self._json({"plan": plan, "derived": _public_derived(derive(plan))})
            elif path == "/api/latex":
                text = (self._json_body().get("text") or "").strip()
                try:
                    latex = to_latex(text)
                    self._json({"latex": latex, "mathml": mathml(latex)})
                except EquationError as exc:
                    self._json({"error": str(exc)})
            elif path in ("/api/export/docx", "/api/export/pptx"):
                self._export(path.rsplit("/", 1)[1])
            elif path == "/api/import":
                self._import()
            else:
                self._json({"error": "unknown endpoint"}, 404)
        except (ValueError, EquationError) as exc:
            self._json({"error": str(exc)}, 400)
        except Exception as exc:  # report, keep serving
            self._json({"error": f"{type(exc).__name__}: {exc}"}, 500)

    def _export(self, kind: str) -> None:
        derived = derive(_export_plan(self._json_body()))
        errors = [i for i in derived["issues"] if i["level"] == "error"]
        if errors:
            self._json({"error": "수식 오류를 먼저 고치세요", "issues": errors}, 400)
            return
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / f"out.{kind}"
            if kind == "docx":
                write_docx(derived, out)
            else:
                write_pptx(derived, out)
            data = out.read_bytes()
        name = EXPORT_NAMES[kind]
        disposition = f"attachment; filename=\"equations.{kind}\"; filename*=UTF-8''{quote(name)}"
        self._send(200, data, EXPORT_TYPES[kind], {"Content-Disposition": disposition})

    def _import(self) -> None:
        name = unquote(self.headers.get("X-Filename") or "upload.docx")
        suffix = Path(name).suffix.lower()
        if suffix not in (".docx", ".pptx"):
            raise ValueError(".docx 또는 .pptx 파일만 가져올 수 있습니다")
        data = self._body()
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / f"{Path(name).stem or 'upload'}{suffix}"
            src.write_bytes(data)
            plan = import_office(src)
        plan["title"] = Path(name).stem
        self._json({"plan": plan, "text": plan_to_text(plan)})


def main(argv: Optional[list[str]] = None) -> int:
    configure_utf8_stdio()
    parser = argparse.ArgumentParser(description="Local web editor for reduction-project equations.")
    parser.add_argument("--port", type=int, default=None)
    parser.add_argument("--plan", type=Path, default=None, help="plan JSON (or quick-text .txt) to open")
    parser.add_argument("--no-browser", action="store_true")
    args = parser.parse_args(argv)
    if args.plan:
        if args.plan.suffix.lower() in (".txt", ".md"):
            INITIAL["plan"] = plan_from_text(args.plan.read_text(encoding="utf-8"))
        else:
            INITIAL["plan"] = json.loads(args.plan.read_text(encoding="utf-8"))
    port = args.port or find_free_port(8095)
    httpd = ThreadingHTTPServer(("127.0.0.1", port), Handler)
    url = f"http://127.0.0.1:{port}/"
    print(f"수식·모니터링 인자 편집기: {url}  (종료: Ctrl+C)", flush=True)
    if not args.no_browser:
        threading.Timer(0.6, lambda: webbrowser.open(url)).start()
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
