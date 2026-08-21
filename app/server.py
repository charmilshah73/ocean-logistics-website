"""Local development server for the Logistics Control Tower."""
from __future__ import annotations

import io
import json
import shutil
from pathlib import Path

from flask import Flask, jsonify, request, send_file, send_from_directory

from app.admin_service import process_upload
from app.config import CURRENT, JSON_FILE, LEGACY_WORKBOOK, ROOT, WORKBOOK_FILE
from app.data_loader import build_payload, load_records, load_workbook_records, resolve_workbook, slim
from app.export_data import export_json, read_status
from app.export_excel import build_detail_workbook, detail_filename
from app.etl import run_etl
from app.publish import list_history, rollback

STATIC = ROOT / "static"

app = Flask(__name__, static_folder=str(STATIC), static_url_path="/static")
app.config["MAX_CONTENT_LENGTH"] = 80 * 1024 * 1024
app.config["SEND_FILE_MAX_AGE_DEFAULT"] = 0


@app.after_request
def disable_cache(response):
    path = request.path or ""
    if path in ("/", "/admin") or path.startswith("/static/"):
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
    return response


def bootstrap_data() -> None:
    """Publish JSON from legacy workbook on first run."""
    if JSON_FILE.exists():
        return
    try:
        workbook = resolve_workbook()
    except FileNotFoundError:
        return
    export_json(workbook, version="bootstrap")
    if not WORKBOOK_FILE.exists() and workbook != WORKBOOK_FILE:
        WORKBOOK_FILE.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(workbook, WORKBOOK_FILE)


@app.get("/")
def index():
    return send_from_directory(STATIC, "index.html")


@app.get("/index.html")
def index_html():
    return send_from_directory(STATIC, "index.html")


@app.get("/admin")
def admin_page():
    return send_from_directory(STATIC, "admin.html")


@app.get("/api/data")
def api_data():
    try:
        if JSON_FILE.exists():
            payload = json.loads(JSON_FILE.read_text(encoding="utf-8"))
        else:
            workbook = resolve_workbook()
            records = [slim(r) for r in load_workbook_records(workbook)]
            payload = build_payload(records, workbook, "live")
        resp = jsonify(payload)
        resp.headers["Cache-Control"] = "no-store"
        return resp
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.get("/api/status")
def api_status():
    status = read_status()
    if not status.get("rowCount") and JSON_FILE.exists():
        payload = json.loads(JSON_FILE.read_text(encoding="utf-8"))
        status["rowCount"] = payload.get("rowCount", len(payload.get("records", [])))
        status["modified"] = payload.get("modified")
        status["version"] = payload.get("version")
    resp = jsonify(status)
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.post("/api/export/detail")
def export_detail():
    """Export filtered detailed-report rows to Excel."""
    try:
        payload = request.get_json(silent=True) or {}
        records = payload.get("records") or []
        columns = payload.get("columns") or []
        if not records:
            return jsonify({"error": "No rows to export. Apply filters or clear the search box."}), 400
        if not columns and records:
            columns = list(records[0].keys())
        data = build_detail_workbook(records, columns)
        return send_file(
            io.BytesIO(data),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=detail_filename(),
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.post("/api/refresh")
def api_refresh():
    """Re-run ETL against data/input and republish."""
    try:
        from app.config import INPUT, STAGING
        from app.etl import etl_summary
        from app.publish import publish_workbook

        output_dir = STAGING / "manual_refresh" / "output"
        code = run_etl(INPUT, output_dir)
        if code != 0:
            return jsonify({"error": "ETL failed. Check data/input for source files."}), 400
        summary = etl_summary(output_dir)
        wb_path = Path(summary["workbook"])
        manifest = publish_workbook(wb_path, upload_id="manual_refresh", warnings=summary.get("warnings"))
        return jsonify(manifest)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.post("/admin/upload")
def admin_upload():
    files = request.files.getlist("files")
    try:
        manifest = process_upload(files)
        return jsonify({"ok": True, **manifest})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


@app.get("/admin/history")
def admin_history():
    return jsonify({"history": list_history()})


@app.post("/admin/rollback/<version>")
def admin_rollback(version: str):
    try:
        manifest = rollback(version)
        return jsonify({"ok": True, **manifest})
    except Exception as exc:
        return jsonify({"ok": False, "error": str(exc)}), 400


def _lan_urls(port: int) -> list[str]:
    import socket

    urls = [f"http://127.0.0.1:{port}"]
    try:
        hostname = socket.gethostname()
        for info in socket.getaddrinfo(hostname, None, socket.AF_INET):
            ip = info[4][0]
            if ip.startswith("127.") or ip.startswith("169.254."):
                continue
            url = f"http://{ip}:{port}"
            if url not in urls:
                urls.append(url)
    except OSError:
        pass
    return urls


def main() -> None:
    bootstrap_data()
    from app.config import HOST, PORT

    print(f"\nLogistics Control Tower")
    print(f"Listening on {HOST}:{PORT}")
    for url in _lan_urls(PORT):
        print(f"Open: {url}")
    print(f"Admin: http://127.0.0.1:{PORT}/admin")
    if JSON_FILE.exists():
        status = read_status()
        print(f"Rows: {status.get('rowCount', '?')} · {status.get('modified', 'unknown')}\n")
    elif LEGACY_WORKBOOK.exists():
        print(f"Bootstrap source: {LEGACY_WORKBOOK}\n")
    else:
        print("No data yet — upload files at /admin\n")
    # Reloader keeps a leftover 127.0.0.1 socket; LAN demos need a single bind.
    app.run(host=HOST, port=PORT, debug=False, threaded=True, use_reloader=False)


if __name__ == "__main__":
    main()
