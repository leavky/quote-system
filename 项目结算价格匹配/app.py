from __future__ import annotations

import os
import traceback
import uuid
from pathlib import Path
from urllib.parse import quote

from flask import Flask, jsonify, request, send_file
from werkzeug.serving import make_server

from settlement import build_match_result, export_result, public_record, reprice_result


BASE_DIR = Path(__file__).resolve().parent
DEFAULT_LEDGER = BASE_DIR / "zc检测费用台帐明细20260725.xls"
DEFAULT_PRICE = BASE_DIR / "检测项目结算计费价格汇总.xlsx"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "outputs"
ALLOWED_SUFFIXES = {".xls", ".xlsx"}
SESSIONS: dict[str, dict] = {}

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


def error(message: str, status: int = 400):
    return jsonify({"error": message}), status


def save_upload(storage, session_id: str, role: str) -> Path | None:
    if storage is None or not storage.filename:
        return None
    suffix = Path(storage.filename).suffix.lower()
    if suffix not in ALLOWED_SUFFIXES:
        raise ValueError("只支持 .xls 或 .xlsx 文件")
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    target = UPLOAD_DIR / f"{session_id}-{role}{suffix}"
    storage.save(target)
    return target


def session_payload(session_id: str, session: dict) -> dict:
    result = session["result"]
    return {
        "session_id": session_id,
        "ledger_name": session["ledger_name"],
        "price_name": session["price_name"],
        "systems": result["systems"],
        "priority": result["priority"],
        "summary": result["summary"],
        "records": [public_record(record) for record in result["records"]],
    }


@app.get("/")
def index():
    return app.send_static_file("index.html")


@app.post("/api/load")
def load_files():
    try:
        session_id = uuid.uuid4().hex[:12]
        ledger_upload = save_upload(request.files.get("ledger"), session_id, "ledger")
        price_upload = save_upload(request.files.get("price"), session_id, "price")
        ledger_path = ledger_upload or DEFAULT_LEDGER
        price_path = price_upload or DEFAULT_PRICE
        if not ledger_path.exists():
            return error(f"未找到默认台账文件：{DEFAULT_LEDGER.name}")
        if not price_path.exists():
            return error(f"未找到默认价格文件：{DEFAULT_PRICE.name}")
        result = build_match_result(ledger_path, price_path)
        session = {
            "ledger_path": ledger_path,
            "price_path": price_path,
            "ledger_name": Path(request.files.get("ledger").filename).name if ledger_upload else ledger_path.name,
            "price_name": Path(request.files.get("price").filename).name if price_upload else price_path.name,
            "result": result,
            "downloads": {},
        }
        SESSIONS[session_id] = session
        return jsonify(session_payload(session_id, session))
    except Exception as exc:
        traceback.print_exc()
        return error(str(exc), 500)


@app.post("/api/reprice")
def reprice():
    try:
        payload = request.get_json(silent=True) or {}
        session_id = str(payload.get("session_id") or "")
        if session_id not in SESSIONS:
            return error("匹配会话已失效，请重新加载文件", 404)
        priority = payload.get("priority")
        if not isinstance(priority, list):
            return error("价格优先级格式不正确")
        session = SESSIONS[session_id]
        reprice_result(session["result"], priority)
        return jsonify(session_payload(session_id, session))
    except Exception as exc:
        traceback.print_exc()
        return error(str(exc), 500)


@app.post("/api/export")
def export():
    try:
        payload = request.get_json(silent=True) or {}
        session_id = str(payload.get("session_id") or "")
        if session_id not in SESSIONS:
            return error("匹配会话已失效，请重新加载文件", 404)
        row_ids = payload.get("row_ids")
        if row_ids is not None and not isinstance(row_ids, list):
            return error("导出范围格式不正确")
        session = SESSIONS[session_id]
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        filename = f"项目结算价格匹配结果-{session_id}.xlsx"
        output_path = OUTPUT_DIR / filename
        export_result(
            session["result"],
            output_path,
            row_ids=row_ids,
            ledger_name=session["ledger_name"],
            price_name=session["price_name"],
        )
        session["downloads"][filename] = output_path
        return jsonify(
            {
                "filename": filename,
                "download_url": f"/download/{session_id}/{quote(filename)}",
                "exported": len(row_ids) if row_ids is not None else len(session["result"]["records"]),
            }
        )
    except Exception as exc:
        traceback.print_exc()
        return error(str(exc), 500)


@app.get("/download/<session_id>/<path:filename>")
def download(session_id: str, filename: str):
    session = SESSIONS.get(session_id)
    path = session and session["downloads"].get(Path(filename).name)
    if not path or not Path(path).exists():
        return error("导出文件不存在", 404)
    return send_file(
        path,
        as_attachment=True,
        download_name=Path(path).name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def create_server(host: str = "127.0.0.1", port: int = 8010):
    return make_server(host, port, app, threaded=True)


def main() -> None:
    host = os.getenv("SETTLEMENT_HOST", "127.0.0.1")
    port = int(os.getenv("SETTLEMENT_PORT", "8010"))
    print(f"项目结算价格匹配系统已启动：http://{host}:{port}", flush=True)
    create_server(host, port).serve_forever()


if __name__ == "__main__":
    main()
