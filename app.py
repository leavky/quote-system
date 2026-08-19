from __future__ import annotations
import traceback
import uuid
from pathlib import Path
from urllib.parse import quote
from flask import Flask, jsonify, request, send_file
from werkzeug.serving import make_server
from const import INDEX_HTML
from matcher import (
    DEFAULT_PRICE,
    OUTPUT_DIR,
    QUOTE_TEMPLATE,
    UPLOAD_DIR,
    append_price_aliases,
    export_matches,
    load_price_items,
    match_quote,
    quote_line_from_match,
    rank_price_items_for_line,
    search_price_items,
)

SESSIONS: dict[str, dict] = {}


flask_app = Flask(__name__)
flask_app.config["TRUSTED_HOSTS"] = [
    "127.0.0.1",
    "127.0.0.1:8000",
    "localhost",
    "localhost:8000",
]


def json_error(message: str, status: int = 400):
    return jsonify({"error": message}), status


def save_upload(file_storage, session_id: str) -> Path:
    filename = Path(file_storage.filename or "").name
    suffix = Path(filename).suffix.lower()
    if suffix not in {".xlsx", ".xls"}:
        raise ValueError("只支持 Excel 文件")
    UPLOAD_DIR.mkdir(exist_ok=True)
    path = UPLOAD_DIR / f"{session_id}-{filename}"
    file_storage.save(path)
    return path


@flask_app.get("/")
def index():
    return INDEX_HTML


@flask_app.get("/api/search")
def api_search():
    if not DEFAULT_PRICE.exists():
        return json_error("未找到默认报价库：检测报价表.xlsx")
    query = request.args.get("q", "")
    tab = request.args.get("tab", "推荐") or "推荐"
    mode = request.args.get("mode", "fuzzy")
    session_id = request.args.get("session_id", "")
    match_id = request.args.get("match_id", "")
    items = None
    if session_id and match_id and session_id in SESSIONS:
        session = SESSIONS[session_id]
        match = next((m for m in session.get("matches", []) if m.get("id") == match_id), None)
        if match:
            items = rank_price_items_for_line(
                DEFAULT_PRICE,
                quote_line_from_match(match),
                query=query,
                tab=tab,
                mode=mode,
            )
    if items is None:
        items = search_price_items(DEFAULT_PRICE, query, limit=None, mode=mode)
        if tab and tab != "推荐":
            items = [item for item in items if item.get("sheet") == tab]
            items.sort(key=lambda item: (item.get("row_number") or 0, item.get("seq") or ""))
    sheets = []
    for item in load_price_items(DEFAULT_PRICE):
        if item.sheet not in sheets:
            sheets.append(item.sheet)
    return jsonify({"items": items, "tabs": ["推荐", *sheets]})


@flask_app.post("/api/match")
def api_match():
    try:
        session_id = uuid.uuid4().hex[:12]
        if not DEFAULT_PRICE.exists():
            return json_error("未找到默认报价库：检测报价表.xlsx")
        quote_file = request.files.get("quote")
        if quote_file is None or not quote_file.filename:
            return json_error("请上传外部报价清单")
        quote_path = save_upload(quote_file, session_id)
        result = match_quote(DEFAULT_PRICE, quote_path)
        SESSIONS[session_id] = result
        return jsonify({"session_id": session_id, **result})
    except Exception as exc:
        traceback.print_exc()
        return json_error(str(exc), 500)


@flask_app.post("/api/export")
def api_export():
    try:
        payload = request.get_json(silent=True) or {}
        session_id = payload.get("session_id") or uuid.uuid4().hex[:12]
        matches = payload.get("matches")
        if not isinstance(matches, list):
            return json_error("没有可导出的匹配结果")
        OUTPUT_DIR.mkdir(exist_ok=True)
        filename = f"报价匹配结果-{session_id}.xlsx"
        export_matches(matches, OUTPUT_DIR / filename)
        return jsonify({"download_url": f"/download?file={quote(filename)}"})
    except Exception as exc:
        traceback.print_exc()
        return json_error(str(exc), 500)


@flask_app.post("/api/learn_alias")
def api_learn_alias():
    try:
        payload = request.get_json(silent=True) or {}
        item = payload.get("item")
        if not isinstance(item, dict):
            return json_error("缺少匹配条目信息")
        result = append_price_aliases(
            DEFAULT_PRICE,
            item,
            project_alias=payload.get("project_alias", ""),
            parameter_alias=payload.get("parameter_alias", ""),
        )
        return jsonify(result)
    except Exception as exc:
        traceback.print_exc()
        return json_error(str(exc), 500)


@flask_app.get("/download")
def download():
    name = request.args.get("file", "")
    path = (OUTPUT_DIR / name).resolve()
    output_root = OUTPUT_DIR.resolve()
    if not str(path).startswith(str(output_root)) or not path.exists():
        return json_error("文件不存在", 404)
    return send_file(
        path,
        as_attachment=True,
        download_name=path.name,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@flask_app.get("/download_template")
def download_template():
    if not QUOTE_TEMPLATE.exists():
        return json_error("未找到导入模板：报价清单导入模板.xlsx", 404)
    return send_file(
        QUOTE_TEMPLATE,
        as_attachment=True,
        download_name="报价清单导入模板.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def create_server(host: str = "127.0.0.1", port: int = 8000):
    UPLOAD_DIR.mkdir(exist_ok=True)
    OUTPUT_DIR.mkdir(exist_ok=True)
    return make_server(host, port, flask_app, threaded=True)


def main() -> None:
    server = create_server()
    print("报价查询系统已启动：http://127.0.0.1:8000")
    server.serve_forever()


if __name__ == "__main__":
    main()
