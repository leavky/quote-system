from __future__ import annotations

import shutil
import uuid
from pathlib import Path
from typing import Any

from const import INDEX_HTML
from matcher import (
    DEFAULT_PRICE,
    QUOTE_TEMPLATE,
    append_price_aliases,
    export_matches,
    load_price_items,
    match_quote,
    quote_line_from_match,
    rank_price_items_for_line,
    search_price_items,
)

SESSIONS: dict[str, dict[str, Any]] = {}


def make_response(data: dict[str, Any], ok: bool = True, status: int = 200) -> dict[str, Any]:
    return {"ok": ok, "status": status, "data": data}


def json_error(message: str, status: int = 400) -> dict[str, Any]:
    return make_response({"error": message}, ok=False, status=status)


class Api:
    def __init__(self) -> None:
        self.window = None

    def choose_quote_file(self) -> dict[str, Any]:
        import webview

        paths = self.window.create_file_dialog(
            webview.OPEN_DIALOG,
            allow_multiple=False,
            file_types=("Excel 文件 (*.xlsx;*.xls)", "所有文件 (*.*)"),
        )
        if not paths:
            return {"path": "", "name": ""}
        path = Path(paths[0])
        return {"path": str(path), "name": path.name}

    def match(self, quote_path: str) -> dict[str, Any]:
        try:
            if not DEFAULT_PRICE.exists():
                return json_error("未找到默认报价库：检测报价表.xlsx")
            path = Path(quote_path)
            if not path.exists():
                return json_error("请先选择外部报价清单")
            session_id = uuid.uuid4().hex[:12]
            result = match_quote(DEFAULT_PRICE, path)
            SESSIONS[session_id] = result
            return make_response({"session_id": session_id, **result})
        except Exception as exc:
            return json_error(str(exc), 500)

    def search(self, params: dict[str, Any]) -> dict[str, Any]:
        try:
            if not DEFAULT_PRICE.exists():
                return json_error("未找到默认报价库：检测报价表.xlsx")
            query = str(params.get("q") or "")
            tab = str(params.get("tab") or "推荐")
            mode = str(params.get("mode") or "fuzzy")
            session_id = str(params.get("session_id") or "")
            match_id = str(params.get("match_id") or "")
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
            return make_response({"items": items, "tabs": ["推荐", *sheets]})
        except Exception as exc:
            return json_error(str(exc), 500)

    def learn_alias(self, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            item = payload.get("item")
            if not isinstance(item, dict):
                return json_error("缺少匹配条目信息")
            result = append_price_aliases(
                DEFAULT_PRICE,
                item,
                project_alias=payload.get("project_alias", ""),
                parameter_alias=payload.get("parameter_alias", ""),
            )
            return make_response(result)
        except Exception as exc:
            return json_error(str(exc), 500)

    def export(self, payload: dict[str, Any]) -> dict[str, Any]:
        try:
            import webview

            matches = payload.get("matches")
            if not isinstance(matches, list):
                return json_error("没有可导出的匹配结果")
            session_id = payload.get("session_id") or uuid.uuid4().hex[:12]
            default_name = f"报价匹配结果-{session_id}.xlsx"
            paths = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename=default_name,
                file_types=("Excel 文件 (*.xlsx)", "所有文件 (*.*)"),
            )
            if not paths:
                return make_response({"cancelled": True})
            output_path = Path(paths if isinstance(paths, str) else paths[0])
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            export_matches(matches, output_path)
            return make_response({"path": str(output_path), "message": f"已保存：{output_path}"})
        except Exception as exc:
            return json_error(str(exc), 500)

    def download_template(self) -> dict[str, Any]:
        try:
            import webview

            if not QUOTE_TEMPLATE.exists():
                return json_error("未找到导入模板：报价清单导入模板.xlsx", 404)
            paths = self.window.create_file_dialog(
                webview.SAVE_DIALOG,
                save_filename="报价清单导入模板.xlsx",
                file_types=("Excel 文件 (*.xlsx)", "所有文件 (*.*)"),
            )
            if not paths:
                return make_response({"cancelled": True})
            output_path = Path(paths if isinstance(paths, str) else paths[0])
            if output_path.suffix.lower() != ".xlsx":
                output_path = output_path.with_suffix(".xlsx")
            output_path.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(QUOTE_TEMPLATE, output_path)
            return make_response({"path": str(output_path), "message": f"已保存：{output_path}"})
        except Exception as exc:
            return json_error(str(exc), 500)


BRIDGE_SCRIPT = r"""
<script>
window.nativeQuotePath = '';

function nativeResponse(result) {
  return {
    ok: Boolean(result && result.ok),
    status: result && result.status ? result.status : 200,
    json: async () => result && result.data ? result.data : {},
  };
}

window.fetch = async function(url, options = {}) {
  const target = new URL(String(url), 'http://pywebview.local');
  let result;
  if (target.pathname === '/api/search') {
    result = await window.pywebview.api.search(Object.fromEntries(target.searchParams.entries()));
  } else if (target.pathname === '/api/match') {
    result = await window.pywebview.api.match(window.nativeQuotePath || '');
  } else if (target.pathname === '/api/export') {
    result = await window.pywebview.api.export(JSON.parse(options.body || '{}'));
  } else if (target.pathname === '/api/learn_alias') {
    result = await window.pywebview.api.learn_alias(JSON.parse(options.body || '{}'));
  } else {
    result = { ok: false, status: 404, data: { error: '未知请求：' + target.pathname } };
  }
  return nativeResponse(result);
};
</script>
"""


DESKTOP_SCRIPT = r"""
<script>
async function chooseQuoteFile() {
  const file = await window.pywebview.api.choose_quote_file();
  if (!file || !file.path) return;
  window.nativeQuotePath = file.path;
  document.getElementById('fileName').textContent = file.name || file.path;
}

window.downloadTemplate = async function() {
  const result = await window.pywebview.api.download_template();
  const data = result && result.data ? result.data : {};
  if (!result.ok) {
    document.getElementById('message').innerHTML = `<div class="error">${data.error || '保存模板失败'}</div>`;
  } else if (!data.cancelled) {
    document.getElementById('message').innerHTML = `<div class="notice">${data.message || '模板已保存'}</div>`;
  }
};

window.downloadMatches = async function() {
  if (!state.matches.length) return;
  const button = document.getElementById('downloadBtn');
  button.disabled = true;
  try {
    const result = await window.pywebview.api.export({ session_id: state.sessionId, matches: state.matches });
    const data = result && result.data ? result.data : {};
    if (!result.ok) throw new Error(data.error || '保存失败');
    if (!data.cancelled) {
      document.getElementById('message').innerHTML = `<div class="notice">${data.message || '结果已保存'}</div>`;
    }
  } catch (err) {
    document.getElementById('message').innerHTML = `<div class="error">${err.message}</div>`;
  } finally {
    button.disabled = false;
  }
};

document.addEventListener('DOMContentLoaded', () => {
  const picker = document.querySelector('.file-picker');
  const input = document.getElementById('quoteFile');
  picker.addEventListener('click', (event) => {
    event.preventDefault();
    chooseQuoteFile();
  });
  input.removeAttribute('required');
});
</script>
"""


def build_html() -> str:
    html = INDEX_HTML
    html = html.replace(
        "onclick=\"window.location.href='/download_template'\"",
        'onclick="downloadTemplate()"',
    )
    html = html.replace("<script>\nlet state", BRIDGE_SCRIPT + "\n<script>\nlet state", 1)
    html = html.replace("</body>", DESKTOP_SCRIPT + "\n</body>", 1)
    return html


def main() -> None:
    try:
        import webview
    except ImportError:
        raise SystemExit("缺少 pywebview，请先运行：uv sync 或 pip install pywebview")

    api = Api()
    window = webview.create_window(
        "报价查询系统",
        html=build_html(),
        js_api=api,
        width=1280,
        height=820,
        min_size=(960, 620),
        confirm_close=True,
    )
    api.window = window
    webview.start()


if __name__ == "__main__":
    main()
