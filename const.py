from __future__ import annotations


INDEX_HTML = """<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>报价查询系统</title>
  <style>
    :root { color-scheme: light; --ink:#172033; --muted:#667085; --line:#d8dee8; --bg:#f6f8fb; --panel:#ffffff; --blue:#2563eb; --green:#0f9f6e; --amber:#b7791f; --red:#c2410c; }
    * { box-sizing: border-box; }
    body { margin:0; font-family:-apple-system,BlinkMacSystemFont,"Segoe UI","PingFang SC","Microsoft YaHei",sans-serif; color:var(--ink); background:var(--bg); }
    main { padding:16px; }
    .toolbar { display:grid; grid-template-columns:minmax(360px, 1fr) auto; gap:16px; align-items:end; background:var(--panel); border:1px solid var(--line); padding:16px; border-radius:8px; }
    .toolbar-upload { flex:1; min-width:300px; }
    .toolbar-actions { display:flex; gap:8px; align-items:center; align-self:end; }
    label { display:block; font-size:13px; color:#344054; margin-bottom:6px; font-weight:600; }
    .file-input { position:absolute; width:1px; height:1px; opacity:0; pointer-events:none; }
    .file-picker { height:42px; border:1px solid var(--line); border-radius:6px; background:#fff; display:flex; align-items:center; overflow:hidden; max-width:720px; }
    .file-picker-btn { height:100%; border:0; border-right:1px solid var(--line); background:#f3f6fb; color:var(--ink); padding:0 16px; font-weight:700; cursor:pointer; display:inline-flex; align-items:center; justify-content:center; line-height:1; }
    .file-name { flex:1; min-width:0; height:100%; padding:0 12px; color:var(--muted); white-space:nowrap; overflow:hidden; text-overflow:ellipsis; font-size:14px; display:flex; align-items:center; }
    .btn { border:0; padding:0 18px; border-radius:6px; font-weight:700; cursor:pointer; font-size:13px; display:inline-flex; align-items:center; justify-content:center; gap:6px; height:42px; min-height:42px; white-space:nowrap; }
    .btn-primary { background:var(--blue); color:#fff; }
    .btn-success { background:#0f9f6e; color:#fff; }
    .btn:disabled { background:#98a2b3; cursor:not-allowed; }
    .table-wrap { background:var(--panel); border:1px solid var(--line); border-radius:8px; overflow:auto; max-height:calc(100vh - 148px); }
    table { border-collapse:collapse; width:100%; min-width:900px; }
    th, td { padding:9px 10px; border-bottom:1px solid #edf1f6; text-align:left; vertical-align:middle; font-size:13px; }
    th { position:sticky; top:0; background:#f9fafb; z-index:1; color:#344054; white-space:nowrap; font-weight:600; }
    tr:hover td { background:#fbfdff; }
    .pill { display:inline-flex; padding:2px 8px; border-radius:999px; font-weight:700; font-size:11px; white-space:nowrap; }
    .pill-row { display:flex; gap:5px; flex-wrap:wrap; }
    .auto    { background:#dff8eb; color:#067647; }
    .review  { background:#fff4d6; color:#92400e; }
    .missing { background:#ffe8df; color:#9a3412; }
    .muted   { color:var(--muted); font-size:12px; }
    .price   { font-variant-numeric:tabular-nums; text-align:right; }
    .hl      { color:#0f9f6e; font-weight:700; }
    .match-btn { width:100%; min-width:260px; border:1px solid var(--line); border-radius:6px; padding:7px 9px; background:#fff; color:var(--ink); text-align:left; cursor:pointer; font-size:12px; }
    .match-btn:hover { border-color:var(--blue); }
    .modal-backdrop { position:fixed; inset:0; background:rgba(15,23,42,.35); display:none; align-items:center; justify-content:center; padding:24px; z-index:10; }
    .modal { width:min(1180px, 100%); max-height:88vh; background:#fff; border-radius:8px; box-shadow:0 20px 55px rgba(15,23,42,.25); display:flex; flex-direction:column; overflow:hidden; }
    .modal-head { padding:14px 16px; border-bottom:1px solid var(--line); display:flex; align-items:center; justify-content:space-between; gap:12px; }
    .modal-title { font-weight:700; }
    .modal-title-sub { display:flex; gap:8px; flex-wrap:wrap; margin-top:8px; font-size:12px; color:var(--muted); font-weight:600; }
    .context-pill { border:1px solid var(--line); border-radius:6px; padding:5px 8px; background:#fff; }
    .modal-body { padding:14px 16px; overflow:auto; }
    .modal-foot { padding:12px 16px; border-top:1px solid var(--line); display:flex; justify-content:flex-end; gap:8px; }
    .search-row { display:flex; gap:8px; align-items:center; margin-bottom:12px; }
    .search-input { flex:1; min-width:0; border:1px solid var(--line); border-radius:6px; padding:10px; font-size:14px; }
    .search-mode { display:flex; gap:6px; flex:0 0 auto; }
    .mode-btn { border:0; border-radius:6px; background:#f3f6fb; padding:0 12px; min-height:40px; font-size:13px; font-weight:700; color:var(--muted); cursor:pointer; }
    .mode-btn.active { background:#eff6ff; color:#1d4ed8; }
    .tab-row { display:flex; gap:8px; flex-wrap:wrap; margin:0 0 12px; }
    .tab { border:1px solid var(--line); background:#fff; color:var(--ink); border-radius:999px; padding:7px 14px; font-size:12px; font-weight:700; cursor:pointer; }
    .tab.active { border-color:var(--blue); background:#eff6ff; color:#1d4ed8; }
    .pick-table-wrap { border:1px solid var(--line); border-radius:8px; overflow:auto; max-height:52vh; }
    .pick-table { width:100%; border-collapse:collapse; min-width:900px; }
    .pick-table th, .pick-table td { padding:9px 10px; border-bottom:1px solid #edf1f6; font-size:13px; vertical-align:top; text-align:left; }
    .pick-table th { position:sticky; top:0; background:#f9fafb; z-index:1; white-space:nowrap; }
    .pick-table tr { cursor:pointer; }
    .pick-table tr:hover td, .pick-table tr.selected td { background:#eff6ff; }
    .pick-price { width:92px; text-align:right; font-variant-numeric:tabular-nums; }
    .pick-code { width:120px; white-space:nowrap; }
    .pick-seq { width:64px; min-width:64px; white-space:nowrap; }
    .pick-source { width:72px; min-width:72px; white-space:nowrap; }
    .pick-remark { color:var(--muted); font-size:12px; }
    .btn-plain { background:#fff; color:var(--ink); border:1px solid var(--line); }
    .notice { margin:10px 0 0; color:var(--muted); font-size:12px; }
    .summary { display:none; align-items:center; gap:8px; flex-wrap:wrap; margin:12px 0; }
    .summary-item { background:#fff; border:1px solid var(--line); border-radius:6px; padding:7px 10px; font-size:12px; color:#344054; cursor:pointer; }
    .summary-item.active { border-color:var(--blue); background:#eff6ff; color:#1d4ed8; }
    .summary-item strong { font-size:14px; color:var(--ink); margin-left:4px; }
    .error  { background:#fff1f0; color:#9a3412; border:1px solid #ffccc7; padding:12px; border-radius:8px; margin:12px 0; white-space:pre-wrap; }
    @media (max-width:700px) { main { padding:10px; } .toolbar { grid-template-columns:1fr; } }
  </style>
</head>
<body>
  <main>
    <form class="toolbar" id="uploadForm">
      <div class="toolbar-upload">
        <label>外部报价清单（xlsx）</label>
        <label class="file-picker">
          <input class="file-input" id="quoteFile" type="file" name="quote" accept=".xlsx" required />
          <span class="file-picker-btn">选择文件</span>
          <span class="file-name" id="fileName">未选择文件</span>
        </label>
      </div>
      <div class="toolbar-actions">
        <button class="btn btn-primary" id="matchBtn" type="submit">开始匹配</button>
        <button class="btn btn-plain" type="button" onclick="window.location.href='/download_template'">下载导入模板</button>
        <button class="btn btn-success" id="downloadBtn" type="button" onclick="downloadMatches()" disabled>下载结果</button>
      </div>
    </form>
    <div id="message"></div>
    <div class="summary" id="summary"></div>
    <section id="results" style="display:none">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>序号</th>
              <th>检测材料</th>
              <th>检测参数</th>
              <th>组/点数</th>
              <th class="price">单价</th>
              <th class="price">合价</th>
              <th>备注</th>
              <th>状态</th>
              <th>匹配项</th>
            </tr>
          </thead>
          <tbody id="tbody"></tbody>
        </table>
      </div>
    </section>
  </main>
  <div class="modal-backdrop" id="matchModal">
    <div class="modal">
      <div class="modal-head">
        <div>
          <div class="modal-title">选择匹配项</div>
          <div class="modal-title-sub" id="matchContext"></div>
        </div>
        <button class="btn btn-plain" type="button" onclick="closeMatchModal()">关闭</button>
      </div>
      <div class="modal-body">
        <div class="search-row">
          <input class="search-input" id="matchSearch" type="search" placeholder="输入检测项目、检测材料或别名后点击搜索" />
          <div class="search-mode">
            <button class="mode-btn active" id="modeFuzzy" type="button" onclick="setSearchMode('fuzzy')">模糊</button>
            <button class="mode-btn" id="modeExact" type="button" onclick="setSearchMode('exact')">精确</button>
          </div>
          <button class="btn btn-primary" id="matchSearchBtn" type="button" onclick="runMatchSearch()">搜索</button>
        </div>
        <div class="tab-row" id="matchTabs"></div>
        <div class="pick-table-wrap">
          <table class="pick-table">
            <thead>
              <tr>
                <th>序号</th>
                <th>检测项目</th>
                <th>检测材料</th>
                <th>检测参数</th>
                <th>单位</th>
                <th class="pick-price">单价（元）</th>
                <th>备注</th>
                <th>报价编号</th>
                <th>检测项目别名</th>
                <th>检测参数别名</th>
              </tr>
            </thead>
            <tbody id="matchResults"></tbody>
          </table>
        </div>
      </div>
      <div class="modal-foot">
        <button class="btn btn-plain" type="button" onclick="closeMatchModal()">取消</button>
        <button class="btn btn-primary" type="button" onclick="confirmMatchSelection()">确定</button>
      </div>
    </div>
  </div>
<script>
let state = { sessionId: null, matches: [], activeMatchId: null, searchResults: [], selectedItemId: null, tabs: ['推荐'], activeTab: '推荐', searchQuery: '', searchMode: 'fuzzy', statusFilter: 'all' };
const $ = (id) => document.getElementById(id);

function statusClass(s) {
  if (s === '自动匹配') return 'auto';
  if (s === '待确认')   return 'review';
  return 'missing';
}

function money(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? v : n.toFixed(2).replace(/\\.00$/, '');
}

function suggestedPrice(m) {
  const item = m.matched || {};
  if (m.matched_price !== null && m.matched_price !== undefined) return money(m.matched_price);
  if (item.raw_price) return item.raw_price;
  return '';
}

function displayRemark(m) {
  const item = m.matched || {};
  return item.remark || m.matched_remark || '';
}

function currentMatchLabel(m) {
  return m.matched_label || m.matched_code || '请选择';
}

function matchButton(m) {
  return `<button class="match-btn" type="button" onclick="openMatchModal('${m.id}')">${currentMatchLabel(m)}</button>`;
}

function statusPills(m) {
  const pills = [`<span class="pill ${statusClass(m.match_status)}">${m.match_status}</span>`];
  if (m.match_method === '人工确认' || m.manual_confirmed) {
    pills.push('<span class="pill review">手动确认</span>');
  }
  if (m.alias_learned) {
    pills.push('<span class="pill auto">已学习别名</span>');
  }
  return `<div class="pill-row">${pills.join('')}</div>`;
}

function updateStat() {
  const total   = state.matches.length;
  const auto    = state.matches.filter(m => m.match_status === '自动匹配').length;
  const review  = state.matches.filter(m => m.match_status === '待确认').length;
  const missing = total - auto - review;
  const manual  = state.matches.filter(m => m.match_method === '人工确认' || m.manual_confirmed).length;
  return { total, auto, review, missing, manual };
}

function renderSummary() {
  const stat = updateStat();
  if (!stat.total) {
    $('summary').style.display = 'none';
    return;
  }
  $('summary').style.display = 'flex';
  $('summary').innerHTML = `
    <button class="summary-item ${state.statusFilter === 'all' ? 'active' : ''}" type="button" onclick="setStatusFilter('all')">总计<strong>${stat.total}</strong></button>
    <button class="summary-item ${state.statusFilter === 'auto' ? 'active' : ''}" type="button" onclick="setStatusFilter('auto')">自动匹配<strong>${stat.auto}</strong></button>
    <button class="summary-item ${state.statusFilter === 'review' ? 'active' : ''}" type="button" onclick="setStatusFilter('review')">待确认<strong>${stat.review}</strong></button>
    <button class="summary-item ${state.statusFilter === 'missing' ? 'active' : ''}" type="button" onclick="setStatusFilter('missing')">未匹配<strong>${stat.missing}</strong></button>
    <button class="summary-item ${state.statusFilter === 'manual' ? 'active' : ''}" type="button" onclick="setStatusFilter('manual')">手动确认<strong>${stat.manual}</strong></button>
  `;
}

function visibleMatches() {
  if (state.statusFilter === 'auto') return state.matches.filter(m => m.match_status === '自动匹配');
  if (state.statusFilter === 'review') return state.matches.filter(m => m.match_status === '待确认');
  if (state.statusFilter === 'missing') return state.matches.filter(m => m.match_status === '未匹配');
  if (state.statusFilter === 'manual') return state.matches.filter(m => m.match_method === '人工确认' || m.manual_confirmed);
  return state.matches;
}

function renderRows() {
  $('tbody').innerHTML = visibleMatches().map((m, idx) => {
    return `<tr data-id="${m.id}">
      <td class="muted">${m.seq || idx + 1}</td>
      <td>${m.sample_name || ''}</td>
      <td>${m.parameter || m.project_name || ''}</td>
      <td class="price">${money(m.quantity)}</td>
      <td class="price hl">${suggestedPrice(m)}</td>
      <td class="price">${money(m.calculated_total)}</td>
      <td>${displayRemark(m)}</td>
      <td>${statusPills(m)}</td>
      <td>${matchButton(m)}</td>
    </tr>`;
  }).join('');
}

window.setStatusFilter = function(filter) {
  state.statusFilter = filter;
  renderSummary();
  renderRows();
};

function applyItemToMatch(match, item) {
  if (!match || !item) return;
  match.matched = item;
  match.match_status = '自动匹配';
  match.match_method = '人工确认';
  match.manual_confirmed = true;
  match.match_score = item.score || match.match_score || 100;
  match.matched_price = item.price;
  match.matched_price_text = item.raw_price;
  match.matched_code = item.code || item.parameter || item.id;
  match.matched_label = [item.category, item.material || item.parameter].filter(Boolean).join(' / ');
  match.matched_remark = item.remark || '';
  match.price_rule_result = item.price_rule
    ? { requires_selection: true, options: item.price_rule.tiers || [], multipliers: item.price_rule.multipliers || [], message: '复合报价，需选择规格' }
    : null;
  match.price_explanation = match.price_rule_result ? match.price_rule_result.message : '';
  match.calculated_total = (match.price_rule_result || item.price == null)
    ? null : Number(((match.quantity || 1) * item.price).toFixed(2));
}

window.openMatchModal = async function(matchId) {
  state.activeMatchId = matchId;
  $('matchModal').style.display = 'flex';
  const match = state.matches.find((m) => m.id === matchId);
  state.selectedItemId = match && match.matched ? match.matched.id : null;
  $('matchContext').innerHTML = match ? `
    <span class="context-pill">检测材料：${match.sample_name || ''}</span>
    <span class="context-pill">检测参数：${match.parameter || match.project_name || ''}</span>
  ` : '';
  state.activeTab = '推荐';
  state.searchQuery = '';
  $('matchSearch').value = '';
  await searchMatchItems();
  setTimeout(() => $('matchSearch').focus(), 0);
};

window.closeMatchModal = function() {
  $('matchModal').style.display = 'none';
  state.activeMatchId = null;
  state.selectedItemId = null;
};

async function searchMatchItems() {
  const params = new URLSearchParams();
  params.set('q', state.searchQuery.trim());
  params.set('mode', state.searchMode);
  params.set('tab', state.activeTab);
  if (state.sessionId) params.set('session_id', state.sessionId);
  if (state.activeMatchId) params.set('match_id', state.activeMatchId);
  const resp = await fetch(`/api/search?${params.toString()}`);
  const data = await resp.json();
  if (!resp.ok) {
    $('matchResults').innerHTML = `<tr><td colspan="7"><div class="error">${data.error || '搜索失败'}</div></td></tr>`;
    return;
  }
  state.searchResults = data.items || [];
  state.tabs = data.tabs || ['推荐'];
  if (!state.tabs.includes(state.activeTab)) state.activeTab = '推荐';
  state.selectedItemId = state.searchResults.some((x) => x.id === state.selectedItemId) ? state.selectedItemId : null;
  renderSearchTabs();
  renderSearchResults();
}

function renderSearchTabs() {
  $('matchTabs').innerHTML = state.tabs.map((tab) => {
    const active = state.activeTab === tab ? 'active' : '';
    return `<button class="tab ${active}" type="button" onclick="setSearchTab('${tab}')">${tab}</button>`;
  }).join('');
}

function renderSearchResults() {
  const columns = getPickColumns();
  const head = columns.map((col) => {
    const cls = col.className ? ` class="${col.className}"` : '';
    return `<th${cls}>${col.label}</th>`;
  }).join('');
  document.querySelector('.pick-table thead tr').innerHTML = head;
  const rows = state.searchResults;
  $('matchResults').innerHTML = rows.map((item) => {
    const selected = state.selectedItemId === item.id ? 'selected' : '';
    const cells = columns.map((col) => {
      const cls = col.className ? ` class="${col.className}"` : '';
      return `<td${cls}>${col.value(item)}</td>`;
    }).join('');
    return `<tr class="${selected}" onclick="selectSearchItem('${item.id}')">${cells}</tr>`;
  }).join('') || `<tr><td colspan="${columns.length}"><div class="muted">没有找到匹配项</div></td></tr>`;
}

function getPickColumns() {
  const withSheet = state.activeTab === '推荐';
  const columns = [
    { label: '序号', className: 'pick-seq', value: (item) => item.seq || '' },
  ];
  if (withSheet) columns.push({ label: '来源', className: 'pick-source', value: (item) => item.sheet || '' });
  columns.push({ label: '检测项目', value: (item) => item.category || '' });
  columns.push({ label: '检测材料', value: (item) => item.material || '' });
  columns.push(
    { label: '检测参数', value: (item) => item.parameter || '' },
    { label: '单位', className: 'pick-code', value: (item) => item.unit || '' },
    { label: '单价（元）', className: 'pick-price', value: (item) => item.raw_price || item.price || '' },
    { label: '备注', value: (item) => item.remark || '' },
    { label: '报价编号', className: 'pick-code', value: (item) => item.code || '' },
    { label: '检测项目别名', value: (item) => (item.project_aliases || []).join(' / ') },
    { label: '检测参数别名', value: (item) => (item.parameter_aliases || []).join(' / ') },
  );
  return columns;
}

window.selectSearchItem = function(itemId) {
  state.selectedItemId = itemId;
  renderSearchResults();
};

window.setSearchTab = function(tab) {
  state.activeTab = tab;
  searchMatchItems();
};

window.runMatchSearch = function() {
  state.searchQuery = $('matchSearch').value.trim();
  searchMatchItems();
};

window.setSearchMode = function(mode) {
  state.searchMode = mode === 'exact' ? 'exact' : 'fuzzy';
  $('modeFuzzy').classList.toggle('active', state.searchMode === 'fuzzy');
  $('modeExact').classList.toggle('active', state.searchMode === 'exact');
};

window.confirmMatchSelection = async function() {
  const match = state.matches.find((m) => m.id === state.activeMatchId);
  const item = state.searchResults.find((x) => x.id === state.selectedItemId);
  if (!match || !item) return;
  applyItemToMatch(match, item);
  await learnAliasForMatch(match, item);
  renderRows();
  renderSummary();
  closeMatchModal();
};

async function learnAliasForMatch(match, item) {
  const sample = (match.sample_name || '').trim();
  const project = (match.parameter || match.project_name || '').trim();
  if (!sample && !project) return;
  try {
    const resp = await fetch('/api/learn_alias', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ item, project_alias: sample, parameter_alias: project }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || '别名写入失败');
    match.alias_learned = Boolean(data.updated);
    match.alias_message = data.message || '';
    if (data.updated) {
      item.project_aliases = data.project_aliases || item.project_aliases || [];
      item.parameter_aliases = data.parameter_aliases || item.parameter_aliases || [];
      item.aliases = [...item.project_aliases, ...item.parameter_aliases];
      match.matched.project_aliases = item.project_aliases;
      match.matched.parameter_aliases = item.parameter_aliases;
      match.matched.aliases = item.aliases;
    }
  } catch (err) {
    match.alias_learned = false;
    match.alias_message = err.message;
    $('message').innerHTML = `<div class="error">${err.message}</div>`;
  }
}

$('matchSearch').addEventListener('keydown', (e) => {
  if (e.key === 'Enter') {
    e.preventDefault();
    runMatchSearch();
  }
});

$('quoteFile').addEventListener('change', (e) => {
  const file = e.currentTarget.files && e.currentTarget.files[0];
  $('fileName').textContent = file ? file.name : '未选择文件';
});

window.downloadMatches = async function() {
  if (!state.matches.length) return;
  $('downloadBtn').disabled = true;
  try {
    const resp = await fetch('/api/export', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ session_id: state.sessionId, matches: state.matches }),
    });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || '下载失败');
    window.location.href = data.download_url;
  } catch (err) {
    $('message').innerHTML = `<div class="error">${err.message}</div>`;
  } finally {
    $('downloadBtn').disabled = false;
  }
};

$('uploadForm').addEventListener('submit', async (e) => {
  e.preventDefault();
  $('matchBtn').disabled = true;
  $('downloadBtn').disabled = true;
  $('message').innerHTML = '<div class="notice">正在解析和匹配，请稍等...</div>';
  try {
    const resp = await fetch('/api/match', { method: 'POST', body: new FormData(e.currentTarget) });
    const data = await resp.json();
    if (!resp.ok) throw new Error(data.error || '匹配失败');
    state = { ...state, sessionId: data.session_id, matches: data.matches };
    renderRows();
    renderSummary();
    $('results').style.display = '';
    $('message').innerHTML = '';
    $('downloadBtn').disabled = false;
  } catch (err) {
    $('message').innerHTML = `<div class="error">${err.message}</div>`;
  } finally {
    $('matchBtn').disabled = false;
  }
});
</script>
</body>
</html>
"""
