
# -*- coding: utf-8 -*-
"""
人事情報_統合.xlsx を元に、年度・事業所で絞り込める検索アプリ(静的HTML)を自動生成
UI: 「年度」「事業所」のチェックボックスをそのまま使い、ドロップダウン(プルダウン)パネル内に格納
     さらに、両方ともチェックボックスを縦一列に並べる
"""
import pandas as pd
import json
from datetime import datetime
import os

EXCEL_FILE = "人事情報_統合.xlsx"
HTML_FILE = "人事情報_検索_app.html"

# Excel 読み込み（.xlsx は openpyxl を指定）
xl = pd.ExcelFile(EXCEL_FILE, engine="openpyxl")

# すべてのシートから行を抽出（必要列：年度・事業所）
records = []
for sh in xl.sheet_names:
    df = xl.parse(sh)
    df.columns = [str(c).strip() for c in df.columns]
    if "年度" not in df.columns or "事業所" not in df.columns:
        continue
    sub = df.copy()
    # NaNを空文字へ、文字列化
    for col in sub.columns:
        sub[col] = sub[col].apply(lambda x: "" if pd.isna(x) else str(x))
    # 年度・事業所が空の行を除外
    sub = sub[(sub["年度"].str.strip() != "") & (sub["事業所"].str.strip() != "")]
    records.extend(sub.to_dict(orient="records"))

# 選択肢（独立）：年度 と 事業所 の全候補（チェックボックスと同じ内容）
all_years = sorted(list({r.get("年度", "") for r in records if r.get("年度", "").strip() != ""}))
all_sites = sorted(list({r.get("事業所", "") for r in records if r.get("事業所", "").strip() != ""}))
choices = {"年度": all_years, "事業所": all_sites}

# 表示順（存在しない列は自動スキップ）
columns_order = ["年度", "事業所", "辞令", "氏名", "日付", "内容"]

# -------- CSS --------
css = """
* { box-sizing: border-box; }
body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', 'Yu Gothic', Meiryo, sans-serif; margin: 24px; }
h1 { font-size: 1.6rem; margin: 0 0 12px; }
header .meta { color: #666; font-size: .9rem; margin-bottom: 16px; }
.controls { display: flex; gap: 24px; flex-wrap: wrap; margin: 16px 0 12px; align-items: flex-start; }
.controls .group { min-width: 280px; }
.controls .group > label { display:block; font-weight: 600; font-size: .95rem; margin-bottom: 6px; }

.card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin: 12px 0; }
.card h2 { font-size: 1.2rem; margin: 0 0 8px; }
.count { color: #333; font-size: .95rem; margin-bottom: 8px; }

.tablewrap { overflow-x: auto; border: 1px solid #eee; border-radius: 6px; }
table { border-collapse: collapse; width: 100%; min-width: 920px; }
th, td {
  padding: 8px 10px;
  border-bottom: 1px solid #eee;
  text-align: left;
  white-space: nowrap; /* ← 折り返し禁止（横テキスト固定） */
}
th { background: #f8f9fb; position: sticky; top: 0; z-index: 1; }
tr:nth-child(even) td { background: #fcfcff; }
.empty { color: #666; padding: 12px; }
.footer { margin-top: 18px; color: #555; font-size: .9rem; }
button { padding: 8px 12px; font-size: .9rem; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; background: #fff; }
button:hover { background: #f4f5f7; }
.note { color: #777; font-size: .85rem; }
.badges { display:flex; gap:6px; flex-wrap:wrap; margin:8px 0; }
.badge { background:#eef3ff; border:1px solid #cbd6ff; color:#2b4dbb; padding:4px 8px; border-radius:999px; font-size:.8rem; }

/* ---- Dropdown(プルダウン)パネル ---- */
.dropdown { position: relative; }
.dropdown .toggle {
  display: inline-flex; align-items: center; gap:6px;
  padding: 8px 12px; border:1px solid #ccc; border-radius:6px; background:#fff; cursor:pointer;
}
.dropdown .toggle:hover { background:#f4f5f7; }
.dropdown .panel {
  position: absolute; z-index: 20; margin-top: 8px; min-width: 260px;
  background: #fff; border: 1px solid #ddd; border-radius: 8px;
  box-shadow: 0 8px 22px rgba(0,0,0,.08);
  padding: 10px; max-height: 280px; overflow: auto;
}

/* デフォルトのチェック群（横並び・折り返し） */
.checkgroup { display: flex; gap: 8px 12px; flex-wrap: wrap; align-items: center; }
.checkgroup label { display: inline-flex; align-items: center; gap: 6px; white-space: nowrap; font-size: .95rem; }
.checkgroup input[type="checkbox"] { width: 16px; height: 16px; }
.dropdown .ops { display:flex; gap:8px; margin-top: 8px; }
.dropdown .ops button { font-size:.85rem; padding:6px 10px; }

/* ▼ 年度・事業所とも縦一列に並べる（このセレクタが上書きする） */
#years { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; }
#sites  { display: flex; flex-direction: column; gap: 6px; align-items: flex-start; }
"""

# -------- JavaScript --------
js = """
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;

// コンテナ参照（チェックボックス描画領域）
const yearsBox = document.getElementById('years');
const sitesBox = document.getElementById('sites');
const clearYearsBtn = document.getElementById('clearYears');
const clearSitesBtn = document.getElementById('clearSites');
const exportBtn = document.getElementById('export');

// ドロップダウン操作
const yearsToggle = document.getElementById('yearsToggle');
const yearsPanel  = document.getElementById('yearsPanel');
const sitesToggle = document.getElementById('sitesToggle');
const sitesPanel  = document.getElementById('sitesPanel');

function renderCheckboxes(container, name, values) {
  container.innerHTML = values.map((v, i) => {
    const id = `${name}_${i}`;
    return `<label for="${id}">
      <input type="checkbox" id="${id}" name="${name}" value="${v}">
      ${v}
    </label>`;
  }).join('');
}

function renderChoices() {
  renderCheckboxes(yearsBox, 'year', CHOICES['年度']);
  renderCheckboxes(sitesBox, 'site', CHOICES['事業所']);
}

function getCheckedValues(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`))
              .map(el => el.value);
}

function getFiltered() {
  const ys = getCheckedValues('year'); // 選択した年度（複数）
  const ss = getCheckedValues('site'); // 選択した事業所（複数）
  const rows = DATA.filter(r =>
    (ys.length === 0 || ys.includes(r['年度'])) &&
    (ss.length === 0 || ss.includes(r['事業所']))
  );
  return rows;
}

function makeTable(containerId, rows) {
  const wrap = document.getElementById(containerId);
  wrap.innerHTML = '';
  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class="empty">該当するデータがありません。</div>';
    return;
  }
  const cols = COLS.filter(c => rows.some(r => c in r));
  const thead = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r => '<tr>' +
    cols.map(c => `<td>${(r[c] ?? '')}</td>`).join('') +
    '</tr>').join('') + '</tbody>';
  const html = '<div class="tablewrap"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}

function renderSummary(rows) {
  const cnt = rows.length;
  document.getElementById('count').textContent = `${cnt} 件`;
  const ys = getCheckedValues('year');
  const ss = getCheckedValues('site');
  const yBadge = ys.length ? ys.join(' / ') : '全年度';
  const sBadge = ss.length ? ss.join(' / ') : '全事業所';
  const badges = [yBadge, sBadge].map(v => `<span class="badge">${v}</span>`).join('');
  document.getElementById('badges').innerHTML = badges;
}

function renderAll() {
  const rows = getFiltered();
  renderSummary(rows);
  makeTable('tbl', rows);
}

function exportCSV() {
  const rows = getFiltered();
  if (rows.length === 0) { alert('出力対象がありません。'); return; }
  const cols = Array.from(new Set(rows.flatMap(r => Object.keys(r))));
  const header = cols.join(',');
  const csvRows = [header].concat(rows.map(r =>
    cols.map(c => String(r[c] ?? '').replaceAll('"', '""'))
        .map(v => /[",\\n]/.test(v) ? `"${v}"` : v)
        .join(',')
  ));
  const blob = new Blob([csvRows.join('\\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '人事情報_抽出結果.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// チェック状態変化で再描画（委譲）
document.addEventListener('change', (ev) => {
  const t = ev.target;
  if (t && t.matches('input[type="checkbox"][name="year"], input[type="checkbox"][name="site"]')) {
    renderAll();
  }
});

// クリアボタン
clearYearsBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => { el.checked = false; });
  renderAll();
});
clearSitesBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="site"]').forEach(el => { el.checked = false; });
  renderAll();
});

// ドロップダウン開閉（クリック外しで閉じる）
function setupDropdown(toggleEl, panelEl) {
  toggleEl.addEventListener('click', (e) => {
    e.stopPropagation();
    const open = panelEl.getAttribute('data-open') === 'true';
    document.querySelectorAll('.dropdown .panel').forEach(p => p.setAttribute('data-open','false'));
    panelEl.setAttribute('data-open', open ? 'false' : 'true');
    panelEl.style.display = open ? 'none' : 'block';
  });
}
setupDropdown(yearsToggle, yearsPanel);
setupDropdown(sitesToggle, sitesPanel);
document.addEventListener('click', () => {
  document.querySelectorAll('.dropdown .panel').forEach(p => {
    p.setAttribute('data-open','false'); p.style.display = 'none';
  });
});

// 初期化
renderChoices();
renderAll();
"""

# -------- HTML --------
html_template = """
<!doctype html>
<html lang="ja">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>人事情報</title>
<style>[[CSS]]</style>
</head>
<body>
  <header>
    <h1>人事情報</h1>
    

    <div class="controls">
      <!-- 年度 -->
      <div class="group" aria-label="年度のドロップダウン">
        <label>年度</label>
        <div class="dropdown">
          <button id="yearsToggle" class="toggle" aria-expanded="false" aria-controls="yearsPanel">年度を選択</button>
          <div id="yearsPanel" class="panel" role="menu" aria-label="年度の選択肢" data-open="false" style="display:none;">
            <div id="years" class="checkgroup"></div>
            <div class="ops">
              <button id="clearYears" title="年度の選択を全て解除">年度選択クリア</button>
            </div>
          </div>
        </div>
        <div class="note">※未選択は全件表示。</div>
      </div>

      <!-- 事業所 -->
      <div class="group" aria-label="事業所のドロップダウン">
        <label>事業所</label>
        <div class="dropdown">
          <button id="sitesToggle" class="toggle" aria-expanded="false" aria-controls="sitesPanel">事業所を選択</button>
          <div id="sitesPanel" class="panel" role="menu" aria-label="事業所の選択肢" data-open="false" style="display:none;">
            <div id="sites" class="checkgroup"></div>
            <div class="ops">
              <button id="clearSites" title="事業所の選択を全て解除">事業所選択クリア</button>
            </div>
          </div>
        </div>
        <div class="note">※未選択は全件表示。</div>
      </div>

      <div class="group" style="min-width:auto;">
        <button id="export" title="現在の抽出結果をCSVで保存">CSVダウンロード</button>
      </div>
    </div>
  </header>

  <section class="card">
    <h2>選択結果</h2>
    <div class="badges" id="badges"></div>
    <div class="count" id="count"></div>
    <div id="tbl"></div>
  </section>


<script>
[[JS]]
</script>
</body>
</html>
"""

# テンプレ埋め込みと書き出し
html = html_template.replace("[[CSS]]", css) \
    .replace("[[SRCFILE]]", os.path.basename(EXCEL_FILE)) \
    .replace("[[TIMESTAMP]]", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

js_filled = js.replace("__DATA__", json.dumps(records, ensure_ascii=False)) \
    .replace("__CHOICES__", json.dumps(choices, ensure_ascii=False)) \
    .replace("__COLS__", json.dumps(columns_order, ensure_ascii=False))

html = html.replace("[[JS]]", js_filled)

with open(HTML_FILE, "w", encoding="utf-8") as f:
    f.write(html)

print("✓ 生成完了:", HTML_FILE)
