
# -*- coding: utf-8 -*-
"""
人事情報_統合.xlsx を元に、年度・事業所で絞り込める検索アプリ(静的HTML)を自動生成
UI: 参照実装に合わせて、ドロップダウン(ボタン+パネル)内に縦並びチェックボックス、
    パネル上部に「すべて選択」「すべて解除」を配置
"""
import pandas as pd
import json
from datetime import datetime
import os

EXCEL_FILE = "人事情報_統合.xlsx"
HTML_FILE = "index.html"

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

# 選択肢（独立）：年度 と 事業所
all_years = sorted(list({r.get("年度", "") for r in records if r.get("年度", "").strip() != ""}))
all_sites = sorted(list({r.get("事業所", "") for r in records if r.get("事業所", "").strip() != ""}))
choices = {"年度": all_years, "事業所": all_sites}

# 表示順（存在しない列は自動スキップ）
columns_order = ["年度", "事業所", "辞令", "氏名", "日付", "内容"]

# ---------------- CSS ----------------
css = """
* { box-sizing: border-box; }
body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', 'Yu Gothic', Meiryo, sans-serif; margin: 24px; }
h1 { font-size: 1.6rem; margin: 0 0 12px; }
header .meta { color: #666; font-size: .9rem; margin-bottom: 16px; }
.controls { display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0 12px; align-items: flex-start; }
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
  white-space: nowrap;
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

/* ▼ 参照実装に合わせて統一 */
.dropdown { position: relative; display: inline-block; }
.dropdown-toggle {
  padding: 8px 12px; font-size: .95rem; border: 1px solid #ccc;
  border-radius: 6px; background: #fff; cursor: pointer;
}
.dropdown-toggle[aria-expanded="true"] { background: #f4f5f7; }
.dropdown-panel {
  position: absolute; z-index: 1000; min-width: 300px; margin-top: 6px;
  background: #fff; border: 1px solid #ddd; border-radius: 8px;
  box-shadow: 0 8px 24px rgba(0,0,0,.12);
  padding: 10px; display: none;
}
.dropdown-panel.open { display: block; }
.dropdown-actions {
  display: flex; gap: 8px; justify-content: flex-end; margin-bottom: 8px;
}
.dropdown-actions button {
  padding: 6px 10px; font-size: .85rem; border: 1px solid #ccc;
  border-radius: 6px; background: #fff; cursor: pointer;
}
.dropdown-actions button:hover { background: #f4f5f7; }
.checkbox-list {
  display: grid; grid-template-columns: 1fr; gap: 6px;
  max-height: 280px; overflow-y: auto; border: 1px solid #eee;
  padding: 8px; border-radius: 6px; background: #fff;
}
.chk { display: flex; align-items: center; gap: 8px; font-size: .95rem; }
/* ▲ 統一ここまで */
"""

# ---------------- JavaScript ----------------
js = r"""
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;

const exportBtn = document.getElementById('export');

// 年度
const ddYearBtn   = document.getElementById('dd-year-btn');
const ddYearPanel = document.getElementById('dd-year-panel');
const yearList    = document.getElementById('year_list');
const yearSelectAllBtn = document.getElementById('year_select_all');
const yearClearAllBtn  = document.getElementById('year_clear_all');

// 事業所
const ddSiteBtn   = document.getElementById('dd-site-btn');
const ddSitePanel = document.getElementById('dd-site-panel');
const siteList    = document.getElementById('site_list');
const siteSelectAllBtn = document.getElementById('site_select_all');
const siteClearAllBtn  = document.getElementById('site_clear_all');

// 候補描画（縦並び）
function renderYearChoices() {
  yearList.innerHTML = CHOICES['年度']
    .map(y => `<label class="chk"><input type="checkbox" name="year" value="${y}">${y}</label>`)
    .join('');
}
function renderSiteChoices() {
  siteList.innerHTML = CHOICES['事業所']
    .map(s => `<label class="chk"><input type="checkbox" name="site" value="${s}">${s}</label>`)
    .join('');
}

// 選択値取得（name基準）
function getChecked(name) {
  return Array.from(document.querySelectorAll(`input[name="${name}"]:checked`)).map(el => el.value);
}

// フィルタ
function getFiltered() {
  const years = getChecked('year');
  const sites = getChecked('site');
  return DATA.filter(r =>
    (years.length === 0 || years.includes(r['年度'])) &&
    (sites.length === 0 || sites.includes(r['事業所']))
  );
}

// テーブル作成
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
    cols.map(c => `<td>${(r[c] ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  const html = '<div class="tablewrap"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}

// バッジ
function renderBadges() {
  const years = getChecked('year');
  const sites = getChecked('site');
  document.getElementById('badge_year').textContent = years.length ? years.join(', ') : '未選択';
  document.getElementById('badge_site').textContent = sites.length ? sites.join(', ') : '未選択';
}

// 再描画
function renderAll() {
  const rows = getFiltered();
  document.getElementById('count').textContent = `${rows.length} 件`;
  makeTable('tbl', rows);
  renderBadges();
}

// CSV
function exportCSV() {
  const rows = getFiltered();
  if (rows.length === 0) { alert('出力対象がありません。'); return; }
  const cols = Array.from(new Set(rows.flatMap(r => Object.keys(r))));
  const header = cols.join(',');
  const csvRows = [header].concat(rows.map(r =>
    cols.map(c => String(r[c] ?? '').replaceAll('"', '""'))
        .map(v => /[",\n]/.test(v) ? `"${v}"` : v)
        .join(',')
  ));
  const blob = new Blob([csvRows.join('\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '人事情報_抽出結果.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

// 開閉
function toggleDropdown(btn, panel, open) {
  const isOpen = (open != null) ? open : !(panel.classList.contains('open'));
  panel.classList.toggle('open', isOpen);
  btn.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
}
ddYearBtn.addEventListener('click', () => toggleDropdown(ddYearBtn, ddYearPanel));
ddSiteBtn.addEventListener('click', () => toggleDropdown(ddSiteBtn, ddSitePanel));

// 外側クリックで閉じる
document.addEventListener('click', (e) => {
  if (!ddYearBtn.contains(e.target) && !ddYearPanel.contains(e.target)) toggleDropdown(ddYearBtn, ddYearPanel, false);
  if (!ddSiteBtn.contains(e.target) && !ddSitePanel.contains(e.target)) toggleDropdown(ddSiteBtn, ddSitePanel, false);
});
// ESCで閉じる
document.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    toggleDropdown(ddYearBtn, ddYearPanel, false);
    toggleDropdown(ddSiteBtn, ddSitePanel, false);
  }
});

// 一括選択/解除
yearSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = true);
  renderAll();
});
yearClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="year"]').forEach(el => el.checked = false);
  renderAll();
});
siteSelectAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="site"]').forEach(el => el.checked = true);
  renderAll();
});
siteClearAllBtn.addEventListener('click', () => {
  document.querySelectorAll('input[name="site"]').forEach(el => el.checked = false);
  renderAll();
});

// 変更即時反映
document.addEventListener('change', (e) => {
  if (e.target && (e.target.name === 'year' || e.target.name === 'site')) renderAll();
});

// 初期化
(function init() {
  renderYearChoices();
  renderSiteChoices();
  renderAll();
  exportBtn.addEventListener('click', exportCSV);
})();
"""

# ---------------- HTML ----------------
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
        <div class="dropdown">
          <button class="dropdown-toggle" id="dd-year-btn" aria-expanded="false" aria-controls="dd-year-panel">年度を選択（複数可）</button>
          <div class="dropdown-panel" id="dd-year-panel" role="listbox" aria-labelledby="dd-year-btn">
            <div class="dropdown-actions">
              <button id="year_select_all" type="button">すべて選択</button>
              <button id="year_clear_all" type="button">すべて解除</button>
            </div>
            <div id="year_list" class="checkbox-list" aria-label="年度選択"></div>
          </div>
        </div>
      </div>
      <!-- 事業所 -->
      <div class="group" aria-label="事業所のドロップダウン">
        <div class="dropdown">
          <button class="dropdown-toggle" id="dd-site-btn" aria-expanded="false" aria-controls="dd-site-panel">事業所を選択（複数可）</button>
          <div class="dropdown-panel" id="dd-site-panel" role="listbox" aria-labelledby="dd-site-btn">
            <div class="dropdown-actions">
              <button id="site_select_all" type="button">すべて選択</button>
              <button id="site_clear_all" type="button">すべて解除</button>
            </div>
            <div id="site_list" class="checkbox-list" aria-label="事業所選択"></div>
          </div>
        </div>
      </div>

      <div class="group" style="min-width:auto;">
        <button id="export" title="現在の抽出結果をCSVで保存">CSVダウンロード</button>
        <div class="note">※年度・事業所は複数選択できます（未選択の場合は全件）</div>
      </div>
    </div>

    <div class="note">選択中 → 年度: <span class="badge" id="badge_year"></span> ／ 事業所: <span class="badge" id="badge_site"></span></div>
  </header>

  <section class="card">
    <h2>人事情報－検索結果</h2>
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
html = html_template.replace("[[CSS]]", css)\
    .replace("[[SRCFILE]]", os.path.basename(EXCEL_FILE))\
    .replace("[[TIMESTAMP]]", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))

js_filled = js.replace("__DATA__", json.dumps(records, ensure_ascii=False))\
    .replace("__CHOICES__", json.dumps(choices, ensure_ascii=False))\
    .replace("__COLS__", json.dumps(columns_order, ensure_ascii=False))

html = html.replace("[[JS]]", js_filled)
with open(HTML_FILE, "w", encoding="utf-8") as f:
    f.write(html)
print("✓ 生成完了:", HTML_FILE)
