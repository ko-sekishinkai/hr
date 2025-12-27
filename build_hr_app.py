
# -*- coding: utf-8 -*-
"""
人事情報_統合.xlsx を元に、年度→事業所で絞り込める検索アプリ(静的HTML)を自動生成
以前の build_app.py を参考に、Excel→HTML の一体生成方式
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

# 年度→事業所の選択肢を作成
all_years = sorted(list({r.get("年度", "") for r in records if r.get("年度", "").strip() != ""}))
choices = {"年度": all_years, "事業所By年度": {}}
for y in all_years:
    sites = sorted(list({r.get("事業所", "") for r in records
                         if r.get("年度", "") == y and r.get("事業所", "").strip() != ""}))
    choices["事業所By年度"][y] = sites

# 表示順（存在しない列は自動スキップ）
columns_order = ["年度", "事業所", "辞令", "氏名", "日付", "内容"]

# ---------- CSS ----------
css = """
* { box-sizing: border-box; }
body { font-family: system-ui, -apple-system, 'Segoe UI', Roboto, 'Hiragino Kaku Gothic Pro', 'Noto Sans JP', 'Yu Gothic', Meiryo, sans-serif; margin: 24px; }
h1 { font-size: 1.6rem; margin: 0 0 12px; }
header .meta { color: #666; font-size: .9rem; margin-bottom: 16px; }
.controls { display: flex; gap: 12px; flex-wrap: wrap; margin: 16px 0 12px; }
.controls label { font-weight: 600; font-size: .95rem; }
select { padding: 8px 10px; font-size: .95rem; }
.card { border: 1px solid #ddd; border-radius: 8px; padding: 12px; margin: 12px 0; }
.card h2 { font-size: 1.2rem; margin: 0 0 8px; }
.count { color: #333; font-size: .95rem; margin-bottom: 8px; }
.tablewrap { overflow-x: auto; border: 1px solid #eee; border-radius: 6px; }
table { border-collapse: collapse; width: 100%; min-width: 920px; }
th, td { padding: 8px 10px; border-bottom: 1px solid #eee; text-align: left; }
th { background: #f8f9fb; position: sticky; top: 0; z-index: 1; }
tr:nth-child(even) td { background: #fcfcff; }
.empty { color: #666; padding: 12px; }
.footer { margin-top: 18px; color: #555; font-size: .9rem; }
button { padding: 8px 12px; font-size: .9rem; border: 1px solid #ccc; border-radius: 6px; cursor: pointer; background: #fff; }
button:hover { background: #f4f5f7; }
.note { color: #777; font-size: .85rem; }
.badges { display:flex; gap:6px; flex-wrap:wrap; margin:8px 0; }
.badge { background:#eef3ff; border:1px solid #cbd6ff; color:#2b4dbb; padding:4px 8px; border-radius:999px; font-size:.8rem; }
"""

# ---------- JavaScript ----------
js = """
const DATA = __DATA__;
const CHOICES = __CHOICES__;
const COLS = __COLS__;
const yearSel = document.getElementById('year');
const siteSel = document.getElementById('site');
const exportBtn = document.getElementById('export');

function renderYearChoices() {
  yearSel.innerHTML = '<option value=\"\">年度を選択…</option>' + CHOICES['年度'].map(y => `<option value=\"${y}\">${y}</option>`).join('');
  siteSel.innerHTML = '<option value=\"\">事業所を選択…</option>';
}

function updateSiteChoices() {
  const y = yearSel.value;
  const list = (CHOICES['事業所By年度'][y] || []);
  siteSel.innerHTML = '<option value=\"\">事業所を選択…</option>' + list.map(s => `<option value=\"${s}\">${s}</option>`).join('');
}

function getFiltered() {
  const y = yearSel.value;
  const s = siteSel.value;
  const rows = DATA.filter(r => (!y || r['年度'] === y) && (!s || r['事業所'] === s));
  return rows;
}

function makeTable(containerId, rows) {
  const wrap = document.getElementById(containerId);
  wrap.innerHTML = '';
  if (!rows || rows.length === 0) {
    wrap.innerHTML = '<div class=\"empty\">該当するデータがありません。</div>';
    return;
  }
  const cols = COLS.filter(c => rows.some(r => c in r));
  const thead = '<thead><tr>' + cols.map(c => `<th>${c}</th>`).join('') + '</tr></thead>';
  const tbody = '<tbody>' + rows.map(r => '<tr>' + cols.map(c => `<td>${(r[c] ?? '')}</td>`).join('') + '</tr>').join('') + '</tbody>';
  const html = '<div class=\"tablewrap\"><table>' + thead + tbody + '</table></div>';
  wrap.innerHTML = html;
}

function renderSummary(rows) {
  const cnt = rows.length;
  document.getElementById('count').textContent = `${cnt} 件`;
  const y = yearSel.value || '全年度';
  const s = siteSel.value || '全事業所';
  const badges = [y, s].map(v => `<span class=\"badge\">${v}</span>`).join('');
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
    cols.map(c => String(r[c] ?? '').replaceAll('\"', '\"\"'))
        .map(v => /[\",\\n]/.test(v) ? `\"${v}\"` : v)
        .join(',')
  ));
  const blob = new Blob([csvRows.join('\\n')], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = '人事情報_抽出結果.csv';
  document.body.appendChild(a); a.click();
  document.body.removeChild(a); URL.revokeObjectURL(url);
}

yearSel.addEventListener('change', () => { updateSiteChoices(); renderAll(); });
siteSel.addEventListener('change', () => { renderAll(); });
exportBtn.addEventListener('click', exportCSV);

renderYearChoices();
renderAll();
"""

# ---------- HTML ----------
html_template = """
<!doctype html>
<html lang=\"ja\">
<head>
<meta charset=\"utf-8\" />
<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\" />
<title>人事情報 検索アプリ (年度×事業所)</title>
<style>[[CSS]]</style>
</head>
<body>
  <header>
    <h1>人事情報 検索アプリ <span style=\"font-size:.9rem;color:#555\">(年度×事業所)</span></h1>
    <div class=\"meta\">ソース: [[SRCFILE]] ／ 生成日時: [[TIMESTAMP]]</div>
    <div class=\"controls\">
      <div>
        <label for=\"year\">年度</label><br>
        <select id=\"year\" aria-label=\"年度選択\"></select>
      </div>
      <div>
        <label for=\"site\">事業所</label><br>
        <select id=\"site\" aria-label=\"事業所選択\"></select>
      </div>
      <div style=\"align-self: end;\">
        <button id=\"export\" title=\"現在の抽出結果をCSVで保存\">CSVダウンロード</button>
        <div class=\"note\">※年度を選ぶと事業所の選択肢が絞り込まれます。</div>
      </div>
    </div>
  </header>

  <section class=\"card\">
    <h2>抽出結果</h2>
    <div class=\"badges\" id=\"badges\"></div>
    <div class=\"count\" id=\"count\"></div>
    <div id=\"tbl\"></div>
  </section>

  <div class=\"footer\">このページはExcelから自動生成されています。毎年Excelを更新したら、Pythonで再生成してください。</div>

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
