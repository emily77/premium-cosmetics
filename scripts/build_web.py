#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
產生單檔自包含網頁 web/index.html

特性
  - 無外部依賴（CSS/JS/資料全部內嵌），雙擊即可開啟
  - 可拖到 Netlify Drop / Cloudflare Pages / GitHub Pages 免費託管
  - 可上傳 Google Drive 備份
  - 支援深／淺色切換、表格排序、原料搜尋、列印

用法：python scripts/build_web.py
"""

import csv
import json
import os

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
WEB = os.path.join(ROOT, "web")
OUT = os.path.join(WEB, "index.html")


def load_json(name, default=None):
    p = os.path.join(DATA, name)
    if not os.path.exists(p):
        return default
    with open(p, encoding="utf-8") as f:
        return json.load(f)


def load_csv(name):
    p = os.path.join(DATA, name)
    if not os.path.exists(p):
        return []
    with open(p, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main():
    os.makedirs(WEB, exist_ok=True)

    fin = load_json("finance_model.json")
    products = load_json("products.json")["products"]
    cost_rows = load_csv("cost_report.csv")
    detail = load_csv("cost_detail.csv")
    prices = load_json("ingredient_prices.json", {}).get("prices", {})
    master = load_json("ingredient_master.json", {}).get("ingredients", {})

    # 敏感度（重算，與 cost_engine 同公式）
    a = fin["assumptions"]
    sens = []
    for r in cost_rows:
        raw = float(r["raw_material_twd"] or 0)
        if not raw:
            continue
        pack = float(r["packaging_twd"] or 0)
        fill = float(r["filling_twd"] or 0)
        qc = float(r["qc_twd"] or 0)
        row = {"sku": r["sku_code"], "name": r["name_zh"], "base": float(r["total_cost_twd"])}
        for d in (-0.30, -0.15, 0.15, 0.30):
            base = raw * (1 + d) + pack + fill + qc
            row[f"d{int(d*100)}"] = round(base * 1.20 * 1.08, 1)
        sens.append(row)

    # 價格來源統計
    layer_count = {}
    for r in detail:
        layer_count[r["price_layer"]] = layer_count.get(r["price_layer"], 0) + 1

    # 抓取失敗清單
    failed = [{"inci": k, "reason": v.get("reason", "")}
              for k, v in prices.items() if v.get("status") != "ok"]

    blob = {
        "fin": fin,
        "products": products,
        "cost": cost_rows,
        "detail": detail,
        "sens": sens,
        "layerCount": layer_count,
        "failed": failed,
        "masterCount": len(master),
        "priceOk": sum(1 for v in prices.values() if v.get("status") == "ok"),
    }

    html = TEMPLATE.replace("__DATA__", json.dumps(blob, ensure_ascii=False))
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    size = os.path.getsize(OUT)
    print(f"完成 → {OUT}  ({size:,} bytes)")
    print(f"  14 支 SKU、{len(detail)} 筆原料明細、{len(failed)} 筆抓取失敗")


TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Rokaizumi 頂級化妝品線 — 完整資料儀表板</title>
<style>
:root{
  --bg:#0d1117; --bg2:#161b22; --bg3:#1c2128; --line:#30363d;
  --fg:#e6edf3; --fg2:#8b949e; --accent:#58a6ff; --accent2:#3fb950;
  --warn:#d29922; --danger:#f85149; --purple:#bc8cff; --pink:#f778ba;
  --shadow:0 2px 12px rgba(0,0,0,.4);
}
[data-theme="light"]{
  --bg:#f6f8fa; --bg2:#ffffff; --bg3:#eef1f4; --line:#d0d7de;
  --fg:#1f2328; --fg2:#656d76; --accent:#0969da; --accent2:#1a7f37;
  --warn:#9a6700; --danger:#cf222e; --purple:#8250df; --pink:#bf3989;
  --shadow:0 2px 12px rgba(0,0,0,.08);
}
*{box-sizing:border-box;margin:0;padding:0}
body{
  font-family:"Segoe UI","Microsoft JhengHei","PingFang TC",system-ui,sans-serif;
  background:var(--bg); color:var(--fg); line-height:1.65; font-size:15px;
}
a{color:var(--accent);text-decoration:none}
a:hover{text-decoration:underline}

/* ---------- 版面 ---------- */
.wrap{display:flex;min-height:100vh}
nav{
  width:230px;flex:0 0 230px;background:var(--bg2);border-right:1px solid var(--line);
  padding:20px 0;position:sticky;top:0;height:100vh;overflow-y:auto;
}
nav h1{font-size:15px;padding:0 18px 14px;color:var(--accent);letter-spacing:.5px}
nav .sub{font-size:11px;color:var(--fg2);padding:0 18px 16px;line-height:1.5}
nav a{
  display:block;padding:8px 18px;font-size:13.5px;color:var(--fg2);
  border-left:3px solid transparent;
}
nav a:hover{background:var(--bg3);color:var(--fg);text-decoration:none}
nav a.on{border-left-color:var(--accent);color:var(--accent);background:var(--bg3)}
main{flex:1;padding:28px 36px 80px;max-width:1400px;overflow-x:hidden}

/* ---------- 標題 ---------- */
h2{
  font-size:22px;margin:44px 0 6px;padding-bottom:10px;
  border-bottom:2px solid var(--line);scroll-margin-top:20px;
}
h2:first-of-type{margin-top:0}
h2 .n{color:var(--accent);font-weight:400;margin-right:8px}
h3{font-size:16px;margin:26px 0 10px;color:var(--accent)}
h4{font-size:14px;margin:18px 0 8px;color:var(--fg2);font-weight:600}
.lead{color:var(--fg2);font-size:13.5px;margin-bottom:18px}

/* ---------- 卡片 ---------- */
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(180px,1fr));gap:14px;margin:18px 0}
.kpi{
  background:var(--bg2);border:1px solid var(--line);border-radius:10px;
  padding:16px 18px;box-shadow:var(--shadow);
}
.kpi .lbl{font-size:12px;color:var(--fg2);margin-bottom:6px}
.kpi .val{font-size:24px;font-weight:700;letter-spacing:-.5px}
.kpi .unit{font-size:12px;color:var(--fg2);font-weight:400;margin-left:3px}
.kpi .note{font-size:11.5px;color:var(--fg2);margin-top:6px}
.kpi.g .val{color:var(--accent2)}
.kpi.b .val{color:var(--accent)}
.kpi.w .val{color:var(--warn)}
.kpi.r .val{color:var(--danger)}
.kpi.p .val{color:var(--purple)}

/* ---------- 表格 ---------- */
.tw{overflow-x:auto;margin:14px 0;border:1px solid var(--line);border-radius:10px}
table{width:100%;border-collapse:collapse;font-size:13px;background:var(--bg2)}
th{
  background:var(--bg3);padding:10px 12px;text-align:left;font-weight:600;
  color:var(--fg2);font-size:12px;white-space:nowrap;border-bottom:1px solid var(--line);
  position:sticky;top:0;
}
th.s{cursor:pointer;user-select:none}
th.s:hover{color:var(--accent)}
td{padding:9px 12px;border-bottom:1px solid var(--line);vertical-align:top}
tr:last-child td{border-bottom:none}
tbody tr:hover{background:var(--bg3)}
td.n,th.n{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}
tfoot td{background:var(--bg3);font-weight:700;border-top:2px solid var(--line)}
.tag{
  display:inline-block;padding:1px 7px;border-radius:20px;font-size:11px;
  font-weight:600;white-space:nowrap;
}
.t-p0{background:rgba(248,81,73,.15);color:var(--danger)}
.t-p1{background:rgba(210,153,34,.15);color:var(--warn)}
.t-p2{background:rgba(139,148,158,.15);color:var(--fg2)}
.t-ok{background:rgba(63,185,80,.15);color:var(--accent2)}
.t-l1{background:rgba(88,166,255,.15);color:var(--accent)}
.t-l2{background:rgba(188,140,255,.15);color:var(--purple)}
.t-l3{background:rgba(63,185,80,.15);color:var(--accent2)}
.t-bad{background:rgba(248,81,73,.15);color:var(--danger)}

/* ---------- 長條圖 ---------- */
.bar{display:flex;align-items:center;gap:10px;margin:5px 0;font-size:12.5px}
.bar .nm{width:150px;flex:0 0 150px;color:var(--fg2);text-align:right;
  overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.bar .tr{flex:1;background:var(--bg3);border-radius:4px;height:20px;overflow:hidden}
.bar .fl{height:100%;border-radius:4px;transition:width .4s}
.bar .vl{width:110px;flex:0 0 110px;text-align:right;font-variant-numeric:tabular-nums}

/* ---------- SKU 卡 ---------- */
.skus{display:grid;grid-template-columns:repeat(auto-fill,minmax(300px,1fr));gap:14px;margin:18px 0}
.sku{
  background:var(--bg2);border:1px solid var(--line);border-radius:10px;
  padding:16px;box-shadow:var(--shadow);border-top:3px solid var(--accent);
}
.sku.p0{border-top-color:var(--danger)}
.sku.p1{border-top-color:var(--warn)}
.sku.p2{border-top-color:var(--fg2)}
.sku .hd{display:flex;justify-content:space-between;align-items:flex-start;gap:8px}
.sku .code{font-size:11px;color:var(--fg2);font-family:Consolas,monospace}
.sku .nm{font-size:15px;font-weight:700;margin:3px 0 8px}
.sku .meta{font-size:12px;color:var(--fg2);margin-bottom:10px}
.sku .row{display:flex;justify-content:space-between;font-size:12.5px;padding:3px 0;
  border-bottom:1px dashed var(--line)}
.sku .row:last-child{border-bottom:none}
.sku .row b{font-variant-numeric:tabular-nums}

/* ---------- 其他 ---------- */
.note{
  background:var(--bg2);border-left:3px solid var(--accent);border-radius:0 8px 8px 0;
  padding:12px 16px;margin:14px 0;font-size:13px;color:var(--fg2);
}
.note.warn{border-left-color:var(--warn)}
.note.danger{border-left-color:var(--danger)}
.note.ok{border-left-color:var(--accent2)}
.note b{color:var(--fg)}
ul,ol{margin:10px 0 10px 22px;font-size:13.5px}
li{margin:5px 0}
code{
  background:var(--bg3);padding:1px 6px;border-radius:4px;
  font-family:Consolas,monospace;font-size:12.5px;color:var(--pink);
}
.toolbar{display:flex;gap:10px;flex-wrap:wrap;margin:12px 0;align-items:center}
input[type=text],select{
  background:var(--bg2);border:1px solid var(--line);color:var(--fg);
  padding:7px 12px;border-radius:7px;font-size:13px;font-family:inherit;
}
input[type=text]{min-width:240px}
input[type=text]:focus,select:focus{outline:none;border-color:var(--accent)}
button{
  background:var(--bg2);border:1px solid var(--line);color:var(--fg);
  padding:7px 14px;border-radius:7px;font-size:13px;cursor:pointer;font-family:inherit;
}
button:hover{border-color:var(--accent);color:var(--accent)}
.theme-btn{position:fixed;top:16px;right:20px;z-index:99;box-shadow:var(--shadow)}
footer{
  margin-top:60px;padding-top:20px;border-top:1px solid var(--line);
  font-size:12px;color:var(--fg2);
}
@media(max-width:900px){
  nav{display:none}
  main{padding:20px 16px 60px}
  .bar .nm{width:90px;flex:0 0 90px}
}
@media print{
  nav,.theme-btn,.toolbar{display:none}
  body{background:#fff;color:#000}
  main{max-width:none;padding:0}
  .tw{break-inside:avoid}
  h2{break-after:avoid}
}
</style>
</head>
<body data-theme="dark">
<button class="theme-btn" onclick="toggleTheme()">🌓 切換主題</button>
<div class="wrap">
<nav>
  <h1>ROKAIZUMI ATELIER</h1>
  <div class="sub">頂級化妝品線<br>完整資料儀表板<br><span id="gen"></span></div>
  <a href="#s1" class="on">1　總覽</a>
  <a href="#s2">2　14 支 SKU</a>
  <a href="#s3">3　成本分析</a>
  <a href="#s4">4　原料明細</a>
  <a href="#s5">5　價格來源</a>
  <a href="#s6">6　敏感度</a>
  <a href="#s7">7　定價檢查</a>
  <a href="#s8">8　投資明細</a>
  <a href="#s9">9　化工廠合作</a>
  <a href="#s10">10　AI Agent 團隊</a>
  <a href="#s11">11　損益與回收</a>
  <a href="#s12">12　獲利分配</a>
  <a href="#s13">13　法遵邊界</a>
  <a href="#s14">14　STO＋／RWA</a>
</nav>
<main>

<h2 id="s1"><span class="n">1</span>總覽</h2>
<p class="lead">Rokaizumi 合同會社（品牌：ROKAIZUMI ATELIER／露海澄研）頂級保養品產品線。代工廠：台灣安成生化科技有限公司。所有 SKU 目前 <code>enabled=false</code>，未經林總核定不得對外使用。</p>
<div class="kpis" id="kpis"></div>
<div class="note warn">
<b>本頁所有數字皆為估算值，非報價、非承諾。</b>
原料價格來自公開網路抓取與業界參考價，實際成本以代工廠正式 RFQ 回覆為準。
</div>

<h2 id="s2"><span class="n">2</span>14 支 SKU</h2>
<p class="lead">六大系列：澄（清潔）／潤（保濕修護）／曜（美白防護）／齡（抗老）／御（頂級油與禮盒）／沁（頭皮）。分三批開發，首批 4 支 P0。</p>
<div class="skus" id="skus"></div>

<h2 id="s3"><span class="n">3</span>成本分析</h2>
<p class="lead">逐支成本結構（TWD／支）。點欄位標題可排序。</p>
<div class="tw"><table id="t-cost">
<thead><tr>
<th class="s" data-k="sku_code">SKU</th>
<th class="s" data-k="name_zh">品名</th>
<th class="s n" data-k="size_ml">容量</th>
<th class="s n" data-k="raw_material_twd">料體</th>
<th class="s n" data-k="packaging_twd">包材</th>
<th class="s n" data-k="filling_twd">充填</th>
<th class="s n" data-k="qc_twd">檢驗</th>
<th class="s n" data-k="exfactory_twd">出廠成本</th>
<th class="s n" data-k="total_cost_twd">總成本</th>
<th class="s n" data-k="suggested_retail_twd">建議售價</th>
<th class="s n" data-k="gross_margin_pct">毛利率</th>
</tr></thead><tbody></tbody></table></div>
<h3>成本結構佔比</h3>
<div id="struct"></div>

<h2 id="s4"><span class="n">4</span>原料明細</h2>
<p class="lead">248 筆逐 SKU × 逐原料成本明細。可搜尋原料名稱或篩選 SKU。</p>
<div class="toolbar">
  <input type="text" id="q" placeholder="搜尋原料名稱（例如 Hyaluronate）">
  <select id="fsku"><option value="">全部 SKU</option></select>
  <select id="flayer">
    <option value="">全部價格層級</option>
    <option value="L1_scraped">L1 自動抓取</option>
    <option value="L2_ref">L2 參考價</option>
    <option value="L3_manual">L3 人工覆寫</option>
  </select>
  <span id="cnt" style="font-size:12.5px;color:var(--fg2)"></span>
</div>
<div class="tw" style="max-height:600px;overflow-y:auto"><table id="t-detail">
<thead><tr>
<th>SKU</th><th>原料（INCI）</th><th>功能類別</th>
<th class="n">比例 %</th><th class="n">單價 USD/kg</th>
<th>價格層級</th><th class="n">每 kg 料體成本 TWD</th><th>備註</th>
</tr></thead><tbody></tbody></table></div>

<h2 id="s5"><span class="n">5</span>價格來源</h2>
<p class="lead">三層價格回退架構。L1 自動抓取 makingcosmetics.com 公開價；L2 業界參考價；L3 代工廠報價（優先度最高）。</p>
<div class="kpis" id="layer-kpi"></div>
<h3>L1 抓取失敗清單</h3>
<p class="lead">以下原料在 makingcosmetics.com 找不到對應商品，已自動回退到 L2 參考價。</p>
<div class="tw" style="max-height:400px;overflow-y:auto"><table id="t-fail">
<thead><tr><th>原料（INCI）</th><th>失敗原因</th></tr></thead><tbody></tbody></table></div>

<h2 id="s6"><span class="n">6</span>敏感度分析</h2>
<p class="lead">料體價格變動 ±15%／±30% 對總成本的影響。</p>
<div class="tw"><table id="t-sens">
<thead><tr><th>SKU</th><th>品名</th><th class="n">基準</th>
<th class="n">-30%</th><th class="n">-15%</th><th class="n">+15%</th><th class="n">+30%</th>
<th class="n">最大變動</th></tr></thead><tbody></tbody></table></div>
<div class="note ok"><b>結論：</b>即使原料漲 30%，總成本最多增加約 15%。成本控制的槓桿在<b>包材與批量</b>，不在原料議價。</div>

<h2 id="s7"><span class="n">7</span>定價合理性檢查</h2>
<p class="lead">零售價 ÷ 出廠成本。精品保養品常見區間為 6–12 倍（含品牌行銷、通路抽成、退換貨）。</p>
<div class="tw"><table id="t-price">
<thead><tr><th>SKU</th><th>品名</th><th class="n">出廠成本</th>
<th class="n">原提案價</th><th class="n">倍率</th>
<th class="n">建議價（7 倍）</th><th class="n">倍率</th><th>判定</th></tr></thead><tbody></tbody></table></div>
<div class="note danger"><b>重要發現：</b>原提案價多數超過 20 倍，最高達 36 倍。這代表必須有極強的品牌力與行銷預算支撐，否則通路會要求更高折讓、或消費者直接轉向競品。</div>

<h2 id="s8"><span class="n">8</span>投資明細</h2>
<p class="lead">第一年一次性開發投資（TWD）。</p>
<div class="kpis" id="inv-kpi"></div>
<div class="tw"><table id="t-inv">
<thead><tr><th>類別</th><th>項目</th><th class="n">金額 TWD</th><th>說明</th></tr></thead>
<tbody></tbody><tfoot></tfoot></table></div>
<h3>各類別佔比</h3>
<div id="inv-bar"></div>

<h2 id="s9"><span class="n">9</span>化工廠合作</h2>
<p class="lead" id="fac-lead"></p>
<div class="note ok" id="fac-note"></div>
<h3>代工廠投入（不出資金、不佔股）</h3>
<div class="tw"><table id="t-fac1"><thead><tr><th>項目</th><th>內容</th><th>說明</th></tr></thead><tbody></tbody></table></div>
<h3>品牌方負擔費用</h3>
<div class="tw"><table id="t-fac2"><thead><tr><th>項目</th><th>金額</th><th>說明</th></tr></thead><tbody></tbody></table></div>
<h3>建議合約條件</h3>
<div class="tw"><table id="t-fac3"><thead><tr><th>條款</th><th>建議內容</th><th>說明</th></tr></thead><tbody></tbody></table></div>
<h3>風險提醒</h3>
<ul id="fac-risk"></ul>

<h2 id="s10"><span class="n">10</span>AI Agent 團隊</h2>
<p class="lead">20 個 Agent 合併成 6 個值班入口（E1～E6，依 D-029）。每月 AI 預算 USD 60。</p>
<div class="kpis" id="ai-kpi"></div>
<div class="tw"><table id="t-ai">
<thead><tr><th>入口</th><th>職能</th><th>對應 Agent</th><th>替代人力</th>
<th class="n">人力月成本</th><th class="n">AI 月成本</th><th>主要職責</th></tr></thead>
<tbody></tbody><tfoot></tfoot></table></div>
<div class="note warn" id="ai-note"></div>

<h2 id="s11"><span class="n">11</span>損益與回收</h2>
<p class="lead">兩情境對照：原提案價 vs 建議價（出廠成本 × 7）。通路假設 70% 自營電商、30% 實體通路。</p>
<div class="kpis" id="pnl-kpi"></div>
<h3>第一年損益表</h3>
<div class="tw"><table id="t-pnl">
<thead><tr><th>項目</th><th class="n">原提案價</th><th class="n">建議價</th><th>說明</th></tr></thead><tbody></tbody></table></div>
<h3>第二年損益表（14 支全開）</h3>
<div class="tw"><table id="t-pnl2">
<thead><tr><th>項目</th><th class="n">原提案價</th><th class="n">建議價</th><th>說明</th></tr></thead><tbody></tbody></table></div>
<h3>投資回收期</h3>
<div class="tw"><table id="t-pb">
<thead><tr><th>情境</th><th class="n">年淨利 TWD</th><th class="n">回收期（年）</th><th>判定</th></tr></thead><tbody></tbody></table></div>

<h2 id="s12"><span class="n">12</span>獲利分配</h2>
<p class="lead">以建議價計算，零售價 100% 的價值鏈分配。混合通路假設：70% 自營、30% 實體通路。</p>
<div id="wf"></div>
<div class="note"><b>怎麼讀：</b>通路抽成吃掉最大一塊（約 15%），其次是產品成本（約 12%）。品牌稅後淨利約 12–15%。這是精品保養品的常態結構。</div>

<h2 id="s13"><span class="n">13</span>法遵邊界</h2>
<div class="note danger"><b>本專案不照抄任何他人產品線。</b>原始需求涉及照抄 tenlead.com（天麗生技／Joli 系列）的品名與成分比例，已改以「品類對標＋原創命名＋公開知識配方架構＋正式 RFQ」的合法路徑執行。</div>
<h3>七項法律風險</h3>
<div class="tw"><table id="t-law">
<thead><tr><th>法規</th><th>條文</th><th>風險行為</th><th>後果</th></tr></thead><tbody></tbody></table></div>
<h3>執行紅線</h3>
<ol>
<li>不使用任何他人品牌名、商標、產品名、文案、圖片。</li>
<li>不對外宣稱「同款」「平替」「同配方」。</li>
<li>成分比例為我方自行設定的起始架構，非任何公司實際配方。</li>
<li>上市前每一支都需法遵（A04）逐項核對法規限量與宣稱。</li>
<li>對外送件、簽約、付費一律先經林總核定。</li>
</ol>

<h2 id="s14"><span class="n">14</span>STO＋／RWA</h2>
<p class="lead">每批商品可產出 RWA 批次憑證，接 STO＋ 多狀態模型。</p>
<h3>RWA 六類資產對映</h3>
<div class="tw"><table id="t-rwa">
<thead><tr><th>資產類別</th><th>本專案對應</th><th>憑證內容</th></tr></thead><tbody></tbody></table></div>
<h3>關鍵決策</h3>
<div class="tw"><table id="t-dec">
<thead><tr><th>編號</th><th>決策</th><th>對本專案的影響</th></tr></thead><tbody></tbody></table></div>

<footer>
<p><b>Rokaizumi 頂級化妝品線 — 完整資料儀表板</b></p>
<p>產出時間：<span id="gen2"></span>　｜　所有 SKU <code>enabled=false</code>　｜　高影響操作需林總核定</p>
<p>本頁為單檔自包含 HTML，無外部依賴。可雙擊開啟、可列印、可上傳任何免費靜態託管。</p>
</footer>
</main>
</div>

<script>
const D = __DATA__;
const F = D.fin;
const fmt = n => (n===null||n===undefined||n==='')?'—':Number(n).toLocaleString('zh-TW');
const fmt0 = n => (n===null||n===undefined||n==='')?'—':Math.round(Number(n)).toLocaleString('zh-TW');
const pct = n => (n===null||n===undefined||n==='')?'—':Number(n).toFixed(1)+'%';
const el = (t,c,h)=>{const e=document.createElement(t);if(c)e.className=c;if(h!==undefined)e.innerHTML=h;return e;};
const $ = s=>document.querySelector(s);

document.getElementById('gen').textContent = F.generated_at.slice(0,10);
document.getElementById('gen2').textContent = F.generated_at.replace('T',' ').slice(0,19);

function toggleTheme(){
  const b=document.body;
  b.dataset.theme = b.dataset.theme==='dark'?'light':'dark';
}

/* ---------- 1 總覽 ---------- */
(function(){
  const inv = F.investment.total_twd;
  const y1o = F.pnl.y1_original, y1s = F.pnl.y1_suggested;
  const kp = [
    ['b','SKU 總數','14','支','六大系列，首批 4 支 P0'],
    ['g','一次性投資',fmt0(inv),'TWD','第一年開發＋首批量產'],
    ['p','第一年營收（建議價）',fmt0(y1s.net_revenue),'TWD','品牌實收，已扣通路抽成'],
    ['g','第一年淨利（建議價）',fmt0(y1s.net_profit),'TWD','淨利率 '+pct(y1s.net_margin_pct)],
    ['w','回收期（建議價）',F.payback.y1_suggested_years,'年','原提案價為 '+F.payback.y1_original_years+' 年'],
    ['b','AI 團隊年省',fmt0(F.ai_team.summary.saving_year_twd),'TWD','4 人降至 2 人 + AI'],
  ];
  const box=$('#kpis');
  kp.forEach(([c,l,v,u,n])=>{
    box.appendChild(el('div','kpi '+c,
      `<div class="lbl">${l}</div><div class="val">${v}<span class="unit">${u}</span></div><div class="note">${n}</div>`));
  });
})();

/* ---------- 2 SKU 卡 ---------- */
(function(){
  const box=$('#skus');
  const cmap={}; D.cost.forEach(r=>cmap[r.sku_code]=r);
  D.products.forEach(p=>{
    const c=cmap[p.sku_code]||{};
    const pr=(p.priority||'').toLowerCase();
    const d=el('div','sku '+pr);
    d.innerHTML=`
      <div class="hd">
        <div><div class="code">${p.sku_code}</div><div class="nm">${p.name_zh}</div></div>
        <span class="tag t-${pr}">${p.priority||'—'}</span>
      </div>
      <div class="meta">${p.form||''}　${p.size_ml?p.size_ml+' mL':(p.pack||'組合')}</div>
      <div class="row"><span>建議售價</span><b>TWD ${fmt(p.suggested_retail_twd)}</b></div>
      <div class="row"><span>出廠成本</span><b>TWD ${fmt0(c.exfactory_twd)}</b></div>
      <div class="row"><span>總成本</span><b>TWD ${fmt0(c.total_cost_twd)}</b></div>
      <div class="row"><span>零售／出廠倍率</span><b>${c.exfactory_twd?(p.suggested_retail_twd/c.exfactory_twd).toFixed(1)+'×':'—'}</b></div>
      <div class="row"><span>狀態</span><b style="color:var(--warn)">enabled=false</b></div>`;
    box.appendChild(d);
  });
})();

/* ---------- 3 成本表 ---------- */
(function(){
  const tb=$('#t-cost tbody');
  const cols=['sku_code','name_zh','size_ml','raw_material_twd','packaging_twd',
    'filling_twd','qc_twd','exfactory_twd','total_cost_twd','suggested_retail_twd','gross_margin_pct'];
  function render(rows){
    tb.innerHTML='';
    rows.forEach(r=>{
      const tr=el('tr');
      tr.innerHTML=`<td><code>${r.sku_code}</code></td><td>${r.name_zh}</td>
        <td class="n">${r.size_ml||'—'}</td>
        <td class="n">${fmt0(r.raw_material_twd)}</td>
        <td class="n">${fmt0(r.packaging_twd)}</td>
        <td class="n">${fmt0(r.filling_twd)}</td>
        <td class="n">${fmt0(r.qc_twd)}</td>
        <td class="n"><b>${fmt0(r.exfactory_twd)}</b></td>
        <td class="n">${fmt0(r.total_cost_twd)}</td>
        <td class="n">${fmt(r.suggested_retail_twd)}</td>
        <td class="n">${pct(r.gross_margin_pct)}</td>`;
      tb.appendChild(tr);
    });
  }
  render(D.cost);
  let dir={};
  document.querySelectorAll('#t-cost th.s').forEach(th=>{
    th.onclick=()=>{
      const k=th.dataset.k; dir[k]=!dir[k];
      const rows=[...D.cost].sort((a,b)=>{
        const x=a[k],y=b[k];
        const nx=parseFloat(x),ny=parseFloat(y);
        const v=(!isNaN(nx)&&!isNaN(ny))?nx-ny:String(x).localeCompare(String(y),'zh-TW');
        return dir[k]?v:-v;
      });
      render(rows);
    };
  });
})();

/* ---------- 3b 成本結構 ---------- */
(function(){
  const box=$('#struct');
  D.cost.forEach(r=>{
    const t=Number(r.total_cost_twd)||1;
    const raw=Number(r.raw_material_twd)/t*100;
    const pack=Number(r.packaging_twd)/t*100;
    const other=100-raw-pack;
    const d=el('div','bar');
    d.innerHTML=`<div class="nm">${r.sku_code}</div>
      <div class="tr" style="display:flex">
        <div class="fl" style="width:${raw}%;background:var(--accent)"></div>
        <div class="fl" style="width:${pack}%;background:var(--purple)"></div>
        <div class="fl" style="width:${other}%;background:var(--bg3)"></div>
      </div>
      <div class="vl">料${raw.toFixed(0)}% 包${pack.toFixed(0)}%</div>`;
    box.appendChild(d);
  });
  box.appendChild(el('div','note',
    '<b>圖例：</b><span style="color:var(--accent)">■</span> 料體　'+
    '<span style="color:var(--purple)">■</span> 包材　<span style="color:var(--fg2)">■</span> 充填＋檢驗＋管銷'));
})();

/* ---------- 4 原料明細 ---------- */
(function(){
  const tb=$('#t-detail tbody');
  const sel=$('#fsku');
  [...new Set(D.detail.map(r=>r.sku_code))].sort().forEach(s=>{
    const o=el('option'); o.value=s; o.textContent=s; sel.appendChild(o);
  });
  const layerName={L1_scraped:'L1 自動抓取',L2_ref:'L2 參考價',L3_manual:'L3 人工覆寫',fallback:'fallback'};
  const layerCls={L1_scraped:'t-l1',L2_ref:'t-l2',L3_manual:'t-l3',fallback:'t-bad'};
  function render(){
    const q=$('#q').value.toLowerCase().trim();
    const fs=sel.value, fl=$('#flayer').value;
    const rows=D.detail.filter(r=>
      (!q||r.inci_name.toLowerCase().includes(q)) &&
      (!fs||r.sku_code===fs) && (!fl||r.price_layer===fl));
    tb.innerHTML='';
    rows.forEach(r=>{
      const tr=el('tr');
      tr.innerHTML=`<td><code>${r.sku_code}</code></td><td>${r.inci_name}</td>
        <td>${r.function_category||'—'}</td>
        <td class="n">${Number(r.pct).toFixed(3)}</td>
        <td class="n">${fmt(r.usd_per_kg)}</td>
        <td><span class="tag ${layerCls[r.price_layer]||'t-p2'}">${layerName[r.price_layer]||r.price_layer}</span></td>
        <td class="n">${fmt(r.cost_twd_per_kg_bulk)}</td>
        <td style="font-size:12px;color:var(--fg2)">${r.note||''}</td>`;
      tb.appendChild(tr);
    });
    $('#cnt').textContent=`顯示 ${rows.length} / ${D.detail.length} 筆`;
  }
  $('#q').oninput=render; sel.onchange=render; $('#flayer').onchange=render;
  render();
})();

/* ---------- 5 價格來源 ---------- */
(function(){
  const lc=D.layerCount;
  const total=Object.values(lc).reduce((a,b)=>a+b,0)||1;
  const box=$('#layer-kpi');
  const defs=[
    ['L1_scraped','L1 自動抓取','makingcosmetics.com 公開價','b'],
    ['L2_ref','L2 內建參考價','業界行情中位數','p'],
    ['L3_manual','L3 人工覆寫','代工廠正式報價','g'],
    ['fallback','fallback','無資料保守值','r'],
  ];
  defs.forEach(([k,l,d,c])=>{
    const v=lc[k]||0;
    box.appendChild(el('div','kpi '+c,
      `<div class="lbl">${l}</div><div class="val">${v}<span class="unit">筆</span></div>
       <div class="note">${d}　${(v/total*100).toFixed(1)}%</div>`));
  });
  const tb=$('#t-fail tbody');
  D.failed.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td>${r.inci}</td><td style="color:var(--fg2)">${r.reason}</td>`;
    tb.appendChild(tr);
  });
})();

/* ---------- 6 敏感度 ---------- */
(function(){
  const tb=$('#t-sens tbody');
  D.sens.forEach(r=>{
    const mx=Math.max(Math.abs(r.d30-r.base),Math.abs(r.dm30-r.base));
    const tr=el('tr');
    tr.innerHTML=`<td><code>${r.sku}</code></td><td>${r.name}</td>
      <td class="n">${fmt0(r.base)}</td>
      <td class="n">${fmt0(r.dm30)}</td><td class="n">${fmt0(r.dm15)}</td>
      <td class="n">${fmt0(r.d15)}</td><td class="n">${fmt0(r.d30)}</td>
      <td class="n">±${(mx/r.base*100).toFixed(1)}%</td>`;
    tb.appendChild(tr);
  });
})();

/* ---------- 7 定價檢查 ---------- */
(function(){
  const tb=$('#t-price tbody');
  const om={}, sm={};
  F.pricing.original.forEach(r=>om[r.sku_code]=r);
  F.pricing.suggested.forEach(r=>sm[r.sku_code]=r);
  F.pricing.original.forEach(o=>{
    const s=sm[o.sku_code];
    let verdict,cls;
    if(o.multiple>20){verdict='顯著偏高';cls='t-bad';}
    else if(o.multiple>12){verdict='偏高';cls='t-p1';}
    else if(o.multiple<4){verdict='偏低';cls='t-p1';}
    else {verdict='合理';cls='t-ok';}
    const tr=el('tr');
    tr.innerHTML=`<td><code>${o.sku_code}</code></td><td>${o.name_zh}</td>
      <td class="n">${fmt0(o.exfactory)}</td>
      <td class="n">${fmt(o.price)}</td>
      <td class="n">${o.multiple}×</td>
      <td class="n" style="color:var(--accent2)">${fmt(s.price)}</td>
      <td class="n" style="color:var(--accent2)">${s.multiple}×</td>
      <td><span class="tag ${cls}">${verdict}</span></td>`;
    tb.appendChild(tr);
  });
})();

/* ---------- 8 投資明細 ---------- */
(function(){
  const inv=F.investment;
  const box=$('#inv-kpi');
  box.appendChild(el('div','kpi g',
    `<div class="lbl">一次性投資總額</div><div class="val">${fmt0(inv.total_twd)}<span class="unit">TWD</span></div>
     <div class="note">約 ${(inv.total_twd/10000).toFixed(0)} 萬元</div>`));
  inv.by_category.forEach(c=>{
    box.appendChild(el('div','kpi b',
      `<div class="lbl">${c.category}</div><div class="val">${fmt0(c.amount_twd)}<span class="unit">TWD</span></div>
       <div class="note">${(c.amount_twd/inv.total_twd*100).toFixed(1)}%</div>`));
  });
  const tb=$('#t-inv tbody');
  inv.rows.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td>${r.category}</td><td>${r.item}</td>
      <td class="n">${fmt(r.amount_twd)}</td>
      <td style="font-size:12px;color:var(--fg2)">${r.note}</td>`;
    tb.appendChild(tr);
  });
  $('#t-inv tfoot').innerHTML=`<tr><td colspan="2">合計</td>
    <td class="n">${fmt(inv.total_twd)}</td><td></td></tr>`;
  const bar=$('#inv-bar');
  const mx=Math.max(...inv.by_category.map(c=>c.amount_twd));
  inv.by_category.forEach(c=>{
    const d=el('div','bar');
    d.innerHTML=`<div class="nm">${c.category}</div>
      <div class="tr"><div class="fl" style="width:${c.amount_twd/mx*100}%;background:var(--accent)"></div></div>
      <div class="vl">${fmt0(c.amount_twd)}</div>`;
    bar.appendChild(d);
  });
})();

/* ---------- 9 化工廠 ---------- */
(function(){
  const f=F.factory;
  $('#fac-lead').textContent=`受託方：${f.name}　｜　角色：${f.role}　｜　代工廠出資金額：TWD ${fmt(f.factory_cash_investment_twd)}`;
  $('#fac-note').innerHTML='<b>核心原則：</b>'+f.factory_note;
  const fill=(sel,rows)=>{
    const tb=$(sel+' tbody');
    rows.forEach(r=>{
      const tr=el('tr');
      tr.innerHTML=`<td>${r[0]}</td><td>${r[1]}</td><td style="color:var(--fg2);font-size:12.5px">${r[2]}</td>`;
      tb.appendChild(tr);
    });
  };
  fill('#t-fac1',f.factory_contribution);
  fill('#t-fac2',f.brand_side_pays);
  fill('#t-fac3',f.contract_terms);
  const ul=$('#fac-risk');
  f.risk_notes.forEach(n=>ul.appendChild(el('li',null,n)));
})();

/* ---------- 10 AI 團隊 ---------- */
(function(){
  const s=F.ai_team.summary;
  const box=$('#ai-kpi');
  const kp=[
    ['b','Agent 總數',s.agent_count,'個','合併成 '+s.entry_count+' 個值班入口'],
    ['w','純人力年成本',fmt0(s.full_human_year_twd),'TWD',s.full_human_headcount+' 人 × 月 '+fmt(s.full_human_month_twd)],
    ['g','AI 年成本',fmt0(s.ai_year_twd),'TWD','USD '+s.ai_month_usd+'/月'],
    ['p','精簡後總成本',fmt0(s.lean_total_year_twd),'TWD',s.lean_headcount+' 人 + AI'],
    ['g','年節省',fmt0(s.saving_year_twd),'TWD','節省 '+pct(s.saving_pct)],
  ];
  kp.forEach(([c,l,v,u,n])=>{
    box.appendChild(el('div','kpi '+c,
      `<div class="lbl">${l}</div><div class="val">${v}<span class="unit">${u}</span></div><div class="note">${n}</div>`));
  });
  const tb=$('#t-ai tbody');
  F.ai_team.rows.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td><b>${r.code}</b></td><td>${r.name}</td><td style="font-size:12px">${r.agents}</td>
      <td>${r.human_role}</td>
      <td class="n">${fmt(r.human_cost_twd)}</td>
      <td class="n" style="color:var(--accent2)">${r.ai_cost_twd}</td>
      <td style="font-size:12px;color:var(--fg2)">${r.duty}</td>`;
    tb.appendChild(tr);
  });
  $('#t-ai tfoot').innerHTML=`<tr><td colspan="4">合計</td>
    <td class="n">${fmt(s.full_human_month_twd)}</td>
    <td class="n">${fmt(s.ai_month_twd)}</td><td></td></tr>`;
  $('#ai-note').innerHTML='<b>誠實說明：</b>'+s.note;
})();

/* ---------- 11 損益 ---------- */
(function(){
  const y1o=F.pnl.y1_original, y1s=F.pnl.y1_suggested;
  const y2o=F.pnl.y2_original, y2s=F.pnl.y2_suggested;
  const box=$('#pnl-kpi');
  const kp=[
    ['b','第一年營收（原價）',fmt0(y1o.net_revenue),'TWD','品牌實收'],
    ['b','第一年營收（建議價）',fmt0(y1s.net_revenue),'TWD','品牌實收'],
    ['g','第一年淨利（原價）',fmt0(y1o.net_profit),'TWD','淨利率 '+pct(y1o.net_margin_pct)],
    ['g','第一年淨利（建議價）',fmt0(y1s.net_profit),'TWD','淨利率 '+pct(y1s.net_margin_pct)],
    ['p','第二年淨利（建議價）',fmt0(y2s.net_profit),'TWD','14 支全開'],
  ];
  kp.forEach(([c,l,v,u,n])=>{
    box.appendChild(el('div','kpi '+c,
      `<div class="lbl">${l}</div><div class="val">${v}<span class="unit">${u}</span></div><div class="note">${n}</div>`));
  });

  const notes={
    '產品成本（出廠）':'含料體、包材、充填、檢驗、代工廠毛利',
    '物流與倉儲':'佔品牌實收 4%',
    '行銷':'佔品牌實收 18%',
    '人事':'2 人 × 月 45,000 × 12',
    '平台與金流':'自營部分 5%',
    '代理佣金':'通路部分 12%',
    '社群貢獻回饋':'佔品牌實收 3%，與代理佣金分開記帳',
    '雜支':'佔品牌實收 3%',
    'AI Agent 維運':'USD 60/月 × 12',
  };
  function fillPnl(sel,o,s){
    const tb=$(sel+' tbody');
    tb.innerHTML='';
    const add=(label,a,b,note,cls)=>{
      const tr=el('tr');
      tr.innerHTML=`<td${cls?' style="font-weight:700"':''}>${label}</td>
        <td class="n"${cls?' style="font-weight:700"':''}>${fmt0(a)}</td>
        <td class="n"${cls?' style="font-weight:700"':''}>${fmt0(b)}</td>
        <td style="font-size:12px;color:var(--fg2)">${note||''}</td>`;
      tb.appendChild(tr);
    };
    add('零售總額（建議售價 × 銷量）',o.gross_retail_revenue,s.gross_retail_revenue,'未扣通路抽成');
    add('品牌實收營收',o.net_revenue,s.net_revenue,'已扣通路抽成與金流手續費',1);
    Object.keys(o.expenses).forEach(k=>{
      add('　− '+k,o.expenses[k],s.expenses[k],notes[k]||'');
    });
    add('費用合計',o.total_expenses,s.total_expenses,'',1);
    add('稅前淨利',o.ebt,s.ebt,'');
    add('營所稅（20%）',o.tax,s.tax,'');
    add('稅後淨利',o.net_profit,s.net_profit,'淨利率 '+pct(o.net_margin_pct)+' / '+pct(s.net_margin_pct),1);
  }
  fillPnl('#t-pnl',y1o,y1s);
  fillPnl('#t-pnl2',y2o,y2s);

  const tb=$('#t-pb tbody');
  const rows=[
    ['第一年｜原提案價',y1o.net_profit,F.payback.y1_original_years],
    ['第一年｜建議價',y1s.net_profit,F.payback.y1_suggested_years],
    ['第二年｜原提案價',y2o.net_profit,F.payback.y2_original_years],
    ['第二年｜建議價',y2s.net_profit,F.payback.y2_suggested_years],
  ];
  rows.forEach(([l,p,y])=>{
    let v,cls;
    if(y===null){v='無法回收';cls='t-bad';}
    else if(y<=1.5){v='快速回收';cls='t-ok';}
    else if(y<=3){v='合理';cls='t-ok';}
    else {v='偏長';cls='t-p1';}
    const tr=el('tr');
    tr.innerHTML=`<td>${l}</td><td class="n">${fmt0(p)}</td>
      <td class="n">${y===null?'—':y}</td>
      <td><span class="tag ${cls}">${v}</span></td>`;
    tb.appendChild(tr);
  });
})();

/* ---------- 12 獲利分配瀑布 ---------- */
(function(){
  const box=$('#wf');
  const colors=['var(--danger)','var(--accent)','var(--warn)','var(--fg2)',
    'var(--purple)','var(--pink)','var(--accent2)','var(--fg2)','var(--warn)','var(--accent2)'];
  F.waterfall.forEach(w=>{
    box.appendChild(el('h3',null,`${w.sku_code}　${w.name_zh}　零售價 TWD ${fmt(w.retail_price)}`));
    const mx=Math.max(...w.items.map(i=>i.amount));
    w.items.forEach((it,i)=>{
      const d=el('div','bar');
      d.innerHTML=`<div class="nm">${it.name}</div>
        <div class="tr"><div class="fl" style="width:${it.amount/mx*100}%;background:${colors[i%colors.length]}"></div></div>
        <div class="vl">${fmt(it.amount)}　${it.pct}%</div>`;
      box.appendChild(d);
      if(it.note){
        const n=el('div',null,`<span style="font-size:11.5px;color:var(--fg2);margin-left:160px">${it.note}</span>`);
        box.appendChild(n);
      }
    });
  });
})();

/* ---------- 13 法遵 ---------- */
(function(){
  const rows=[
    ['商標法','第 68 條','使用相同或近似他人註冊商標於同一或類似商品','民事賠償、刑事責任'],
    ['商標法','第 95 條','意圖販賣而持有、陳列仿冒商標商品','3 年以下有期徒刑'],
    ['著作權法','第 91 條','重製他人語文、美術、攝影著作（文案、圖片、包裝設計）','3 年以下有期徒刑、罰金'],
    ['營業秘密法','第 10 條','以不正方法取得、使用他人配方與製程','民事賠償、刑事責任'],
    ['營業秘密法','第 13-1 條','意圖為自己或第三人不法之利益','1 年以上 10 年以下有期徒刑'],
    ['公平交易法','第 21 條','商品品質、內容為虛偽不實或引人錯誤之表示','罰鍰 5 萬至 2,500 萬元'],
    ['公平交易法','第 22 條','以著名表徵為相同或類似之使用，致混淆','罰鍰、民事賠償'],
    ['公平交易法','第 25 條','足以影響交易秩序之欺罔或顯失公平行為','罰鍰、民事賠償'],
    ['化粧品衛生安全管理法','第 7、8 條','未經登錄、標示不實、成分不符限量','罰鍰、沒入銷毀'],
    ['化粧品衛生安全管理法','第 22 條','宣稱醫療效能','罰鍰 60 萬至 500 萬元'],
    ['民法','第 185 條','代工廠共同參與侵權行為','連帶損害賠償'],
  ];
  const tb=$('#t-law tbody');
  rows.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td>${r[0]}</td><td>${r[1]}</td><td>${r[2]}</td>
      <td style="color:var(--danger);font-size:12.5px">${r[3]}</td>`;
    tb.appendChild(tr);
  });
})();

/* ---------- 14 STO＋／RWA ---------- */
(function(){
  const rwa=[
    ['實體商品批次','每批生產的成品','批號、SKU、生產日、效期、檢驗報告雜湊'],
    ['原料來源憑證','關鍵活性成分','供應商、COA、產地、檢驗結果'],
    ['品質檢驗紀錄','每批出廠檢驗','微生物、重金屬、安定性、留樣編號'],
    ['法規登錄文件','產品登錄與 PIF','登錄字號、PIF 版本、宣稱白名單'],
    ['通路銷售紀錄','出貨與銷售','出貨批號、通路、數量（不含個資）'],
    ['社群貢獻紀錄','消費者與代理貢獻','貢獻類型、點數、與佣金分開記帳'],
  ];
  const tb=$('#t-rwa tbody');
  rwa.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td><b>${r[0]}</b></td><td>${r[1]}</td>
      <td style="font-size:12.5px;color:var(--fg2)">${r[2]}</td>`;
    tb.appendChild(tr);
  });
  const dec=[
    ['D-014','鏈上公開 payload 只含 merkle_root 與 leaf_count','配方與檢驗明細不上鏈，只上雜湊根'],
    ['D-017','消費者端先「不可轉讓、不承諾轉 ST」','避免被認定為有價證券'],
    ['D-018','第一階段面對消費者的市場為台灣＋日本','七國同步為長期目標，非首發'],
    ['D-020','照片、個資、健康資料一律不送境外 API','AI Agent 處理客服時須遵守'],
    ['D-021','代理制度先只做單層推薦','避免多層次傳銷認定風險'],
    ['D-025','高影響操作仍不自行執行','公開發布、付費、簽約、送件、正式發行需林總核定'],
    ['D-029','20 個 Agent 合併成 6 個值班入口；每月 AI 預算 USD 60','本頁第 10 節的成本依據'],
  ];
  const tb2=$('#t-dec tbody');
  dec.forEach(r=>{
    const tr=el('tr');
    tr.innerHTML=`<td><code>${r[0]}</code></td><td>${r[1]}</td>
      <td style="font-size:12.5px;color:var(--fg2)">${r[2]}</td>`;
    tb2.appendChild(tr);
  });
})();

/* ---------- 導覽高亮 ---------- */
(function(){
  const links=[...document.querySelectorAll('nav a')];
  const secs=links.map(a=>document.querySelector(a.getAttribute('href'))).filter(Boolean);
  const obs=new IntersectionObserver(es=>{
    es.forEach(e=>{
      if(e.isIntersecting){
        links.forEach(l=>l.classList.remove('on'));
        const i=secs.indexOf(e.target);
        if(i>=0) links[i].classList.add('on');
      }
    });
  },{rootMargin:'-20% 0px -70% 0px'});
  secs.forEach(s=>obs.observe(s));
})();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    main()
