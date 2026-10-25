#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
原料成本計算引擎（三層價格來源）

L1 自動抓取  data/ingredient_prices.json   （makingcosmetics.com 公開零售/批發價）
L2 內建參考價 data/ingredient_master.json   （ref_price_usd_per_kg，業界行情）
L3 人工覆寫  data/price_overrides.json     （代工廠正式報價，優先度最高）

輸入
  data/products.json          14 支 SKU 主檔
  data/ingredients.csv        248 列成分明細（pct_min / pct_max / q.s.）
  data/ingredient_master.json 密度、參考價
  data/ingredient_prices.json L1 抓取結果
  data/price_overrides.json   L3 覆寫（可不存在）

輸出
  data/cost_detail.csv   逐 SKU × 逐原料的成本明細
  data/cost_report.csv   逐 SKU 成本彙總
  data/cost_report.md    人可讀報告（含敏感度分析）

用法
  python scripts/cost_engine.py
  python scripts/cost_engine.py --usd-twd 32.5
"""

import argparse
import csv
import json
import os
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")

P_PRODUCTS = os.path.join(DATA, "products.json")
P_INGRED = os.path.join(DATA, "ingredients.csv")
P_MASTER = os.path.join(DATA, "ingredient_master.json")
P_PRICES = os.path.join(DATA, "ingredient_prices.json")
P_OVERRIDE = os.path.join(DATA, "price_overrides.json")

OUT_DETAIL = os.path.join(DATA, "cost_detail.csv")
OUT_REPORT = os.path.join(DATA, "cost_report.csv")
OUT_MD = os.path.join(DATA, "cost_report.md")

# ---------- 成本參數（可調） ----------
CFG = {
    "usd_twd": 32.0,          # 匯率 USD → TWD
    "loss_rate": 0.03,        # 製程損耗 3%
    "qc_per_unit_twd": 6.0,   # 每支分攤檢驗費（微生物／安定性／重金屬）
    "factory_margin_pct": 0.20,   # 代工廠管銷＋毛利
    "logistics_pct": 0.08,    # 品牌端物流／倉儲／管銷
    "channel_multiplier": 2.0,    # 零售價 = 通路進貨價 × 2（通路抽成 50%）
    "fallback_usd_per_kg": 25.0,  # 完全無價格資料時的保守值
}

# 包材成本（TWD／支），依 SKU 指定；未指定者用 fallback 公式
PACKAGING_TWD = {
    "RA-C01": 26.0,   # 120 mL 軟管＋翻蓋
    "RA-C02": 34.0,   # 150 mL PET 瓶＋壓頭
    "RA-H03": 42.0,   # 30 mL 玻璃滴管瓶
    "RA-H04": 48.0,   # 50 mL 玻璃霜罐
    "RA-H05": 55.0,   # 5 片／盒（不織布＋鋁袋＋外盒）
    "RA-L06": 42.0,   # 30 mL 玻璃滴管瓶
    "RA-L07": 24.0,   # 40 mL 軟管
    "RA-L08": 96.0,   # 15 mL × 3 玻璃安瓶＋外盒
    "RA-N09": 38.0,   # 15 mL 玻璃罐
    "RA-N10": 48.0,   # 50 mL 玻璃霜罐
    "RA-R11": 46.0,   # 30 mL 深色玻璃滴管瓶
    "RA-S13": 42.0,   # 300 mL PET 壓瓶
    "RA-S14": 40.0,   # 50 mL 噴霧瓶
    "RA-R12": 180.0,  # 禮盒外盒＋內襯（內容物包材另計）
}

# 充填成本（TWD／支）
FILLING_TWD = {
    "RA-H05": 25.0,   # 面膜需含浸＋封口
    "RA-L08": 30.0,   # 安瓶需熔封
    "RA-S13": 18.0,
}


def is_num(s):
    try:
        float(str(s).strip())
        return True
    except (TypeError, ValueError):
        return False


def load_json(path, default=None):
    if not os.path.exists(path):
        return default
    with open(path, encoding="utf-8") as f:
        return json.load(f)


# ---------- 配方比例解析 ----------

def resolve_pcts(rows):
    """把 pct_min/pct_max/q.s. 轉成單一百分比，總和 = 100

    規則：
      1. 有 min/max 者取中位數
      2. q.s. 的 Aqua 吸收剩餘量
      3. 其他 q.s.（pH 調整劑等）固定給 0.1%
      4. 若無 Aqua 且總和 < 100，按比例放大到 100
    """
    numeric, qs = [], []
    for r in rows:
        mn, mx = (r.get("pct_min") or "").strip(), (r.get("pct_max") or "").strip()
        if is_num(mn) and is_num(mx):
            numeric.append((r, (float(mn) + float(mx)) / 2.0))
        else:
            qs.append(r)

    total = sum(v for _, v in numeric)
    out = list(numeric)

    minor = [r for r in qs if (r.get("inci_name") or "").strip() != "Aqua"]
    for r in minor:
        out.append((r, 0.1))
        total += 0.1

    aqua = [r for r in qs if (r.get("inci_name") or "").strip() == "Aqua"]
    if aqua:
        out.append((aqua[0], max(0.0, 100.0 - total)))
    elif total > 0 and abs(total - 100.0) > 0.01:
        scale = 100.0 / total
        out = [(r, v * scale) for r, v in out]
    return out


# ---------- 價格解析（三層） ----------

def resolve_price(inci, master, prices, overrides):
    """回傳 (usd_per_kg, layer, note)"""
    ref = (master.get(inci) or {}).get("ref_price_usd_per_kg")

    ov = overrides.get(inci)
    if isinstance(ov, dict):
        ov = ov.get("usd_per_kg")
    if ov:
        return float(ov), "L3_manual", "代工廠／人工報價"

    p = prices.get(inci) or {}
    if p.get("status") == "ok" and p.get("bulk_usd_per_kg"):
        v = float(p["bulk_usd_per_kg"])
        if ref and (v > ref * 6 or v < ref * 0.15):
            return float(ref), "L2_ref", f"L1 抓取值 {v:.2f} 偏離參考價 {ref:.2f} 過大，已改用參考價"
        return v, "L1_scraped", f"makingcosmetics：{p.get('matched_name')}"

    if ref:
        return float(ref), "L2_ref", "業界參考價"
    return CFG["fallback_usd_per_kg"], "fallback", "無任何價格資料，採保守值"


# ---------- 主計算 ----------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--usd-twd", type=float, default=CFG["usd_twd"], help="USD→TWD 匯率")
    args = ap.parse_args()
    CFG["usd_twd"] = args.usd_twd

    products = load_json(P_PRODUCTS)["products"]
    master = load_json(P_MASTER)["ingredients"]
    prices = (load_json(P_PRICES, {}) or {}).get("prices", {})
    overrides = load_json(P_OVERRIDE, {}) or {}
    overrides = overrides.get("overrides", overrides)

    with open(P_INGRED, encoding="utf-8-sig") as f:
        all_rows = list(csv.DictReader(f))

    by_sku = {}
    for r in all_rows:
        by_sku.setdefault(r["sku_code"], []).append(r)

    detail_rows = []
    summary = []

    for prod in products:
        sku = prod["sku_code"]
        rows = by_sku.get(sku)

        # 禮盒：成本 = 內容物成本合計 + 禮盒包材
        if not rows:
            contents = prod.get("contents") or []
            if not contents:
                continue
            sub = [s for s in summary if s["sku_code"] in contents]
            if len(sub) != len(contents):
                continue
            raw_twd = sum(s["raw_material_twd"] for s in sub)
            pack_twd = PACKAGING_TWD.get(sku, 0.0)
            fill_twd = sum(s["filling_twd"] for s in sub)
            qc_twd = CFG["qc_per_unit_twd"]
            base = raw_twd + pack_twd + fill_twd + qc_twd
            exf = base * (1 + CFG["factory_margin_pct"])
            total = exf * (1 + CFG["logistics_pct"])
            retail = prod.get("suggested_retail_twd") or 0
            summary.append({
                "sku_code": sku,
                "name_zh": prod["name_zh"],
                "form": prod.get("form", ""),
                "size_ml": "",
                "priority": prod.get("priority", ""),
                "fill_grams": "",
                "formula_density": "",
                "raw_material_twd": round(raw_twd, 2),
                "packaging_twd": round(pack_twd, 2),
                "filling_twd": round(fill_twd, 2),
                "qc_twd": round(qc_twd, 2),
                "exfactory_twd": round(exf, 2),
                "total_cost_twd": round(total, 2),
                "suggested_retail_twd": retail,
                "channel_price_twd": round(retail / CFG["channel_multiplier"], 2) if retail else "",
                "gross_margin_pct": round((retail - total) / retail * 100, 1) if retail else "",
                "raw_cost_ratio_pct": round(raw_twd / total * 100, 1) if total else "",
                "l1_count": 0, "l2_count": 0, "l3_count": 0, "fallback_count": 0,
                "note": "禮盒組合（內容物成本合計）",
            })
            continue

        resolved = resolve_pcts(rows)

        # 配方密度：ρ = 1 / Σ(w_i / ρ_i)
        inv = 0.0
        for r, pct in resolved:
            d = (master.get(r["inci_name"]) or {}).get("density") or 1.0
            inv += (pct / 100.0) / d
        density = 1.0 / inv if inv > 0 else 1.0

        size_ml = prod.get("size_ml") or 0
        fill_grams = size_ml * density

        raw_usd_per_kg = 0.0
        layers = {"L1_scraped": 0, "L2_ref": 0, "L3_manual": 0, "fallback": 0}
        for r, pct in resolved:
            inci = r["inci_name"]
            usd_kg, layer, note = resolve_price(inci, master, prices, overrides)
            layers[layer] = layers.get(layer, 0) + 1
            contrib = pct / 100.0 * usd_kg
            raw_usd_per_kg += contrib
            detail_rows.append({
                "sku_code": sku,
                "inci_name": inci,
                "function_category": r.get("function_category", ""),
                "pct": round(pct, 3),
                "usd_per_kg": round(usd_kg, 2),
                "price_layer": layer,
                "cost_usd_per_kg_bulk": round(contrib, 4),
                "cost_twd_per_kg_bulk": round(contrib * CFG["usd_twd"], 2),
                "note": note,
            })

        raw_twd = raw_usd_per_kg * CFG["usd_twd"] * (fill_grams / 1000.0) / (1 - CFG["loss_rate"])
        pack_twd = PACKAGING_TWD.get(sku, 12.0 + 0.12 * size_ml)
        fill_twd = FILLING_TWD.get(sku, max(8.0, 6.0 + 0.06 * size_ml))
        qc_twd = CFG["qc_per_unit_twd"]

        base = raw_twd + pack_twd + fill_twd + qc_twd
        exf = base * (1 + CFG["factory_margin_pct"])
        total = exf * (1 + CFG["logistics_pct"])
        retail = prod.get("suggested_retail_twd") or 0

        summary.append({
            "sku_code": sku,
            "name_zh": prod["name_zh"],
            "form": prod.get("form", ""),
            "size_ml": size_ml,
            "priority": prod.get("priority", ""),
            "fill_grams": round(fill_grams, 1),
            "formula_density": round(density, 3),
            "raw_material_twd": round(raw_twd, 2),
            "packaging_twd": round(pack_twd, 2),
            "filling_twd": round(fill_twd, 2),
            "qc_twd": round(qc_twd, 2),
            "exfactory_twd": round(exf, 2),
            "total_cost_twd": round(total, 2),
            "suggested_retail_twd": retail,
            "channel_price_twd": round(retail / CFG["channel_multiplier"], 2) if retail else "",
            "gross_margin_pct": round((retail - total) / retail * 100, 1) if retail else "",
            "raw_cost_ratio_pct": round(raw_twd / total * 100, 1) if total else "",
            "l1_count": layers.get("L1_scraped", 0),
            "l2_count": layers.get("L2_ref", 0),
            "l3_count": layers.get("L3_manual", 0),
            "fallback_count": layers.get("fallback", 0),
            "note": "",
        })

    # ---------- 輸出 CSV ----------
    if detail_rows:
        with open(OUT_DETAIL, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(detail_rows[0].keys()))
            w.writeheader()
            w.writerows(detail_rows)

    if summary:
        with open(OUT_REPORT, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(summary[0].keys()))
            w.writeheader()
            w.writerows(summary)

    # ---------- 敏感度分析 ----------
    sens = []
    for s in summary:
        if not s["raw_material_twd"]:
            continue
        raw = s["raw_material_twd"]
        row = {"sku_code": s["sku_code"], "name_zh": s["name_zh"], "base": s["total_cost_twd"]}
        for d in (-0.30, -0.15, 0.15, 0.30):
            new_raw = raw * (1 + d)
            base = new_raw + s["packaging_twd"] + s["filling_twd"] + s["qc_twd"]
            total = base * (1 + CFG["factory_margin_pct"]) * (1 + CFG["logistics_pct"])
            row[f"{int(d*100):+d}%"] = round(total, 2)
        sens.append(row)

    # ---------- 輸出 Markdown ----------
    now = datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds")
    l1 = sum(s["l1_count"] for s in summary)
    l2 = sum(s["l2_count"] for s in summary)
    l3 = sum(s["l3_count"] for s in summary)
    fb = sum(s["fallback_count"] for s in summary)
    tot = l1 + l2 + l3 + fb or 1

    lines = []
    lines.append("# 頂級化妝品線 — 原料成本估算報告\n")
    lines.append(f"- 產出時間：{now}")
    lines.append(f"- 匯率假設：USD 1 = TWD {CFG['usd_twd']}")
    lines.append(f"- 製程損耗：{CFG['loss_rate']*100:.0f}%")
    lines.append(f"- 代工廠管銷＋毛利：{CFG['factory_margin_pct']*100:.0f}%")
    lines.append(f"- 品牌端物流／管銷：{CFG['logistics_pct']*100:.0f}%")
    lines.append(f"- 通路倍率：零售價 = 通路進貨價 × {CFG['channel_multiplier']:.0f}\n")
    lines.append("> **本報告為估算值，非報價。** 實際成本以代工廠正式 RFQ 回覆為準。\n")

    lines.append("## 一、價格來源覆蓋率\n")
    lines.append("| 層級 | 說明 | 命中筆數 | 佔比 |")
    lines.append("|---|---|---:|---:|")
    lines.append(f"| L1 自動抓取 | makingcosmetics.com 公開價 | {l1} | {l1/tot*100:.1f}% |")
    lines.append(f"| L2 內建參考價 | 業界行情（含偏離修正） | {l2} | {l2/tot*100:.1f}% |")
    lines.append(f"| L3 人工覆寫 | 代工廠報價 | {l3} | {l3/tot*100:.1f}% |")
    lines.append(f"| fallback | 無資料，保守值 | {fb} | {fb/tot*100:.1f}% |")
    lines.append(f"| **合計** | | **{tot}** | 100% |\n")

    lines.append("## 二、逐支成本彙總（TWD／支）\n")
    lines.append("| SKU | 品名 | 容量 | 料體成本 | 包材 | 充填 | 檢驗 | 出廠成本 | 總成本 | 建議售價 | 毛利率 |")
    lines.append("|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|")
    for s in summary:
        lines.append(
            f"| {s['sku_code']} | {s['name_zh']} | {s['size_ml']} | "
            f"{s['raw_material_twd']:,.0f} | {s['packaging_twd']:,.0f} | {s['filling_twd']:,.0f} | "
            f"{s['qc_twd']:,.0f} | {s['exfactory_twd']:,.0f} | {s['total_cost_twd']:,.0f} | "
            f"{s['suggested_retail_twd']:,} | {s['gross_margin_pct']}% |"
        )
    lines.append("")

    lines.append("## 三、敏感度分析（料體價格變動對總成本的影響）\n")
    if sens:
        lines.append("| SKU | 基準總成本 | -30% | -15% | +15% | +30% |")
        lines.append("|---|---:|---:|---:|---:|---:|")
        for r in sens:
            lines.append(
                f"| {r['sku_code']} | {r['base']:,.0f} | {r['-30%']:,.0f} | {r['-15%']:,.0f} | "
                f"{r['+15%']:,.0f} | {r['+30%']:,.0f} |"
            )
        lines.append("")

    lines.append("## 四、成本結構佔比\n")
    lines.append("| SKU | 料體佔比 | 包材佔比 | 其他 |")
    lines.append("|---|---:|---:|---:|")
    for s in summary:
        if not s["total_cost_twd"]:
            continue
        t = s["total_cost_twd"]
        lines.append(
            f"| {s['sku_code']} | {s['raw_material_twd']/t*100:.1f}% | "
            f"{s['packaging_twd']/t*100:.1f}% | {(t-s['raw_material_twd']-s['packaging_twd'])/t*100:.1f}% |"
        )
    lines.append("")

    lines.append("## 五、定價合理性檢查\n")
    lines.append("判斷基準：零售價 ÷ 出廠成本。精品保養品常見區間為 **6～12 倍**（含品牌行銷、通路抽成、退換貨）。\n")
    lines.append("| SKU | 出廠成本 | 建議售價 | 倍率 | 建議零售價區間（6～8 倍） | 判定 |")
    lines.append("|---|---:|---:|---:|---|---|")
    for s in summary:
        if not s["exfactory_twd"] or not s["suggested_retail_twd"]:
            continue
        mk = s["suggested_retail_twd"] / s["exfactory_twd"]
        lo, hi = s["exfactory_twd"] * 6, s["exfactory_twd"] * 8
        if mk > 12:
            verdict = "偏高（通路推力不足風險）"
        elif mk < 4:
            verdict = "偏低（品牌毛利不足）"
        else:
            verdict = "合理"
        lines.append(
            f"| {s['sku_code']} | {s['exfactory_twd']:,.0f} | {s['suggested_retail_twd']:,} | "
            f"{mk:.1f}× | {lo:,.0f} – {hi:,.0f} | {verdict} |"
        )
    lines.append("")
    lines.append("> 倍率偏高不代表不能賣，但代表**必須有足夠的品牌力與行銷預算支撐**，")
    lines.append("> 否則通路會要求更高折讓、或消費者轉向競品。建議 P0 四支先落在 6～8 倍區間試水溫。\n")

    lines.append("## 六、使用限制（必讀）\n")
    lines.append("1. **L1 抓取價是零售／小批量價**，通常比代工廠批量價高 2–5 倍，只用來交叉驗證，不用來定價。")
    lines.append("2. **L2 參考價是業界行情區間的中位數**，誤差可達 ±40%。")
    lines.append("3. 未計入：國際運費、進口關稅、報關、批量折扣、匯率避險、代工廠開模費、法規登錄費。")
    lines.append("4. 包材成本為概估，實際依選定瓶器、印刷方式（網印／燙金／UV）、訂購量而定。")
    lines.append("5. 檢驗費以每批分攤估算，實際依檢驗項目與批量而定。")
    lines.append("6. 毛利率未計入品牌行銷費、代言費、通路促銷折讓、退換貨損失。")
    lines.append("7. 本報告所有 SKU 仍為 `enabled=false`，未經林總核定不得對外使用。\n")

    lines.append("## 七、如何提高準確度\n")
    lines.append("1. 把代工廠正式報價填入 `data/price_overrides.json`（L3），重跑本引擎即自動採用。")
    lines.append("2. 格式：`{\"overrides\": {\"Squalane\": {\"usd_per_kg\": 28.5, \"source\": \"安成生化 2026-11 報價\"}}}`")
    lines.append("3. 包材實際報價填入 `scripts/cost_engine.py` 的 `PACKAGING_TWD`。")
    lines.append("4. 匯率用 `--usd-twd` 參數覆寫。\n")

    with open(OUT_MD, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))

    print(f"完成：{len(summary)} 支 SKU、{len(detail_rows)} 筆原料明細")
    print(f"  L1 抓取 {l1}／L2 參考 {l2}／L3 覆寫 {l3}／fallback {fb}")
    print(f"  → {OUT_DETAIL}")
    print(f"  → {OUT_REPORT}")
    print(f"  → {OUT_MD}")


if __name__ == "__main__":
    main()
