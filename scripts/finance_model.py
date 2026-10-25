#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
投資與獲利財務模型

產出 data/finance_model.json，供 build_web.py（網頁）與 build_docx.py（Word）共用，
確保兩份輸出的數字完全一致。

模型範圍
  1. 一次性開發投資明細
  2. 化工廠（安成生化）合作條件與投入
  3. AI Agent 團隊成本對比
  4. 兩情境營收預測（原提案價 / 建議價）
  5. 損益表、回收期
  6. 價值鏈獲利分配瀑布

用法：python scripts/finance_model.py
"""

import csv
import json
import os
from datetime import datetime, timezone

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")

P_PRODUCTS = os.path.join(DATA, "products.json")
P_COST = os.path.join(DATA, "cost_report.csv")
OUT = os.path.join(DATA, "finance_model.json")

# ============================================================
# 一、全域假設
# ============================================================
ASSUME = {
    "usd_twd": 32.0,
    "tax_rate": 0.20,              # 營利事業所得稅
    "channel_cut": 0.50,           # 通路抽成（百貨／藥粧）
    "self_direct_share": 0.70,     # 自營電商佔比
    "payment_fee": 0.05,           # 金流手續費（自營部分）
    "logistics_pct": 0.04,         # 物流倉儲（佔實收）
    "marketing_pct": 0.18,         # 行銷（佔實收）
    "agent_commission_pct": 0.12,  # 代理佣金（佔通路部分）
    "community_pct": 0.03,         # 社群貢獻回饋（佔實收）
    "misc_pct": 0.03,              # 雜支（佔實收）
    "headcount": 2,                # 精簡後人力
    "salary_per_month": 45000,     # 平均月薪
    "ai_budget_usd_month": 60,     # D-029：每月 AI 預算
}

# ============================================================
# 二、一次性開發投資（TWD）
# ============================================================
INVESTMENT = [
    ("A", "產品開發", "打樣（4 支 P0 × 3 輪）", 240000,
     "每支每輪 20,000，含配方調整與穩定性初篩"),
    ("A", "產品開發", "安定性與相容性測試", 120000,
     "3 個月加速安定性 + 包材相容性"),
    ("A", "產品開發", "微生物與重金屬檢驗（14 支）", 168000,
     "每支 12,000，含生菌數、黴菌、鉛汞砷鎘"),
    ("A", "產品開發", "功效評估（4 支，人體試驗）", 480000,
     "每支 120,000，含 28 天人體功效與安全性"),
    ("A", "產品開發", "法規登錄與 PIF 建置（14 支）", 280000,
     "每支 20,000，含產品登錄、PIF、標示審查"),

    ("B", "包材與模具", "公版瓶器打樣與測試", 60000,
     "多款公版比對、跌落測試、相容性"),
    ("B", "包材與模具", "私模開模（霜罐 + 滴管瓶 2 款）", 280000,
     "每款 140,000，含模具與首件確認"),
    ("B", "包材與模具", "印刷網版與打樣", 90000,
     "網印／燙金／UV 各版費"),

    ("C", "品牌與通路", "品牌識別與視覺系統", 250000,
     "LOGO、標準字、色彩、輔助圖形、品牌手冊"),
    ("C", "品牌與通路", "包裝設計（14 支）", 210000,
     "每支 15,000，含外盒、瓶身、標籤"),
    ("C", "品牌與通路", "官網與電商系統", 300000,
     "含商品頁、購物車、金流串接、會員系統"),
    ("C", "品牌與通路", "商標註冊（台灣 + 日本，6 類）", 180000,
     "每國每類 15,000，含檢索與申請"),

    ("D", "首批量產", "4 支 P0 × 3,000 支（含 10% 損耗）", 1244000,
     "依成本引擎出廠成本計算"),

    ("E", "行銷啟動", "KOL／部落客合作", 350000,
     "20 位中小型 KOL，含產品寄送"),
    ("E", "行銷啟動", "內容製作（攝影、影片）", 180000,
     "產品攝影、形象影片、短影音素材"),
    ("E", "行銷啟動", "數位廣告（3 個月）", 450000,
     "Meta、Google、LINE 投放"),
    ("E", "行銷啟動", "通路陳列與試用", 120000,
     "藥粧通路陳列架、試用品"),

    ("F", "AI Agent 系統", "建置（一次性）", 150000,
     "20 個 Agent 規格落地、6 個值班入口、知識庫"),
    ("F", "AI Agent 系統", "年度維運（USD 60/月）", 23040,
     "D-029 核定預算"),

    ("G", "營運週轉金", "人事（3 人 × 6 個月）", 810000,
     "籌備期人力"),
    ("G", "營運週轉金", "辦公室與雜支（6 個月）", 180000,
     "租金、設備、行政"),
]

# ============================================================
# 三、化工廠（安成生化）合作條件
# ============================================================
FACTORY = {
    "name": "台灣安成生化科技有限公司",
    "role": "委託製造（OEM／ODM）",
    "factory_cash_investment_twd": 0,
    "factory_note": "代工廠不投入資金、不佔股。所有開發費用由品牌方負擔，避免股權與配方歸屬糾纏。",
    "factory_contribution": [
        ("人力", "配方工程師、品保、生產線人力", "依打樣與量產排程投入"),
        ("設備", "乳化槽、均質機、充填線、包裝線", "既有設備，不需新增"),
        ("品保", "原料檢驗、半成品與成品檢驗、留樣", "依 GMP 規範"),
        ("文件", "GMP 證明、成分表、COA、MSDS", "隨每批提供"),
        ("法規協助", "化粧品登錄所需製造廠資料", "配合提供"),
    ],
    "brand_side_pays": [
        ("打樣費", "每支每輪 20,000", "3 輪為限，超出另計"),
        ("開模費", "每款 140,000", "可談分攤 50% 換 2 年獨家"),
        ("檢驗費", "每支 12,000", "第三方公正檢驗"),
        ("法規登錄費", "每支 20,000", "含 PIF 建置"),
    ],
    "contract_terms": [
        ("配方所有權", "歸品牌方（Rokaizumi）所有", "代工廠不得自行使用或轉售"),
        ("獨家代工", "2 年獨家，換取開模費分攤 50%", "期滿可續約或轉廠"),
        ("價格鎖定", "簽約後 12 個月不調價", "原料漲幅超過 15% 可重議"),
        ("MOQ", "首批 3,000 支／SKU，後續 1,500 支", "低於 MOQ 加收換線費"),
        ("付款條件", "30% 訂金、70% 出貨前付清", "首批可談 40% / 60%"),
        ("品質違約", "不良率 > 2% 全額重製", "含運費與延遲損失"),
        ("保密條款", "配方、客戶、銷量均為機密", "違約賠償 3 年營業額"),
        ("交期", "下單後 45 天交貨", "急單 30 天加收 15%"),
    ],
    "risk_notes": [
        "代工廠若同時代工競品，需在合約中約定競業限制",
        "配方應分階段揭露，避免一次性交付完整配方",
        "關鍵原料建議品牌方自行採購或指定供應商，避免被綁定",
        "每批留樣保存 3 年，以備法規查核與客訴追溯",
    ],
}

# ============================================================
# 四、AI Agent 團隊
# ============================================================
AI_TEAM = [
    ("E1", "法遵與品保", "A04 法遵、A05 品保", "法遵專員 0.5 人", 22500,
     "法規限量核對、宣稱白名單、PIF 文件、標示審查"),
    ("E2", "產品與配方", "A06 配方、A07 成本", "產品經理 0.5 人", 25000,
     "配方架構、打樣追蹤、成本試算、供應商比價"),
    ("E3", "內容與行銷", "A08 內容、A09 社群、A10 廣告", "內容編輯 1 人", 38000,
     "文案、素材腳本、社群排程、廣告投放建議"),
    ("E4", "客服與代理", "A11 客服、A12 代理管理", "客服專員 1 人", 32000,
     "常見問題、客訴分類、代理招募與佣金結算"),
    ("E5", "數據與財務", "A14 記帳、A15 報表、A16 對帳", "會計 0.5 人", 20000,
     "收支記帳、月報、平台對帳、稅務資料整理"),
    ("E6", "供應鏈與物流", "A17 採購、A18 庫存、A19 出貨", "採購 0.5 人", 22000,
     "原料採購、庫存預警、出貨排程、物流追蹤"),
]

# ============================================================
# 五、銷量假設
# ============================================================
SALES_Y1 = {  # 第一年：4 支 P0
    "RA-C01": 6000,
    "RA-C02": 4500,
    "RA-H03": 3500,
    "RA-H04": 2800,
}
SALES_Y2 = {  # 第二年：14 支全開
    "RA-C01": 12000, "RA-C02": 9000, "RA-H03": 7000, "RA-H04": 5600,
    "RA-H05": 4000, "RA-L06": 3500, "RA-L07": 3000, "RA-L08": 1500,
    "RA-N09": 2500, "RA-N10": 1800, "RA-R11": 2000, "RA-R12": 600,
    "RA-S13": 5000, "RA-S14": 3000,
}

# 建議定價倍率（出廠成本 × 倍率，取整到 10 元）
TARGET_MULTIPLE = 7.0


def load_cost():
    out = {}
    with open(P_COST, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            out[r["sku_code"]] = {
                "name_zh": r["name_zh"],
                "form": r["form"],
                "size_ml": r["size_ml"],
                "priority": r["priority"],
                "raw": float(r["raw_material_twd"] or 0),
                "pack": float(r["packaging_twd"] or 0),
                "fill": float(r["filling_twd"] or 0),
                "qc": float(r["qc_twd"] or 0),
                "exfactory": float(r["exfactory_twd"] or 0),
                "total": float(r["total_cost_twd"] or 0),
                "retail": float(r["suggested_retail_twd"] or 0),
            }
    return out


def round10(x):
    return int(round(x / 10.0) * 10)


def pnl(sales, price_map, cost, label):
    """計算單一情境的損益"""
    a = ASSUME
    gross_retail = 0.0
    product_cost = 0.0
    lines = []
    for sku, qty in sales.items():
        c = cost[sku]
        price = price_map[sku]
        rev = qty * price
        pc = qty * c["exfactory"]
        gross_retail += rev
        product_cost += pc
        lines.append({
            "sku_code": sku, "name_zh": c["name_zh"], "qty": qty,
            "unit_price": price, "retail_revenue": rev,
            "unit_exfactory": c["exfactory"], "product_cost": pc,
            "unit_gross": price - c["exfactory"],
            "gross_margin_pct": round((price - c["exfactory"]) / price * 100, 1) if price else 0,
        })

    self_rev = gross_retail * a["self_direct_share"]
    chan_rev = gross_retail * (1 - a["self_direct_share"])
    net_revenue = self_rev * (1 - a["payment_fee"]) + chan_rev * (1 - a["channel_cut"])

    exp = {
        "產品成本（出廠）": product_cost,
        "物流與倉儲": net_revenue * a["logistics_pct"],
        "行銷": net_revenue * a["marketing_pct"],
        "人事": a["headcount"] * a["salary_per_month"] * 12,
        "平台與金流": self_rev * a["payment_fee"],
        "代理佣金": chan_rev * a["agent_commission_pct"],
        "社群貢獻回饋": net_revenue * a["community_pct"],
        "雜支": net_revenue * a["misc_pct"],
        "AI Agent 維運": a["ai_budget_usd_month"] * 12 * a["usd_twd"],
    }
    total_exp = sum(exp.values())
    ebt = net_revenue - total_exp
    tax = max(0.0, ebt * a["tax_rate"])
    net = ebt - tax

    return {
        "label": label,
        "gross_retail_revenue": round(gross_retail, 0),
        "self_direct_revenue": round(self_rev, 0),
        "channel_revenue": round(chan_rev, 0),
        "net_revenue": round(net_revenue, 0),
        "expenses": {k: round(v, 0) for k, v in exp.items()},
        "total_expenses": round(total_exp, 0),
        "ebt": round(ebt, 0),
        "tax": round(tax, 0),
        "net_profit": round(net, 0),
        "net_margin_pct": round(net / net_revenue * 100, 1) if net_revenue else 0,
        "lines": lines,
    }


def waterfall(sku, price, cost):
    """單支 SKU 的價值鏈分配瀑布（以零售價為 100%）"""
    a = ASSUME
    c = cost[sku]
    exf = c["exfactory"]
    # 混合通路：70% 自營、30% 通路
    self_share = a["self_direct_share"]
    chan_share = 1 - self_share

    channel_take = price * chan_share * a["channel_cut"]
    brand_gross = price - channel_take

    items = []
    items.append(("通路抽成", channel_take, "百貨／藥粧 50%，僅 30% 走通路"))
    items.append(("產品成本（出廠）", exf, "含料體、包材、充填、檢驗、代工廠毛利"))
    items.append(("平台與金流", price * self_share * a["payment_fee"], "自營部分 5%"))
    items.append(("物流與倉儲", brand_gross * a["logistics_pct"], "佔品牌實收 4%"))
    items.append(("行銷", brand_gross * a["marketing_pct"], "佔品牌實收 18%"))
    items.append(("代理佣金", price * chan_share * a["agent_commission_pct"], "通路部分 12%"))
    items.append(("社群貢獻回饋", brand_gross * a["community_pct"], "佔品牌實收 3%"))
    items.append(("雜支", brand_gross * a["misc_pct"], "佔品牌實收 3%"))

    used = sum(v for _, v, _ in items)
    ebt = price - used
    tax = max(0.0, ebt * a["tax_rate"])
    net = ebt - tax
    items.append(("營所稅", tax, "20%"))
    items.append(("品牌稅後淨利", net, ""))

    return {
        "sku_code": sku,
        "name_zh": c["name_zh"],
        "retail_price": price,
        "items": [
            {"name": n, "amount": round(v, 1), "pct": round(v / price * 100, 1), "note": note}
            for n, v, note in items
        ],
        "check_sum": round(sum(v for _, v, _ in items), 1),
    }


def main():
    cost = load_cost()
    a = ASSUME

    # ---------- 投資彙總 ----------
    inv_rows = []
    for cat_code, cat, item, amount, note in INVESTMENT:
        inv_rows.append({"cat_code": cat_code, "category": cat, "item": item,
                         "amount_twd": amount, "note": note})
    total_invest = sum(r["amount_twd"] for r in inv_rows)

    by_cat = {}
    for r in inv_rows:
        by_cat.setdefault(r["category"], 0)
        by_cat[r["category"]] += r["amount_twd"]

    # ---------- 定價情境 ----------
    price_original = {k: v["retail"] for k, v in cost.items()}
    price_suggested = {}
    for k, v in cost.items():
        if k == "RA-R12":
            price_suggested[k] = round10(v["exfactory"] * 6.0)
        else:
            price_suggested[k] = round10(v["exfactory"] * TARGET_MULTIPLE)

    # ---------- 損益 ----------
    pnl_y1_orig = pnl(SALES_Y1, price_original, cost, "第一年｜原提案價")
    pnl_y1_sugg = pnl(SALES_Y1, price_suggested, cost, "第一年｜建議價（7 倍）")
    pnl_y2_orig = pnl(SALES_Y2, price_original, cost, "第二年｜原提案價")
    pnl_y2_sugg = pnl(SALES_Y2, price_suggested, cost, "第二年｜建議價（7 倍）")

    # ---------- 回收期 ----------
    def payback(net_profit):
        if net_profit <= 0:
            return None
        return round(total_invest / net_profit, 2)

    # ---------- AI 團隊 ----------
    ai_rows = []
    human_total = 0
    for code, name, agents, human_role, human_cost, duty in AI_TEAM:
        human_total += human_cost
        ai_rows.append({
            "code": code, "name": name, "agents": agents,
            "human_role": human_role, "human_cost_twd": human_cost,
            "ai_cost_twd": round(a["ai_budget_usd_month"] / len(AI_TEAM) * a["usd_twd"], 1),
            "duty": duty,
        })
    ai_month = a["ai_budget_usd_month"] * a["usd_twd"]
    ai_year = ai_month * 12
    human_year = human_total * 12
    lean_human_year = a["headcount"] * a["salary_per_month"] * 12

    ai_summary = {
        "agent_count": 20,
        "entry_count": 6,
        "full_human_headcount": 4,
        "full_human_month_twd": human_total,
        "full_human_year_twd": human_year,
        "ai_month_usd": a["ai_budget_usd_month"],
        "ai_month_twd": round(ai_month, 0),
        "ai_year_twd": round(ai_year, 0),
        "lean_headcount": a["headcount"],
        "lean_human_year_twd": lean_human_year,
        "lean_total_year_twd": round(lean_human_year + ai_year, 0),
        "saving_year_twd": round(human_year - (lean_human_year + ai_year), 0),
        "saving_pct": round((human_year - (lean_human_year + ai_year)) / human_year * 100, 1),
        "note": "AI 不取代人，是讓人從 4 人降到 2 人。留下的人負責判斷、核定與對外關係。",
    }

    # ---------- 價值鏈瀑布 ----------
    wf = [waterfall(k, price_suggested[k], cost) for k in SALES_Y1]

    # ---------- 輸出 ----------
    out = {
        "schema_version": "1.0",
        "generated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "project": "Rokaizumi 頂級化妝品線",
        "status": "估算模型，非報價、非承諾；所有 SKU 仍 enabled=false",
        "assumptions": ASSUME,
        "investment": {
            "rows": inv_rows,
            "by_category": [{"category": k, "amount_twd": v} for k, v in by_cat.items()],
            "total_twd": total_invest,
        },
        "factory": FACTORY,
        "ai_team": {"rows": ai_rows, "summary": ai_summary},
        "pricing": {
            "target_multiple": TARGET_MULTIPLE,
            "original": [{"sku_code": k, "name_zh": cost[k]["name_zh"],
                          "exfactory": cost[k]["exfactory"], "price": price_original[k],
                          "multiple": round(price_original[k] / cost[k]["exfactory"], 1)}
                         for k in cost],
            "suggested": [{"sku_code": k, "name_zh": cost[k]["name_zh"],
                           "exfactory": cost[k]["exfactory"], "price": price_suggested[k],
                           "multiple": round(price_suggested[k] / cost[k]["exfactory"], 1)}
                          for k in cost],
        },
        "pnl": {
            "y1_original": pnl_y1_orig,
            "y1_suggested": pnl_y1_sugg,
            "y2_original": pnl_y2_orig,
            "y2_suggested": pnl_y2_sugg,
        },
        "payback": {
            "y1_original_years": payback(pnl_y1_orig["net_profit"]),
            "y1_suggested_years": payback(pnl_y1_sugg["net_profit"]),
            "y2_original_years": payback(pnl_y2_orig["net_profit"]),
            "y2_suggested_years": payback(pnl_y2_sugg["net_profit"]),
        },
        "waterfall": wf,
        "cost": cost,
    }

    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)

    print(f"完成 → {OUT}")
    print(f"  一次性投資：TWD {total_invest:,}")
    print(f"  第一年淨利（原價）：TWD {pnl_y1_orig['net_profit']:,.0f}　回收 "
          f"{out['payback']['y1_original_years']} 年")
    print(f"  第一年淨利（建議價）：TWD {pnl_y1_sugg['net_profit']:,.0f}　回收 "
          f"{out['payback']['y1_suggested_years']} 年")
    print(f"  AI 團隊年省：TWD {ai_summary['saving_year_twd']:,}（{ai_summary['saving_pct']}%）")


if __name__ == "__main__":
    main()
