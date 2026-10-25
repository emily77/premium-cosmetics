#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
原料價格自動抓取（L1 自動層）

來源：makingcosmetics.com（公開零售/批發價，USD）
方法：SFCC 端點
  1. Search-UpdateGrid  → 取得 pid 清單
  2. Product-ShowQuickView → 取得規格（Size）清單
  3. Product-Variation → 取得單一規格價格

輸出：data/ingredient_prices.json
  每筆含 source / fetched_at / url / tiers[{size, price_usd, usd_per_kg}]
  抓不到者標記 status=failed，由 cost_engine 回退到參考價。

用法：
  python scripts/fetch_prices.py                # 全部抓
  python scripts/fetch_prices.py --only Squalane
  python scripts/fetch_prices.py --refresh      # 忽略快取重抓
"""

import argparse
import json
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone
from difflib import SequenceMatcher

BASE = "https://www.makingcosmetics.com"
SITE = "/on/demandware.store/Sites-makingcosmetics-Site/en_US"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA = os.path.join(ROOT, "data")
MASTER = os.path.join(DATA, "ingredient_master.json")
OUT = os.path.join(DATA, "ingredient_prices.json")

DELAY = 1.2          # 每次請求間隔（秒），避免打爆對方
TIMEOUT = 25
MIN_MATCH = 0.55     # 名稱相似度門檻
MIN_COVERAGE = 0.5   # 關鍵詞覆蓋率門檻


def http_get(url, retries=2):
    for attempt in range(retries + 1):
        try:
            req = urllib.request.Request(url, headers={
                "User-Agent": UA,
                "Accept": "text/html,application/json,*/*",
                "Accept-Language": "en-US,en;q=0.9",
            })
            with urllib.request.urlopen(req, timeout=TIMEOUT) as r:
                return r.read().decode("utf-8", errors="replace")
        except Exception as e:
            if attempt == retries:
                return None
            time.sleep(1.5 * (attempt + 1))
    return None


# ---------- 規格解析 ----------

UNIT_TO_G = {
    "g": 1.0, "gram": 1.0, "grams": 1.0,
    "kg": 1000.0, "kgs": 1000.0, "kilogram": 1000.0,
    "oz": 28.3495, "ounce": 28.3495, "ounces": 28.3495,
    "lb": 453.592, "lbs": 453.592, "pound": 453.592, "pounds": 453.592,
    "ml": 1.0, "l": 1000.0, "liter": 1000.0, "litre": 1000.0,
    "gal": 3785.41, "gallon": 3785.41,
}


def parse_size_to_grams(text):
    """'2.2lbs / 1kg' → 1000.0（取最大等效重量）；無法解析回 None"""
    if not text:
        return None
    t = text.lower().replace(",", "")
    best = None
    for m in re.finditer(r"([\d.]+)\s*([a-z]+)", t):
        try:
            val = float(m.group(1))
        except ValueError:
            continue
        unit = m.group(2)
        if unit in UNIT_TO_G:
            grams = val * UNIT_TO_G[unit]
            if best is None or grams > best:
                best = grams
    return best


# ---------- 名稱比對 ----------

def norm(s):
    s = s.lower()
    s = re.sub(r"[^a-z0-9 ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def similarity(a, b):
    return SequenceMatcher(None, norm(a), norm(b)).ratio()


def tokens(s):
    return [t for t in norm(s).split() if len(t) >= 3]


def coverage(target, cand):
    """目標名稱的關鍵詞有多少比例出現在候選名稱中（0~1）"""
    tt = tokens(target)
    if not tt:
        return 0.0
    ct = set(tokens(cand))
    return sum(1 for t in tt if t in ct) / len(tt)


def pick_best(candidates, target):
    """candidates: [(pid, name)]；回傳 (pid, name, score) 或 None

    同時要求「字串相似度」與「關鍵詞覆蓋率」，避免
    Caprylic/Capric Triglyceride 被誤配成 Triglyceride、
    Tocopherol 被誤配成 Tocotrienols 這類假命中。
    """
    scored = []
    for pid, name in candidates:
        cov = coverage(target, name)
        score = 0.5 * similarity(name, target) + 0.5 * cov
        scored.append((score, cov, pid, name))
    if not scored:
        return None
    scored.sort(reverse=True)
    score, cov, pid, name = scored[0]
    if score < MIN_MATCH or cov < MIN_COVERAGE:
        return None
    return pid, name, score


# ---------- 抓取流程 ----------

def search_pids(query):
    url = f"{BASE}{SITE}/Search-UpdateGrid?q={urllib.parse.quote(query)}&start=0&sz=12"
    html = http_get(url)
    if not html:
        return []
    out = []
    # 每個 product 區塊：data-pid="X" ... <img alt="NAME" ...>
    for block in re.split(r'<div class="product"', html)[1:]:
        m = re.search(r'data-pid="([^"]+)"', block)
        if not m:
            continue
        pid = m.group(1)
        name = None
        for pat in (
            r'title="Quick View for ([^"]+)"',
            r'<img[^>]*\balt="([^"]+)"',
            r'class="product-name"[^>]*>\s*<a[^>]*>([^<]+)</a>',
        ):
            nm = re.search(pat, block)
            if nm:
                name = nm.group(1).strip()
                break
        out.append((pid, name or pid))
    return out


def get_sizes(pid):
    url = f"{BASE}{SITE}/Product-ShowQuickView?pid={urllib.parse.quote(pid)}"
    raw = http_get(url)
    if not raw:
        return []
    try:
        j = json.loads(raw)
    except Exception:
        return []
    sizes = []
    for attr in j.get("product", {}).get("variationAttributes", []) or []:
        if attr.get("id", "").lower() == "size":
            for v in attr.get("values", []) or []:
                if v.get("value"):
                    sizes.append(v["value"])
    return sizes


def get_price(pid, size):
    q = urllib.parse.quote(size)
    url = f"{BASE}{SITE}/Product-Variation?dwvar_{pid}_Size={q}&pid={pid}&quantity=1"
    raw = http_get(url)
    if not raw:
        return None
    try:
        j = json.loads(raw)
    except Exception:
        return None
    p = j.get("product", {}).get("price", {})
    for key in ("sales", "price"):
        node = p.get(key) or {}
        val = node.get("value")
        if val is not None:
            try:
                return float(val)
            except (TypeError, ValueError):
                pass
    return None


def fetch_one(inci, meta):
    kw = meta.get("search")
    if not kw:
        return {"status": "skipped", "reason": "no search keyword"}

    cands = search_pids(kw)
    time.sleep(DELAY)
    if not cands:
        return {"status": "failed", "reason": "no search results"}

    best = pick_best(cands, inci)
    if not best:
        return {"status": "failed", "reason": "no name match", "candidates": [c[1] for c in cands[:5]]}
    pid, name, score = best

    sizes = get_sizes(pid)
    time.sleep(DELAY)
    if not sizes:
        return {"status": "failed", "reason": "no sizes", "matched": name, "pid": pid}

    tiers = []
    for s in sizes:
        price = get_price(pid, s)
        time.sleep(DELAY)
        if price is None:
            continue
        grams = parse_size_to_grams(s)
        if not grams or grams <= 0:
            continue
        tiers.append({
            "size": s,
            "grams": round(grams, 2),
            "price_usd": round(price, 2),
            "usd_per_kg": round(price / (grams / 1000.0), 2),
        })

    if not tiers:
        return {"status": "failed", "reason": "no parsable price tiers", "matched": name, "pid": pid}

    tiers.sort(key=lambda x: x["grams"])
    bulk = tiers[-1]  # 最大包裝 = 最接近批量價

    return {
        "status": "ok",
        "source": "makingcosmetics.com",
        "matched_name": name,
        "pid": pid,
        "match_score": round(score, 3),
        "url": f"{BASE}/{pid}.html",
        "fetched_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
        "tiers": tiers,
        "bulk_usd_per_kg": bulk["usd_per_kg"],
        "bulk_size": bulk["size"],
        "retail_usd_per_kg": tiers[0]["usd_per_kg"],
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--only", help="只抓單一原料（INCI 名稱）")
    ap.add_argument("--refresh", action="store_true", help="忽略既有結果重抓")
    ap.add_argument("--limit", type=int, default=0, help="最多抓幾筆（測試用）")
    args = ap.parse_args()

    with open(MASTER, encoding="utf-8") as f:
        master = json.load(f)["ingredients"]

    existing = {}
    if os.path.exists(OUT) and not args.refresh:
        try:
            with open(OUT, encoding="utf-8") as f:
                existing = json.load(f).get("prices", {})
        except Exception:
            existing = {}

    targets = [args.only] if args.only else list(master.keys())
    if args.limit:
        targets = targets[: args.limit]

    results = dict(existing)
    ok = fail = skip = 0

    for i, inci in enumerate(targets, 1):
        if inci not in master:
            print(f"[{i}/{len(targets)}] {inci}  → 不在主檔，跳過")
            continue
        if inci in results and results[inci].get("status") == "ok" and not args.refresh:
            print(f"[{i}/{len(targets)}] {inci}  → 已有結果，跳過（--refresh 可重抓）")
            skip += 1
            continue

        print(f"[{i}/{len(targets)}] {inci} ...", end=" ", flush=True)
        r = fetch_one(inci, master[inci])
        results[inci] = r
        if r["status"] == "ok":
            print(f"OK  {r['matched_name']}  bulk={r['bulk_usd_per_kg']} USD/kg ({r['bulk_size']})")
            ok += 1
        elif r["status"] == "skipped":
            print(f"SKIP  {r.get('reason')}")
            skip += 1
        else:
            print(f"FAIL  {r.get('reason')}")
            fail += 1

        # 每筆即時落盤，避免中斷全毀
        with open(OUT, "w", encoding="utf-8") as f:
            json.dump({
                "schema_version": "1.0",
                "source_primary": "makingcosmetics.com",
                "currency": "USD",
                "note": "自動抓取之公開零售/批發價，非代工廠報價；bulk_usd_per_kg 取最大包裝換算。",
                "updated_at": datetime.now(timezone.utc).astimezone().isoformat(timespec="seconds"),
                "prices": results,
            }, f, ensure_ascii=False, indent=2)

    print()
    print(f"完成：成功 {ok}、失敗 {fail}、跳過 {skip}，共 {len(results)} 筆 → {OUT}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
