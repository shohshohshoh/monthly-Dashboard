#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
営業日報 Excel → サマリーJSON 変換スクリプト

使い方:
    python create_summary.py 2026 5
    → frontend/public/data/summary_2026_5.json を出力
"""
import sys
import json
import argparse
from pathlib import Path
from collections import defaultdict
import openpyxl

ROOT     = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT / "frontend" / "public" / "data"

def load_summary(year: int, month: int) -> dict:
    filename = f"★営業日報{year}年{month}月.xlsx"
    path     = DATA_DIR / filename
    if not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)

    # ── データシート（日次売上） ──────────────────────────
    ws   = wb["データ"]
    rows = list(ws.iter_rows(min_row=1, max_row=54, values_only=True))
    COL0 = 5

    date_r  = rows[0][COL0:COL0+31]
    hol_r   = rows[1][COL0:COL0+31]
    wday_r  = rows[3][COL0:COL0+31]
    sales_r = rows[4][COL0:COL0+31]

    daily = []
    for d, h, w, s in zip(date_r, hol_r, wday_r, sales_r):
        if d is not None:
            daily.append({
                "day":       d.day,
                "weekday":   w or "",
                "is_closed": h == "休",
                "sales":     int(s or 0),
            })

    working_days = sum(1 for r in daily if not r["is_closed"])
    closed_days  = sum(1 for r in daily if r["is_closed"])

    # ── 商品別シート（FOOD / DRINK Top） ─────────────────
    ws_item = wb["商品別"]
    food_map, drink_map = defaultdict(int), defaultdict(int)
    for r in ws_item.iter_rows(min_row=3, max_row=95, values_only=True):
        if r[4] and r[7]:  food_map[r[4]]  += int(r[7]  or 0)
        if r[8] and r[11]: drink_map[r[8]] += int(r[11] or 0)

    top_food  = [{"name": k, "amount": v}
                 for k, v in sorted(food_map.items(),  key=lambda x: x[1], reverse=True)[:6]]
    top_drink = [{"name": k, "amount": v}
                 for k, v in sorted(drink_map.items(), key=lambda x: x[1], reverse=True)[:6]]

    # ── 集計値（データシートの合計行などから取得） ────────
    # ★ 実データExcelのセル位置に合わせて調整してください
    def cell(ws, row, col):
        v = ws.cell(row=row, column=col).value
        return int(v or 0)

    ws_d = wb["データ"]
    total_sales  = cell(ws_d, 5, 4)   # 純売上高合計（仮）
    food_sales   = cell(ws_d, 6, 4)   # FOOD売上合計（仮）
    drink_sales  = cell(ws_d, 7, 4)   # DRINK売上合計（仮）
    lunch_pax    = cell(ws_d, 8, 4)   # 昼食客数（仮）
    dinner_pax   = cell(ws_d, 9, 4)   # 夕食客数（仮）
    lunch_amt    = cell(ws_d,10, 4)   # 昼食売上（仮）
    dinner_amt   = cell(ws_d,11, 4)   # 夕食売上（仮）
    cash         = cell(ws_d,12, 4)   # 現金（仮）
    jcb          = cell(ws_d,13, 4)   # JCB（仮）
    chiba        = cell(ws_d,14, 4)   # 千葉銀行（仮）
    aqua         = cell(ws_d,15, 4)   # アクアコイン（仮）
    paypay       = cell(ws_d,16, 4)   # PayPay（仮）
    kake         = cell(ws_d,17, 4)   # 売掛（仮）
    baiten       = cell(ws_d,18, 4)   # 売店（仮）
    sonota       = cell(ws_d,19, 4)   # その他（仮）

    summary = {
        "meta": {
            "year":         year,
            "month":        month,
            "label":        f"{year}年{month}月",
            "filename":     filename,
            "working_days": working_days,
            "closed_days":  closed_days,
        },
        "kpi": {
            "total_sales":  total_sales,
            "food_sales":   food_sales,
            "drink_sales":  drink_sales,
            "baiten_sales": baiten,
            "sonota_sales": sonota,
            "lunch_pax":    lunch_pax,
            "dinner_pax":   dinner_pax,
            "lunch_amt":    lunch_amt,
            "dinner_amt":   dinner_amt,
        },
        "payment": [
            {"label": "現金",         "amount": cash},
            {"label": "JCB",          "amount": jcb},
            {"label": "千葉銀行",     "amount": chiba},
            {"label": "アクアコイン", "amount": aqua},
            {"label": "PayPay",       "amount": paypay},
            {"label": "売掛金",       "amount": kake},
        ],
        "category": [
            {"label": "FOOD",   "amount": food_sales},
            {"label": "DRINK",  "amount": drink_sales},
            {"label": "売店",   "amount": baiten},
            {"label": "その他", "amount": sonota},
        ],
        "daily":     daily,
        "top_food":  top_food,
        "top_drink": top_drink,
    }
    return summary


def main():
    parser = argparse.ArgumentParser(description="営業日報 → サマリーJSON")
    parser.add_argument("year",  type=int, help="年（例: 2026）")
    parser.add_argument("month", type=int, help="月（例: 5）")
    args = parser.parse_args()

    print(f"読み込み中: ★営業日報{args.year}年{args.month}月.xlsx")
    summary = load_summary(args.year, args.month)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = OUT_DIR / f"summary_{args.year}_{args.month}.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(summary, f, ensure_ascii=False, indent=2)

    print(f"出力完了: {out_path}")


if __name__ == "__main__":
    main()
