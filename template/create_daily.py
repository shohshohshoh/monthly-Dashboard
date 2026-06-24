#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
★営業日報yyyy年m月.xlsx → daily_yyyy_m.xlsx 変換スクリプト

使い方:
    python create_daily.py 2026 5
    → frontend/public/data/daily_2026_5.xlsx を出力

変換元（データシート横持ち形式）→ 変換先（縦持ち日次テーブル）
"""
import sys
from pathlib import Path
from datetime import datetime
import openpyxl

ROOT    = Path(__file__).parent.parent
SRC_DIR = ROOT / "data"
OUT_DIR = ROOT / "frontend" / "public" / "data"

# データシートの行番号（固定レイアウト）
R_TEIKYU    = 2   # 定休（日曜=✗）
R_KYUJITSU  = 3   # 祝日FLG
R_YOUBI     = 4   # 曜日
R_JUN       = 5   # 純売上高
R_CASH      = 7   # 現金
R_TOTAL     = 10  # 総売上高
R_JCB       = 15  # JCB金額
R_CHIBA     = 16  # 千葉銀行金額
R_AQUA_S    = 20; R_AQUA_E = 24   # アクアコイン
R_PAY_S     = 25; R_PAY_E  = 34   # PayPay
R_FURU_S    = 35; R_FURU_E = 36   # ふるさと納税
R_URIKAKE   = 37  # 売掛金計
R_FOOD      = 41  # FOOD
R_DRINK     = 42  # DRINK
R_BAITEN    = 43  # 売店
R_OTHER     = 44  # その他
R_HIRU_KAK  = 47  # 昼食客数
R_YU_KAK    = 48  # 夕食客数
R_HIRU_AMT  = 50  # 昼食売上
R_YU_AMT    = 51  # 夕食売上


def _int(val):
    try:
        return int(val) if val is not None else 0
    except (ValueError, TypeError):
        return 0


def _sum(ws, r_start, r_end, col):
    return sum(_int(ws.cell(r, col).value) for r in range(r_start, r_end + 1))


def create_daily(year: int, month: int) -> Path:
    src = SRC_DIR / f"★営業日報{year}年{month}月.xlsx"
    if not src.exists():
        raise FileNotFoundError(f"{src} が見つかりません")

    out = OUT_DIR / f"daily_{year}_{month}.xlsx"
    OUT_DIR.mkdir(parents=True, exist_ok=True)

    wb_src    = openpyxl.load_workbook(str(src), data_only=True)
    ws_data   = wb_src["データ"]
    ws_shohin = wb_src["商品別"]

    # 行1 でデータ列を特定（列F以降、datetime値）
    date_cols = []
    for col in range(6, ws_data.max_column + 1):
        val = ws_data.cell(1, col).value
        if val is None:
            break
        if isinstance(val, datetime) and val.year == year and val.month == month:
            date_cols.append((col, val))

    # ── 日次データ ──
    KEYS = ["日付", "日", "曜日", "定休", "祝日", "昼天気", "夜天気",
            "総売上高", "純売上高", "現金", "JCB", "千葉銀行",
            "アクアコイン", "PayPay", "ふるさと納税", "売掛金",
            "FOOD", "DRINK", "売店", "その他",
            "昼食客数", "夕食客数", "昼食売上", "夕食売上"]

    daily_rows = []
    for col, dt in date_cols:
        teikyu   = ws_data.cell(R_TEIKYU, col).value or None
        kyujitsu = ws_data.cell(R_KYUJITSU, col).value or None
        youbi    = ws_data.cell(R_YOUBI, col).value or ""

        total = _int(ws_data.cell(R_TOTAL, col).value)
        if total == 0:
            teikyu = "休"

        row = {
            "日付":         dt.strftime("%Y/%m/%d"),
            "日":           dt.day,
            "曜日":         youbi,
            "定休":         teikyu,
            "祝日":         kyujitsu,
            "昼天気":       None,
            "夜天気":       None,
            "総売上高":     total,
            "純売上高":     _int(ws_data.cell(R_JUN,      col).value),
            "現金":         _int(ws_data.cell(R_CASH,     col).value),
            "JCB":          _int(ws_data.cell(R_JCB,      col).value),
            "千葉銀行":     _int(ws_data.cell(R_CHIBA,    col).value),
            "アクアコイン": _sum(ws_data, R_AQUA_S, R_AQUA_E, col),
            "PayPay":       _sum(ws_data, R_PAY_S,  R_PAY_E,  col),
            "ふるさと納税": _sum(ws_data, R_FURU_S, R_FURU_E, col),
            "売掛金":       _int(ws_data.cell(R_URIKAKE,  col).value),
            "FOOD":         _int(ws_data.cell(R_FOOD,     col).value),
            "DRINK":        _int(ws_data.cell(R_DRINK,    col).value),
            "売店":         _int(ws_data.cell(R_BAITEN,   col).value),
            "その他":       _int(ws_data.cell(R_OTHER,    col).value),
            "昼食客数":     _int(ws_data.cell(R_HIRU_KAK, col).value),
            "夕食客数":     _int(ws_data.cell(R_YU_KAK,   col).value),
            "昼食売上":     _int(ws_data.cell(R_HIRU_AMT, col).value),
            "夕食売上":     _int(ws_data.cell(R_YU_AMT,   col).value),
        }
        daily_rows.append(row)

    # ── 商品別データ ──
    SHOHIN_KEYS = ["日付", "日", "曜日", "順位",
                   "F商品名", "F単価", "F数量", "F金額",
                   "D商品名", "D単価", "D数量", "D金額"]

    shohin_rows = []
    for r in range(3, ws_shohin.max_row + 1):
        vals = [ws_shohin.cell(r, c).value for c in range(1, 13)]
        if vals[0] is None:
            continue
        dt = vals[0]
        if not isinstance(dt, datetime):
            continue
        shohin_rows.append({
            "日付":    dt.strftime("%Y/%m/%d"),
            "日":      dt.day,
            "曜日":    vals[2] or "",
            "順位":    vals[3] or "",
            "F商品名": vals[4],  "F単価": int(vals[5]  or 0),
            "F数量":   int(vals[6]  or 0), "F金額": int(vals[7]  or 0),
            "D商品名": vals[8],  "D単価": int(vals[9]  or 0),
            "D数量":   int(vals[10] or 0), "D金額": int(vals[11] or 0),
        })

    # ── Excel出力 ──
    wb_out = openpyxl.Workbook()

    ws1 = wb_out.active
    ws1.title = f"{year}年{month}月_日次データ"
    ws1.cell(1, 1, f"{year}年{month}月 日次データ")
    for ci, h in enumerate(KEYS, 1):
        ws1.cell(2, ci, h)
    for ri, row in enumerate(daily_rows, 3):
        for ci, key in enumerate(KEYS, 1):
            ws1.cell(ri, ci, row[key])

    ws2 = wb_out.create_sheet("商品別")
    ws2.cell(1, 5, "FOOD")
    ws2.cell(1, 9, "DRINK")
    for ci, h in enumerate(SHOHIN_KEYS, 1):
        ws2.cell(2, ci, h)
    for ri, row in enumerate(shohin_rows, 3):
        for ci, key in enumerate(SHOHIN_KEYS, 1):
            ws2.cell(ri, ci, row[key])

    wb_out.save(str(out))
    print(f"保存: {out}  ({len(daily_rows)}日, 商品別{len(shohin_rows)}行)")
    return out


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("使い方: python create_daily.py <year> <month>")
        sys.exit(1)
    create_daily(int(sys.argv[1]), int(sys.argv[2]))
