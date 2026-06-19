#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
営業日報 Excel → 日次データExcel 変換スクリプト

使い方:
    python create_daily.py 2026 5
    → frontend/public/data/daily_2026_5.xlsx を出力
"""
import argparse
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT     = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT / "frontend" / "public" / "data"

HDR_FILL = PatternFill("solid", fgColor="08123A")
HDR_FONT = Font(bold=True, color="00B4FF", size=10)
VAL_FONT = Font(color="D0E8FF", size=10)
TOT_FONT = Font(bold=True, color="00F5A0", size=10)
ROW_FILL = PatternFill("solid", fgColor="030718")
ALT_FILL = PatternFill("solid", fgColor="04091F")
THIN     = Side(style="thin", color="0F2D4A")
BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER   = Alignment(horizontal="center", vertical="center")
RIGHT    = Alignment(horizontal="right",  vertical="center")
LEFT     = Alignment(horizontal="left",   vertical="center")

COLUMNS = [
    ("日付",        10, CENTER, None),
    ("日",           5, CENTER, '0'),
    ("曜日",         6, CENTER, '@'),
    ("定休",         6, CENTER, '@'),
    ("祝日",         6, CENTER, '@'),
    ("総売上高",    14, RIGHT,  '#,##0'),
    ("純売上高",    14, RIGHT,  '#,##0'),
    ("現金",        14, RIGHT,  '#,##0'),
    ("JCB",         14, RIGHT,  '#,##0'),
    ("千葉銀行",    14, RIGHT,  '#,##0'),
    ("アクアコイン",14, RIGHT,  '#,##0'),
    ("PayPay",      12, RIGHT,  '#,##0'),
    ("売掛金",      12, RIGHT,  '#,##0'),
    ("FOOD",        14, RIGHT,  '#,##0'),
    ("DRINK",       14, RIGHT,  '#,##0'),
    ("売店",        12, RIGHT,  '#,##0'),
    ("その他",      12, RIGHT,  '#,##0'),
    ("昼食客数",    10, RIGHT,  '#,##0'),
    ("夕食客数",    10, RIGHT,  '#,##0'),
    ("昼食売上",    14, RIGHT,  '#,##0'),
    ("夕食売上",    14, RIGHT,  '#,##0'),
]
SUM_KEYS = {"総売上高","純売上高","現金","JCB","千葉銀行","アクアコイン",
            "PayPay","売掛金","FOOD","DRINK","売店","その他",
            "昼食客数","夕食客数","昼食売上","夕食売上"}


def load_daily(year: int, month: int) -> list[dict]:
    path = DATA_DIR / f"★営業日報{year}年{month}月.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["データ"]

    COL0, COLS = 6, 31

    def row_vals(r):
        return [ws.cell(row=r, column=COL0 + i).value for i in range(COLS)]

    dates    = row_vals(1)
    hol_flg  = row_vals(2)   # 定休FLG（"休"）
    holi_flg = row_vals(3)   # 祝日FLG（"祝"）
    wdays    = row_vals(4)   # 曜日
    net_s    = row_vals(5)   # 純売上高
    cash_r   = row_vals(7)   # 現金
    total_s  = row_vals(10)  # 総売上高（税込）
    jcb_r    = row_vals(15)  # JCB
    chiba_r  = row_vals(16)  # 千葉銀行
    aqua_r   = row_vals(20)  # アクアコイン
    paypay_r = row_vals(25)  # PayPay
    kake_r   = row_vals(37)  # 売掛金
    food_r   = row_vals(41)  # FOOD
    drink_r  = row_vals(42)  # DRINK
    baiten_r = row_vals(43)  # 売店
    sonota_r = row_vals(44)  # その他
    lpax_r   = row_vals(47)  # 昼食客数
    dpax_r   = row_vals(48)  # 夕食客数
    lamt_r   = row_vals(50)  # 昼食売上
    damt_r   = row_vals(51)  # 夕食売上

    rows = []
    for i, d in enumerate(dates):
        if d is None:
            continue
        day = d.day if hasattr(d, "day") else int(d)
        rows.append({
            "日付":         f"{year}/{month:02d}/{day:02d}",
            "日":           day,
            "曜日":         wdays[i] or "",
            "定休":         "休" if hol_flg[i] == "休" else "",
            "祝日":         holi_flg[i] or "",
            "総売上高":     int(total_s[i]  or 0),
            "純売上高":     int(net_s[i]    or 0),
            "現金":         int(cash_r[i]   or 0),
            "JCB":          int(jcb_r[i]    or 0),
            "千葉銀行":     int(chiba_r[i]  or 0),
            "アクアコイン": int(aqua_r[i]   or 0),
            "PayPay":       int(paypay_r[i] or 0),
            "売掛金":       int(kake_r[i]   or 0),
            "FOOD":         int(food_r[i]   or 0),
            "DRINK":        int(drink_r[i]  or 0),
            "売店":         int(baiten_r[i] or 0),
            "その他":       int(sonota_r[i] or 0),
            "昼食客数":     int(lpax_r[i]   or 0),
            "夕食客数":     int(dpax_r[i]   or 0),
            "昼食売上":     int(lamt_r[i]   or 0),
            "夕食売上":     int(damt_r[i]   or 0),
        })
    return rows


def write_excel(year: int, month: int, rows: list[dict], out_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月_日次データ"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"   # ヘッダー固定

    n_cols = len(COLUMNS)

    # タイトル行
    ws.merge_cells(f"A1:{get_column_letter(n_cols)}1")
    t = ws["A1"]
    t.value     = f"{year}年{month}月 日次データ"
    t.font      = Font(bold=True, color="00B4FF", size=13)
    t.fill      = HDR_FILL
    t.alignment = CENTER
    t.border    = BORDER
    ws.row_dimensions[1].height = 28

    # ヘッダー行
    for ci, (lbl, width, _, _) in enumerate(COLUMNS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = width
        c = ws.cell(row=2, column=ci, value=lbl)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[2].height = 20

    # データ行
    for ri, row in enumerate(rows, start=3):
        fill = ROW_FILL if ri % 2 == 0 else ALT_FILL
        for ci, (key, _, align, fmt) in enumerate(COLUMNS, 1):
            val = row[key]
            c   = ws.cell(row=ri, column=ci, value=val)
            c.font      = VAL_FONT
            c.fill      = fill
            c.border    = BORDER
            c.alignment = align
            if fmt:
                c.number_format = fmt
        ws.row_dimensions[ri].height = 18

    # 合計行
    last = 3 + len(rows)
    ws.merge_cells(f"A{last}:E{last}")
    tc = ws.cell(row=last, column=1, value="合 計")
    tc.font = HDR_FONT; tc.fill = HDR_FILL
    tc.alignment = CENTER; tc.border = BORDER

    for ci, (key, _, _, fmt) in enumerate(COLUMNS, 1):
        if key in SUM_KEYS:
            total = sum(r[key] for r in rows)
            c = ws.cell(row=last, column=ci, value=total)
            c.font          = TOT_FONT
            c.fill          = HDR_FILL
            c.border        = BORDER
            c.number_format = fmt or '#,##0'
            c.alignment     = RIGHT
    ws.row_dimensions[last].height = 20

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def main():
    parser = argparse.ArgumentParser(description="営業日報 → 日次データExcel")
    parser.add_argument("year",  type=int, help="年（例: 2026）")
    parser.add_argument("month", type=int, help="月（例: 5）")
    args = parser.parse_args()

    print(f"読み込み中: ★営業日報{args.year}年{args.month}月.xlsx")
    rows = load_daily(args.year, args.month)
    print(f"  {len(rows)}日分のデータを取得")

    out_path = OUT_DIR / f"daily_{args.year}_{args.month}.xlsx"
    write_excel(args.year, args.month, rows, out_path)
    print(f"出力完了: {out_path}")


if __name__ == "__main__":
    main()
