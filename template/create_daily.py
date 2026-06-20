#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
営業日報 Excel → 日次データExcel 変換スクリプト

使い方:
    python create_daily.py 2026 5
    → frontend/public/data/daily_2026_5.xlsx を出力

行構造（データシート）:
  行20-24 : アクアコイン（1件1行 → 日毎に合算）
  行25-34 : ペイペイ    （1件1行 → 日毎に合算）
  行35-36 : ふるさと納税（1件1行 → 日毎に合算）
  行37    : 売掛金 計   （既に日次合計）
"""
import argparse
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT     = Path(__file__).parent.parent
DATA_DIR = ROOT / "data"
OUT_DIR  = ROOT / "frontend" / "public" / "data"

# 木更津の天気データ（tenki.jp より取得）
# 形式: {year: {month: {day: (昼天気, 夜天気)}}}  晴/曇/雨 の3種類
WEATHER_DB = {
    2026: {
        5: {
             1: ("雨", "晴"),   # 雨のち晴
             2: ("晴", "曇"),   # 晴のち曇
             3: ("曇", "雨"),   # 曇のち雨
             4: ("曇", "曇"),   # 曇時々雨
             5: ("曇", "晴"),   # 曇のち晴
             6: ("晴", "晴"),   # 晴
             7: ("曇", "曇"),   # 曇一時雨
             8: ("曇", "曇"),   # 曇
             9: ("曇", "晴"),   # 曇のち晴
            10: ("晴", "晴"),   # 晴
            11: ("晴", "晴"),   # 晴
            12: ("晴", "晴"),   # 晴
            13: ("晴", "曇"),   # 晴のち曇
            14: ("曇", "曇"),   # 曇
            15: ("曇", "晴"),   # 曇のち晴
            16: ("晴", "晴"),   # 晴
            17: ("晴", "晴"),   # 晴
            18: ("晴", "晴"),   # 晴
            19: ("晴", "曇"),   # 晴のち曇
            20: ("曇", "雨"),   # 曇のち雨
            21: ("雨", "雨"),   # 雨
            22: ("雨", "曇"),   # 雨のち曇
            23: ("曇", "曇"),   # 曇
            24: ("曇", "曇"),   # 曇一時雨
            25: ("曇", "曇"),   # 曇時々晴
            26: ("曇", "曇"),   # 曇
            27: ("曇", "曇"),   # 曇
            28: ("曇", "曇"),   # 曇
            29: ("晴", "晴"),   # 晴時々曇
            30: ("晴", "曇"),   # 晴のち曇
            31: ("曇", "晴"),   # 曇のち晴
        }
    }
}

def get_weather(year: int, month: int, day: int) -> tuple[str, str]:
    """昼天気・夜天気を返す（晴/曇/雨）。未登録の場合は空文字"""
    try:
        return WEATHER_DB[year][month][day]
    except KeyError:
        return ("", "")


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

# (列キー, 列幅, 配置, 書式)
COLUMNS = [
    ("日付",         10, CENTER, None),
    ("日",            5, CENTER, '0'),
    ("曜日",          6, CENTER, '@'),
    ("定休",          6, CENTER, '@'),
    ("祝日",          6, CENTER, '@'),
    ("昼天気",         6, CENTER, '@'),
    ("夜天気",         6, CENTER, '@'),
    ("総売上高",     14, RIGHT,  '#,##0'),
    ("純売上高",     14, RIGHT,  '#,##0'),
    ("現金",         14, RIGHT,  '#,##0'),
    ("JCB",          14, RIGHT,  '#,##0'),
    ("千葉銀行",     14, RIGHT,  '#,##0'),
    ("アクアコイン", 14, RIGHT,  '#,##0'),
    ("PayPay",       12, RIGHT,  '#,##0'),
    ("ふるさと納税", 14, RIGHT,  '#,##0'),
    ("売掛金",       12, RIGHT,  '#,##0'),
    ("FOOD",         14, RIGHT,  '#,##0'),
    ("DRINK",        14, RIGHT,  '#,##0'),
    ("売店",         12, RIGHT,  '#,##0'),
    ("その他",       12, RIGHT,  '#,##0'),
    ("昼食客数",     10, RIGHT,  '#,##0'),
    ("夕食客数",     10, RIGHT,  '#,##0'),
    ("昼食売上",     14, RIGHT,  '#,##0'),
    ("夕食売上",     14, RIGHT,  '#,##0'),
]
SUM_KEYS = {
    "総売上高", "純売上高", "現金", "JCB", "千葉銀行",
    "アクアコイン", "PayPay", "ふるさと納税", "売掛金",
    "FOOD", "DRINK", "売店", "その他",
    "昼食客数", "夕食客数", "昼食売上", "夕食売上",
}


def load_shohin(year: int, month: int) -> list[dict]:
    """商品別シートから日次ランキングデータを読み込む"""
    path = DATA_DIR / f"★営業日報{year}年{month}月.xlsx"
    wb   = openpyxl.load_workbook(path, data_only=True)
    ws   = wb["商品別"]

    rows = []
    for r in range(3, ws.max_row + 1):
        d = ws.cell(row=r, column=1).value
        if d is None:
            continue
        day = d.day if hasattr(d, "day") else int(d)
        rows.append({
            "日付":   f"{year}/{month:02d}/{day:02d}",
            "日":     day,
            "曜日":   ws.cell(row=r, column=3).value or "",
            "順位":   ws.cell(row=r, column=4).value or "",
            "F商品名": ws.cell(row=r, column=5).value or "",
            "F単価":   int(ws.cell(row=r, column=6).value or 0),
            "F数量":   int(ws.cell(row=r, column=7).value or 0),
            "F金額":   int(ws.cell(row=r, column=8).value or 0),
            "D商品名": ws.cell(row=r, column=9).value or "",
            "D単価":   int(ws.cell(row=r, column=10).value or 0),
            "D数量":   int(ws.cell(row=r, column=11).value or 0),
            "D金額":   int(ws.cell(row=r, column=12).value or 0),
        })
    return rows


def load_daily(year: int, month: int) -> list[dict]:
    path = DATA_DIR / f"★営業日報{year}年{month}月.xlsx"
    if not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {path}")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["データ"]

    COL0, COLS = 6, 31

    def row_vals(r: int) -> list:
        return [ws.cell(row=r, column=COL0 + i).value for i in range(COLS)]

    def sum_rows(r_start: int, r_end: int) -> list[int]:
        """複数行を日付列ごとに合算（1件1行の個別入力を集計）"""
        totals = [0] * COLS
        for r in range(r_start, r_end + 1):
            for i, v in enumerate(row_vals(r)):
                totals[i] += int(v or 0)
        return totals

    dates    = row_vals(1)
    hol_flg  = row_vals(2)    # 定休FLG（"休"）
    holi_flg = row_vals(3)    # 祝日FLG（"祝"）
    wdays    = row_vals(4)    # 曜日
    net_s    = row_vals(5)    # 純売上高
    cash_r   = row_vals(7)    # 現金
    total_s  = row_vals(10)   # 総売上高（税込）
    jcb_r    = row_vals(15)   # JCB
    chiba_r  = row_vals(16)   # 千葉銀行
    aqua_r   = sum_rows(20, 24)  # アクアコイン（行20-24 合算）
    paypay_r = sum_rows(25, 34)  # PayPay    （行25-34 合算）
    furusato = sum_rows(35, 36)  # ふるさと納税（行35-36 合算）
    kake_r   = row_vals(37)   # 売掛金 計（既に日次合計）
    food_r   = row_vals(41)   # FOOD
    drink_r  = row_vals(42)   # DRINK
    baiten_r = row_vals(43)   # 売店
    sonota_r = row_vals(44)   # その他
    lpax_r   = row_vals(47)   # 昼食客数
    dpax_r   = row_vals(48)   # 夕食客数
    lamt_r   = row_vals(50)   # 昼食売上
    damt_r   = row_vals(51)   # 夕食売上

    rows = []
    for i, d in enumerate(dates):
        if d is None:
            continue
        day = d.day if hasattr(d, "day") else int(d)
        hiru_t, yoru_t = get_weather(year, month, day)
        rows.append({
            "日付":         f"{year}/{month:02d}/{day:02d}",
            "日":           day,
            "曜日":         wdays[i] or "",
            "定休":         "休" if hol_flg[i] == "休" else "",
            "祝日":         holi_flg[i] or "",
            "昼天気":       hiru_t,
            "夜天気":       yoru_t,
            "総売上高":     int(total_s[i]  or 0),
            "純売上高":     int(net_s[i]    or 0),
            "現金":         int(cash_r[i]   or 0),
            "JCB":          int(jcb_r[i]    or 0),
            "千葉銀行":     int(chiba_r[i]  or 0),
            "アクアコイン": aqua_r[i],
            "PayPay":       paypay_r[i],
            "ふるさと納税": furusato[i],
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


def write_excel(year: int, month: int, rows: list[dict],
                shohin: list[dict], out_path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月_日次データ"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

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
            c.font = VAL_FONT; c.fill = fill
            c.border = BORDER; c.alignment = align
            if fmt:
                c.number_format = fmt
        ws.row_dimensions[ri].height = 18

    # 合計行
    last = 3 + len(rows)
    ws.merge_cells(f"A{last}:G{last}")
    tc = ws.cell(row=last, column=1, value="合 計")
    tc.font = HDR_FONT; tc.fill = HDR_FILL
    tc.alignment = CENTER; tc.border = BORDER

    for ci, (key, _, _, fmt) in enumerate(COLUMNS, 1):
        if key in SUM_KEYS:
            total = sum(r[key] for r in rows)
            c = ws.cell(row=last, column=ci, value=total)
            c.font = TOT_FONT; c.fill = HDR_FILL
            c.border = BORDER; c.alignment = RIGHT
            c.number_format = fmt or '#,##0'
    ws.row_dimensions[last].height = 20

    add_shohin_sheet(wb, year, month, shohin)

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)


def add_shohin_sheet(wb: openpyxl.Workbook, year: int, month: int, rows: list[dict]):
    """商品別シートを追加"""
    ws = wb.create_sheet("商品別")
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    # 列定義: (キー, 幅, 配置, 書式)
    SCOLS = [
        ("日付",   10, CENTER, None),
        ("日",      5, CENTER, '0'),
        ("曜日",    6, CENTER, '@'),
        ("順位",    6, CENTER, '@'),
        ("F商品名", 22, Alignment(horizontal="left", vertical="center"), '@'),
        ("F単価",  10, RIGHT,  '#,##0'),
        ("F数量",   8, RIGHT,  '#,##0'),
        ("F金額",  12, RIGHT,  '#,##0'),
        ("D商品名", 22, Alignment(horizontal="left", vertical="center"), '@'),
        ("D単価",  10, RIGHT,  '#,##0'),
        ("D数量",   8, RIGHT,  '#,##0'),
        ("D金額",  12, RIGHT,  '#,##0'),
    ]
    n_cols = len(SCOLS)

    # カテゴリヘッダー行
    food_font  = Font(bold=True, color="FF9F1C", size=10)
    drink_font = Font(bold=True, color="00B4FF", size=10)
    ws.merge_cells("E1:H1")
    fc = ws["E1"]
    fc.value = "FOOD"; fc.font = food_font
    fc.fill = HDR_FILL; fc.alignment = CENTER; fc.border = BORDER
    ws.merge_cells("I1:L1")
    dc = ws["I1"]
    dc.value = "DRINK"; dc.font = drink_font
    dc.fill = HDR_FILL; dc.alignment = CENTER; dc.border = BORDER
    for ci in range(1, 5):
        c = ws.cell(row=1, column=ci)
        c.fill = HDR_FILL; c.border = BORDER
    ws.row_dimensions[1].height = 20

    # ヘッダー行
    for ci, (lbl, width, _, _) in enumerate(SCOLS, 1):
        ws.column_dimensions[get_column_letter(ci)].width = width
        c = ws.cell(row=2, column=ci, value=lbl)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = CENTER; c.border = BORDER
    ws.row_dimensions[2].height = 20

    # データ行（日付が変わるたびに色を交互）
    day_colors = {}
    color_idx  = 0
    for row in rows:
        if row["日"] not in day_colors:
            day_colors[row["日"]] = color_idx % 2
            color_idx += 1

    for ri, row in enumerate(rows, start=3):
        fill = ROW_FILL if day_colors[row["日"]] == 0 else ALT_FILL
        for ci, (key, _, align, fmt) in enumerate(SCOLS, 1):
            val = row[key]
            c   = ws.cell(row=ri, column=ci, value=val)
            c.font = VAL_FONT; c.fill = fill
            c.border = BORDER; c.alignment = align
            if fmt and fmt != '@':
                c.number_format = fmt
        ws.row_dimensions[ri].height = 18

    # FOOD/DRINK の合計行
    last = 3 + len(rows)
    ws.merge_cells(f"A{last}:D{last}")
    tc = ws.cell(row=last, column=1, value="合 計")
    tc.font = HDR_FONT; tc.fill = HDR_FILL
    tc.alignment = CENTER; tc.border = BORDER

    for ci, key in [(7, "F数量"), (8, "F金額"), (11, "D数量"), (12, "D金額")]:
        total = sum(r[key] for r in rows)
        c = ws.cell(row=last, column=ci, value=total)
        c.font = TOT_FONT; c.fill = HDR_FILL
        c.border = BORDER; c.alignment = RIGHT
        c.number_format = '#,##0'
    for ci in [5, 6, 9, 10]:
        c = ws.cell(row=last, column=ci)
        c.fill = HDR_FILL; c.border = BORDER
    ws.row_dimensions[last].height = 20


def main():
    parser = argparse.ArgumentParser(description="営業日報 → 日次データExcel")
    parser.add_argument("year",  type=int, help="年（例: 2026）")
    parser.add_argument("month", type=int, help="月（例: 5）")
    args = parser.parse_args()

    out_path = OUT_DIR / f"daily_{args.year}_{args.month}.xlsx"

    # 既存ファイルの上書き確認
    if out_path.exists():
        ans = input(f"\n{out_path.name} は既に存在します。上書きしますか？ [y/N]: ").strip().lower()
        if ans != "y":
            print("キャンセルしました。")
            return

    print(f"読み込み中: ★営業日報{args.year}年{args.month}月.xlsx")
    rows   = load_daily(args.year, args.month)
    shohin = load_shohin(args.year, args.month)
    print(f"  日次データ: {len(rows)}日分")
    print(f"  商品別データ: {len(shohin)}行")

    write_excel(args.year, args.month, rows, shohin, out_path)
    print(f"出力完了: {out_path}")


if __name__ == "__main__":
    main()
