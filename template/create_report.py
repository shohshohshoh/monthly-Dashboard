#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日次データExcel → 分析レポートExcel 生成スクリプト

使い方:
    python create_report.py 2026 5
    → frontend/public/data/report_2026_5.xlsx を出力

チャート構成:
  ① 日次売上トレンド        折れ線 + 面
  ② 曜日別 平均売上         縦棒（祝日/振替/日曜 → 祝日グループ）
  ③ 支払方法別 売上構成      ドーナツ
  ④ カテゴリ別 売上構成      円グラフ + 横棒
  ⑤ 昼食 vs 夕食 比較       グループ棒（売上/客数/客単価）
  ⑥ 天気別 売上分布          箱ひげ図（晴/曇/雨）
  ⑦ FOOD/DRINK ランキング   横棒（金額・数量）
  ⑨ 客単価 日次推移          折れ線（昼食・夕食）
"""
import argparse
import io
from pathlib import Path
from collections import defaultdict
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image as XLImage
import matplotlib.font_manager as fm

ROOT    = Path(__file__).parent.parent
IN_DIR  = ROOT / "frontend" / "public" / "data"
OUT_DIR = ROOT / "frontend" / "public" / "data"

_avail = {f.name for f in fm.fontManager.ttflist}
JP = next((f for f in ["Yu Gothic", "Meiryo", "MS Gothic"] if f in _avail), "sans-serif")
plt.rcParams.update({"font.family": JP, "axes.unicode_minus": False})

C = {
    "bg":   "#03071e", "card": "#08123a",
    "c1":   "#00b4ff", "c2":   "#9b5de5",
    "c3":   "#00f5a0", "c4":   "#ff9f1c",
    "c5":   "#ff3264", "c6":   "#ffd60a",
    "c7":   "#00ecec", "grid": "#0f2d4a",
    "text": "#d0e8ff", "sub":  "#4a6a8a",
}

HDR_FILL = PatternFill("solid", fgColor="08123A")
HDR_FONT = Font(bold=True, color="00B4FF", size=11)
TTL_FONT = Font(bold=True, color="00B4FF", size=14)
VAL_FONT = Font(bold=True, color="D0E8FF", size=11)
THIN     = Side(style="thin", color="0F2D4A")
BORDER   = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
CENTER   = Alignment(horizontal="center", vertical="center")
RIGHT    = Alignment(horizontal="right",  vertical="center")
LEFT     = Alignment(horizontal="left",   vertical="center")


def dark_ax(ax, fig=None):
    if fig:
        fig.patch.set_facecolor(C["bg"])
    ax.set_facecolor(C["card"])
    ax.tick_params(colors=C["text"], labelsize=9)
    for sp in ax.spines.values():
        sp.set_edgecolor(C["grid"])
    ax.grid(color=C["grid"], ls="--", lw=0.5, alpha=0.7)


def to_img(fig, dpi=130):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    buf.seek(0)
    plt.close(fig)
    return buf


# ══════════════════════════════════════════════════════════════
# データ読み込み
# ══════════════════════════════════════════════════════════════
def load_data(year: int, month: int):
    path = IN_DIR / f"daily_{year}_{month}.xlsx"
    if not path.exists():
        raise FileNotFoundError(
            f"データファイルが見つかりません: {path}\n"
            "先に create_daily.py を実行してください。"
        )
    wb  = openpyxl.load_workbook(path)
    ws1 = wb.active

    keys = [
        "日付", "日", "曜日", "定休", "祝日",
        "昼天気", "夜天気",
        "総売上高", "純売上高",
        "現金", "JCB", "千葉銀行", "アクアコイン", "PayPay", "ふるさと納税", "売掛金",
        "FOOD", "DRINK", "売店", "その他",
        "昼食客数", "夕食客数", "昼食売上", "夕食売上",
    ]

    daily = []
    for r in range(3, ws1.max_row + 1):
        v = [ws1.cell(row=r, column=c).value for c in range(1, len(keys) + 1)]
        if v[0] is None or str(v[0]).startswith("合"):
            continue
        row = dict(zip(keys, v))
        for k in keys[7:]:          # 総売上高 以降を int に変換
            row[k] = int(row[k] or 0)
        daily.append(row)

    ws2    = wb["商品別"]
    s_keys = ["日付", "日", "曜日", "順位",
              "F商品名", "F単価", "F数量", "F金額",
              "D商品名", "D単価", "D数量", "D金額"]
    shohin = []
    for r in range(3, ws2.max_row + 1):
        v = [ws2.cell(row=r, column=c).value for c in range(1, 13)]
        if v[0] is None or str(v[0]).startswith("合"):
            continue
        row = dict(zip(s_keys, v))
        for k in ["F単価", "F数量", "F金額", "D単価", "D数量", "D金額"]:
            row[k] = int(row[k] or 0)
        shohin.append(row)

    return daily, shohin


# ══════════════════════════════════════════════════════════════
# チャート定義
# ══════════════════════════════════════════════════════════════

def chart_1(daily):
    """① 日次売上トレンド（総売上高のみ・折れ線 + 面）"""
    op    = [d for d in daily if d["定休"] != "休"]
    xs    = [d["日"] for d in op]
    total = [d["総売上高"] / 10000 for d in op]
    avg   = np.mean(total)

    fig, ax = plt.subplots(figsize=(11, 4.5))
    dark_ax(ax, fig)
    ax.fill_between(xs, total, alpha=0.15, color=C["c1"])
    ax.plot(xs, total, color=C["c1"], lw=2.2, marker="o", ms=4, label="総売上高")
    ax.axhline(avg, color=C["c4"], lw=1.4, ls=":", label=f"平均 {avg:.0f}万円")
    ax.set_xlabel("日", color=C["text"])
    ax.set_ylabel("売上（万円）", color=C["text"])
    ax.xaxis.label.set_color(C["text"]); ax.yaxis.label.set_color(C["text"])
    ax.set_title("① 日次売上トレンド（総売上高）",
                 color=C["text"], fontsize=13, fontweight="bold")
    ax.legend(facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9)
    fig.tight_layout()
    return to_img(fig)


def chart_2(daily):
    """② 曜日別平均売上（日曜・祝日・振替 → 祝日グループ）"""
    # 定休（休）を除外し、日曜 or 祝日フラグ → "祝日" に統合
    order  = ["火", "水", "木", "金", "土", "祝日"]
    colors = [C["c1"], C["c1"], C["c1"], C["c4"], C["c3"], C["c5"]]

    wmap = defaultdict(list)
    for d in daily:
        if d["定休"] == "休":
            continue
        if d["曜日"] == "日" or d["祝日"]:
            key = "祝日"
        else:
            key = d["曜日"]
        if key in order:
            wmap[key].append(d["総売上高"] / 10000)

    avgs = [np.mean(wmap[w]) if wmap[w] else 0 for w in order]
    cnts = [len(wmap[w]) for w in order]

    fig, ax = plt.subplots(figsize=(8, 4.5))
    dark_ax(ax, fig)
    bars = ax.bar(order, avgs, color=colors, edgecolor=C["bg"], width=0.6)
    for bar, v, n in zip(bars, avgs, cnts):
        if v > 0:
            ax.text(bar.get_x() + bar.get_width() / 2, v + max(avgs) * 0.02,
                    f"{v:.0f}万\nn={n}", ha="center", va="bottom",
                    fontsize=8.5, color=C["text"])
    ax.set_ylabel("平均総売上高（万円）", color=C["text"])
    ax.yaxis.label.set_color(C["text"])
    ax.set_ylim(0, max(avgs) * 1.3)

    patches = [
        mpatches.Patch(color=C["c1"], label="平日（火〜木）"),
        mpatches.Patch(color=C["c4"], label="金曜日"),
        mpatches.Patch(color=C["c3"], label="土曜日"),
        mpatches.Patch(color=C["c5"], label="祝日（日・祝・振替）"),
    ]
    ax.legend(handles=patches, facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9)
    ax.set_title("② 曜日別 平均売上（総売上高）",
                 color=C["text"], fontsize=13, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_3(daily):
    """③ 支払方法別構成（ドーナツ）"""
    lbls   = ["現金", "JCB", "千葉銀行", "アクアコイン", "PayPay", "ふるさと納税", "売掛金"]
    vals   = [sum(d[k] for d in daily) for k in lbls]
    colors = [C["c1"], C["c3"], C["c4"], C["c2"], C["c5"], C["c6"], C["c7"]]
    total  = sum(vals)

    fig, ax = plt.subplots(figsize=(8, 5.5))
    fig.patch.set_facecolor(C["bg"])
    ax.set_facecolor(C["bg"]); ax.set_aspect("equal")

    ax.pie(vals, colors=[(0, 0, 0, 0.25)] * len(vals), startangle=90,
           counterclock=False,
           wedgeprops=dict(width=0.52, edgecolor="none"),
           radius=1.05, center=(0.05, -0.05))

    _, _, ats = ax.pie(
        vals, colors=colors, startangle=90,
        autopct=lambda p: f"{p:.1f}%" if p > 2 else "",
        wedgeprops=dict(width=0.52, edgecolor=C["bg"], linewidth=2),
        pctdistance=0.78, counterclock=False
    )
    for at in ats:
        at.set_color("white"); at.set_fontsize(8.5); at.set_fontweight("bold")

    ax.add_patch(plt.Circle((0, 0), 0.38, color=C["card"], zorder=10))
    ax.text(0, 0.1, "総売上", ha="center", va="center",
            fontsize=9, color=C["sub"], zorder=11)
    ax.text(0, -0.18, f"¥{total/10000:.0f}万", ha="center", va="center",
            fontsize=12, color=C["c1"], fontweight="bold", zorder=11)

    patches = [mpatches.Patch(color=co,
                               label=f"{l}  ¥{v/10000:.0f}万 ({v/total*100:.1f}%)")
               for co, l, v in zip(colors, lbls, vals)]
    ax.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.22),
              ncol=2, fontsize=8, facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"])
    ax.set_title("③ 支払方法別 売上構成",
                 color=C["text"], fontsize=13, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_4(daily):
    """④ カテゴリ別売上構成（円グラフ + 横棒）"""
    cats   = ["FOOD", "DRINK", "売店", "その他"]
    vals   = [sum(d[k] for d in daily) for k in cats]
    colors = [C["c4"], C["c1"], C["c3"], C["c2"]]
    total  = sum(vals)

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.5))
    fig.patch.set_facecolor(C["bg"])

    ax1.set_facecolor(C["bg"]); ax1.set_aspect("equal")
    _, _, ats = ax1.pie(vals, colors=colors, startangle=90,
                        autopct=lambda p: f"{p:.1f}%",
                        wedgeprops=dict(edgecolor=C["bg"], linewidth=2),
                        counterclock=False)
    for at in ats:
        at.set_color("white"); at.set_fontsize(9); at.set_fontweight("bold")
    patches = [mpatches.Patch(color=co, label=l) for co, l in zip(colors, cats)]
    ax1.legend(handles=patches, loc="lower center", bbox_to_anchor=(0.5, -0.12),
               ncol=2, fontsize=9, facecolor=C["card"], edgecolor=C["grid"],
               labelcolor=C["text"])
    ax1.set_title("構成比", color=C["text"], fontsize=10)

    dark_ax(ax2)
    bars = ax2.barh(cats[::-1], [v / 10000 for v in vals[::-1]],
                    color=colors[::-1], edgecolor=C["bg"])
    for bar, v in zip(bars, vals[::-1]):
        ax2.text(bar.get_width() + 0.3,
                 bar.get_y() + bar.get_height() / 2,
                 f"¥{v/10000:.0f}万  ({v/total*100:.1f}%)",
                 va="center", fontsize=8.5, color=C["text"])
    ax2.set_xlabel("売上（万円）", color=C["text"])
    ax2.xaxis.label.set_color(C["text"])
    ax2.set_title("金額", color=C["text"], fontsize=10)
    ax2.set_xlim(0, max(vals) / 10000 * 1.45)

    fig.suptitle("④ カテゴリ別 売上構成",
                 color=C["text"], fontsize=13, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_5(daily):
    """⑤ 昼食 vs 夕食 比較（売上/客数/客単価）"""
    op    = [d for d in daily if d["定休"] != "休"]
    l_amt = sum(d["昼食売上"]  for d in op)
    d_amt = sum(d["夕食売上"]  for d in op)
    l_pax = sum(d["昼食客数"] for d in op)
    d_pax = sum(d["夕食客数"] for d in op)
    l_unit = l_amt / l_pax if l_pax else 0
    d_unit = d_amt / d_pax if d_pax else 0

    fig, axes = plt.subplots(1, 3, figsize=(11, 4.5))
    fig.patch.set_facecolor(C["bg"])
    labels = ["昼食", "夕食"]
    colors = [C["c4"], C["c1"]]

    for ax, (vals, ylabel, unit, fmt) in zip(axes, [
        ([l_amt / 10000, d_amt / 10000], "売上（万円）",  "万", "{:.0f}万"),
        ([l_pax,         d_pax],         "客数（名）",    "名", "{:,.0f}名"),
        ([l_unit,        d_unit],        "客単価（円）", "円", "¥{:,.0f}"),
    ]):
        dark_ax(ax)
        bars = ax.bar(labels, vals, color=colors, width=0.5, edgecolor=C["bg"])
        for bar, v in zip(bars, vals):
            ax.text(bar.get_x() + bar.get_width() / 2,
                    bar.get_height() * 1.02,
                    fmt.format(v), ha="center", fontsize=9, color=C["text"])
        ax.set_ylabel(ylabel, color=C["text"])
        ax.yaxis.label.set_color(C["text"])
        ax.set_ylim(0, max(vals) * 1.25)

    axes[0].set_title("売上金額", color=C["text"], fontsize=10)
    axes[1].set_title("来客数",   color=C["text"], fontsize=10)
    axes[2].set_title("客単価",   color=C["text"], fontsize=10)
    axes[2].yaxis.set_major_formatter(
        plt.FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    fig.suptitle("⑤ 昼食 vs 夕食 比較",
                 color=C["text"], fontsize=13, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_6(daily):
    """⑥ 天気別 売上分布（箱ひげ図 晴/曇/雨）"""
    groups = {"晴": [], "曇": [], "雨": []}
    for d in daily:
        if d["定休"] == "休":
            continue
        t = d.get("昼天気", "")
        if t in groups:
            groups[t].append(d["総売上高"] / 10000)

    lbls   = [k for k in ["晴", "曇", "雨"] if groups[k]]
    data   = [groups[k] for k in lbls]
    colors = {"晴": C["c6"], "曇": C["sub"], "雨": C["c1"]}

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4.5))
    dark_ax(ax1, fig); dark_ax(ax2)

    bp = ax1.boxplot(data, tick_labels=lbls, patch_artist=True,
                     medianprops=dict(color=C["c3"], lw=2.5),
                     whiskerprops=dict(color=C["text"], lw=1.2),
                     capprops=dict(color=C["text"], lw=1.2),
                     flierprops=dict(markerfacecolor=C["c5"], markersize=5,
                                     markeredgecolor=C["c5"]))
    for patch, k in zip(bp["boxes"], lbls):
        patch.set_facecolor(colors[k]); patch.set_alpha(0.55)
    ax1.set_ylabel("総売上高（万円）", color=C["text"])
    ax1.yaxis.label.set_color(C["text"])
    ax1.set_title("分布（箱ひげ図）", color=C["text"], fontsize=10)

    avgs = [np.mean(groups[k]) for k in lbls]
    stds = [np.std(groups[k])  for k in lbls]
    bars = ax2.bar(lbls, avgs,
                   color=[colors[k] for k in lbls],
                   width=0.5, edgecolor=C["bg"],
                   yerr=stds, capsize=6,
                   error_kw=dict(ecolor=C["text"], elinewidth=1.5))
    for bar, v, n in zip(bars, avgs, [len(groups[k]) for k in lbls]):
        ax2.text(bar.get_x() + bar.get_width() / 2,
                 bar.get_height() + max(stds) * 0.1,
                 f"{v:.0f}万\nn={n}日", ha="center", fontsize=9, color=C["text"])
    ax2.set_ylabel("平均総売上高（万円）", color=C["text"])
    ax2.yaxis.label.set_color(C["text"])
    ax2.set_ylim(0, max(avgs) * 1.45)
    ax2.set_title("平均比較（±σ）", color=C["text"], fontsize=10)

    fig.suptitle("⑥ 天気別 売上分布（昼天気）",
                 color=C["text"], fontsize=13, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_7(shohin):
    """⑦ FOOD / DRINK 売れ筋ランキング Top10（横棒 左右2段）"""
    def aggregate(key_name, key_qty, key_amt):
        m = defaultdict(lambda: {"数量": 0, "金額": 0})
        for r in shohin:
            if r[key_name]:
                m[r[key_name]]["数量"] += r[key_qty]
                m[r[key_name]]["金額"] += r[key_amt]
        top = sorted(m.items(), key=lambda x: x[1]["金額"], reverse=True)[:10]
        return (
            [x[0][:12] for x in top][::-1],
            [x[1]["金額"] / 10000 for x in top][::-1],
            [x[1]["数量"] for x in top][::-1],
        )

    f_names, f_amts, f_qtys = aggregate("F商品名", "F数量", "F金額")
    d_names, d_amts, d_qtys = aggregate("D商品名", "D数量", "D金額")

    fig, axes = plt.subplots(2, 2, figsize=(13, 9))
    fig.patch.set_facecolor(C["bg"])
    (ax_fa, ax_fq), (ax_da, ax_dq) = axes

    def _hbar(ax, names, vals, label, color_base, unit):
        dark_ax(ax)
        grad = [plt.cm.YlOrRd(0.35 + 0.65 * i / max(len(names) - 1, 1))
                for i in range(len(names))] if "YlOrRd" in color_base else \
               [plt.cm.Blues(0.35 + 0.65 * i / max(len(names) - 1, 1))
                for i in range(len(names))]
        bars = ax.barh(names, vals, color=grad, edgecolor=C["bg"])
        for bar, v in zip(bars, vals):
            ax.text(bar.get_width() + max(vals) * 0.015,
                    bar.get_y() + bar.get_height() / 2,
                    f"{v:.1f}{unit}" if unit == "万" else f"{v:,.0f}{unit}",
                    va="center", fontsize=8, color=C["text"])
        ax.set_xlabel(label, color=C["text"])
        ax.xaxis.label.set_color(C["text"])
        ax.set_xlim(0, max(vals) * 1.32)

    _hbar(ax_fa, f_names, f_amts, "売上金額（万円）", "YlOrRd", "万")
    ax_fa.set_title("FOOD  売上金額 Top10", color=C["c4"], fontsize=10, fontweight="bold")

    _hbar(ax_fq, f_names, f_qtys, "販売数量（個）", "YlOrRd", "個")
    ax_fq.set_title("FOOD  販売数量 Top10", color=C["c4"], fontsize=10, fontweight="bold")

    _hbar(ax_da, d_names, d_amts, "売上金額（万円）", "Blues", "万")
    ax_da.set_title("DRINK 売上金額 Top10", color=C["c1"], fontsize=10, fontweight="bold")

    _hbar(ax_dq, d_names, d_qtys, "販売数量（杯）", "Blues", "杯")
    ax_dq.set_title("DRINK 販売数量 Top10", color=C["c1"], fontsize=10, fontweight="bold")

    fig.suptitle("⑦ FOOD / DRINK 売れ筋ランキング Top10",
                 color=C["text"], fontsize=14, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_9(daily):
    """⑨ 客単価 日次推移（昼食・夕食 折れ線）"""
    op = [d for d in daily
          if d["定休"] != "休" and d["昼食客数"] > 0 and d["夕食客数"] > 0]
    days   = [d["日"] for d in op]
    l_unit = [d["昼食売上"] / d["昼食客数"] for d in op]
    d_unit = [d["夕食売上"] / d["夕食客数"] for d in op]

    fig, ax = plt.subplots(figsize=(11, 4.5))
    dark_ax(ax, fig)
    ax.fill_between(days, l_unit, alpha=0.1, color=C["c4"])
    ax.fill_between(days, d_unit, alpha=0.1, color=C["c1"])
    ax.plot(days, l_unit, color=C["c4"], lw=2, marker="o", ms=4, label="昼食客単価")
    ax.plot(days, d_unit, color=C["c1"], lw=2, marker="s", ms=4, label="夕食客単価")
    ax.axhline(np.mean(l_unit), color=C["c4"], lw=1, ls=":", alpha=0.6)
    ax.axhline(np.mean(d_unit), color=C["c1"], lw=1, ls=":", alpha=0.6)
    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    ax.set_xlabel("日", color=C["text"])
    ax.set_ylabel("客単価（円）", color=C["text"])
    ax.xaxis.label.set_color(C["text"]); ax.yaxis.label.set_color(C["text"])
    ax.set_title("⑨ 客単価 日次推移",
                 color=C["text"], fontsize=13, fontweight="bold")
    ax.legend(facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9)
    fig.tight_layout()
    return to_img(fig)


# ══════════════════════════════════════════════════════════════
# レポートExcel 書き出し
# ══════════════════════════════════════════════════════════════
CHART_META = [
    ("① 日次売上トレンド",        (11, 4.5), "折れ線 + 面"),
    ("② 曜日別 平均売上",         ( 8, 4.5), "縦棒グラフ"),
    ("③ 支払方法別 構成",         ( 8, 5.5), "ドーナツグラフ"),
    ("④ カテゴリ別 構成",         (10, 4.5), "円グラフ + 横棒"),
    ("⑤ 昼食 vs 夕食",           (11, 4.5), "グループ棒グラフ"),
    ("⑥ 天気別 売上分布",         (10, 4.5), "箱ひげ図 + 棒グラフ"),
    ("⑦ FOOD・DRINK ランキング",   (13, 9.0), "横棒グラフ"),
    ("⑨ 客単価 日次推移",         (11, 4.5), "折れ線グラフ"),
]


def _sheet_title(ws, title: str, ncols: int = 10):
    ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
    t = ws["A1"]
    t.value = title
    t.font = HDR_FONT; t.fill = HDR_FILL
    t.alignment = CENTER; t.border = BORDER
    ws.row_dimensions[1].height = 26


def write_index(ws, year, month, daily):
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 2

    ws.merge_cells("B1:L1")
    t = ws["B1"]
    t.value = f"{year}年{month}月 営業日報 分析レポート"
    t.font = TTL_FONT; t.fill = HDR_FILL
    t.alignment = CENTER; t.border = BORDER
    ws.row_dimensions[1].height = 36

    kpis = [
        ("総売上高（税込）", f"¥{sum(d['総売上高'] for d in daily)/10000:.0f}万円", C["c1"]),
        ("FOOD売上",         f"¥{sum(d['FOOD'] for d in daily)/10000:.0f}万円",     C["c4"]),
        ("DRINK売上",        f"¥{sum(d['DRINK'] for d in daily)/10000:.0f}万円",    C["c1"]),
        ("総来客数",         f"{sum(d['昼食客数']+d['夕食客数'] for d in daily):,}名", C["c2"]),
        ("稼働日数",         f"{sum(1 for d in daily if d['定休']!='休')}日",        C["c7"]),
    ]
    for i, (lbl, val, color) in enumerate(kpis):
        col = 2 + i * 2
        lc = get_column_letter(col); vc = get_column_letter(col + 1)
        ws.column_dimensions[lc].width = 14
        ws.column_dimensions[vc].width = 14

        cl = ws.cell(row=3, column=col, value=lbl)
        cl.font = Font(color=color[1:], size=10, bold=True)
        cl.fill = HDR_FILL; cl.alignment = CENTER; cl.border = BORDER
        ws.merge_cells(f"{lc}3:{vc}3")

        cv = ws.cell(row=4, column=col, value=val)
        cv.font = Font(color=color[1:], size=14, bold=True)
        cv.fill = PatternFill("solid", fgColor="030718")
        cv.alignment = CENTER; cv.border = BORDER
        ws.merge_cells(f"{lc}4:{vc}4")
    ws.row_dimensions[3].height = 22
    ws.row_dimensions[4].height = 30

    ws.merge_cells("B6:L6")
    s = ws["B6"]
    s.value = "■ 分析チャート一覧"
    s.font = Font(bold=True, color="00F5A0", size=11)
    s.fill = HDR_FILL; s.alignment = LEFT; s.border = BORDER

    for ci, lbl in enumerate(["#", "チャート名", "グラフ種類"], 1):
        c = ws.cell(row=7, column=ci + 1, value=lbl)
        c.font = HDR_FONT; c.fill = HDR_FILL
        c.alignment = CENTER; c.border = BORDER
    ws.column_dimensions["B"].width = 5
    ws.column_dimensions["C"].width = 30
    ws.column_dimensions["D"].width = 24

    for i, (name, _, chart_type) in enumerate(CHART_META, 1):
        r = 7 + i
        for ci, (val, font) in enumerate([
            (i,          VAL_FONT),
            (name,       VAL_FONT),
            (chart_type, Font(color="4A6A8A", size=10)),
        ], 2):
            c = ws.cell(row=r, column=ci, value=val)
            c.font = font; c.fill = HDR_FILL
            c.alignment = LEFT if ci == 3 else CENTER
            c.border = BORDER
        ws.row_dimensions[r].height = 20


def write_chart_sheet(wb, sheet_name: str, img_buf, w_in: float, h_in: float):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False
    _sheet_title(ws, sheet_name)
    img = XLImage(img_buf)
    img.width  = int(w_in * 96)
    img.height = int(h_in * 96)
    ws.add_image(img, "A2")


def write_report(year: int, month: int, imgs, daily):
    wb = openpyxl.Workbook()
    ws_idx = wb.active
    ws_idx.title = "サマリー"
    write_index(ws_idx, year, month, daily)

    for img, (name, (w, h), _) in zip(imgs, CHART_META):
        write_chart_sheet(wb, name, img, w, h)

    out = OUT_DIR / f"report_{year}_{month}.xlsx"
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


# ══════════════════════════════════════════════════════════════
# メイン
# ══════════════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description="日次データ → 分析レポートExcel")
    parser.add_argument("year",  type=int)
    parser.add_argument("month", type=int)
    args = parser.parse_args()

    out_path = OUT_DIR / f"report_{args.year}_{args.month}.xlsx"
    if out_path.exists():
        ans = input(f"\n{out_path.name} は既に存在します。上書きしますか？ [y/N]: ").strip().lower()
        if ans != "y":
            print("キャンセルしました。"); return

    print(f"データ読み込み中: daily_{args.year}_{args.month}.xlsx")
    daily, shohin = load_data(args.year, args.month)
    print(f"  日次: {len(daily)}日  商品別: {len(shohin)}行")

    print("チャート生成中...")
    chart_fns = [
        (chart_1, daily),
        (chart_2, daily),
        (chart_3, daily),
        (chart_4, daily),
        (chart_5, daily),
        (chart_6, daily),
        (chart_7, shohin),
        (chart_9, daily),
    ]
    imgs = []
    for i, (fn, arg) in enumerate(chart_fns, 1):
        imgs.append(fn(arg))
        print(f"  {i}/{len(chart_fns)} 完了")

    out = write_report(args.year, args.month, imgs, daily)
    print(f"\n出力完了: {out}")
    print(f"  シート構成: サマリー + {len(CHART_META)}チャートシート")


if __name__ == "__main__":
    main()
