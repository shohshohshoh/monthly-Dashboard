#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
日次データExcel → 分析レポートExcel 生成スクリプト

使い方:
    python create_report.py 2026 5
    → frontend/public/data/report_2026_5.xlsx を出力

チャート構成:
  ① 日次売上トレンド        折れ線 + 面（グロウ効果）
  ② 曜日別 平均売上         積み上げ棒（昼・夜）
  ③ 支払方法別 売上構成      ドーナツ（ラベル直接表示）
  ④ カテゴリ別 売上構成      円グラフ（ラベル直接表示）
  ⑤ 昼食 vs 夕食 比較       二重軸グラフ
  ⑥ 天気別 売上分布          箱ひげ図（昼天気→昼食、夜天気→夕食）
  ⑦ FOOD/DRINK ランキング   横棒（金額+数量テキスト）
  ⑨ 客単価 日次推移          折れ線（グロウ効果）
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
from matplotlib.colors import LinearSegmentedColormap
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
plt.rcParams.update({
    "font.family":        JP,
    "axes.unicode_minus": False,
    "axes.labelpad":      6,
    "xtick.major.pad":    5,
    "ytick.major.pad":    5,
})

C = {
    "bg":   "#03071e", "card": "#080f30",
    "c1":   "#00c8ff", "c2":   "#a855f7",
    "c3":   "#00ffa3", "c4":   "#ffb020",
    "c5":   "#ff3264", "c6":   "#ffe135",
    "c7":   "#00ecec", "c8":   "#ff6b6b",
    "grid": "#0d2040", "text": "#cce8ff", "sub":  "#3a5a7a",
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


# ══════════════════════════════════════════════════════════════
# 描画ヘルパー
# ══════════════════════════════════════════════════════════════

def dark_ax(ax, fig=None):
    """ダークテーマの軸スタイル（上・右スパイン削除、y軸グリッドのみ）"""
    if fig:
        fig.patch.set_facecolor(C["bg"])
    ax.set_facecolor(C["card"])
    ax.tick_params(colors=C["text"], labelsize=9, length=3,
                   direction="out", pad=4, width=0.8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    ax.spines["left"].set_edgecolor(C["grid"])
    ax.spines["bottom"].set_edgecolor(C["grid"])
    ax.grid(color=C["grid"], ls="--", lw=0.6, alpha=0.6, axis="y", zorder=0)


def neon_line(ax, xs, ys, color, lw=2.2, ms=5, marker="o", label="", zorder=3):
    """グロウ（発光）効果付き折れ線"""
    for width, alpha in [(9, 0.03), (6, 0.07), (3.5, 0.14)]:
        ax.plot(xs, ys, color=color, lw=width, alpha=alpha,
                solid_capstyle="round", zorder=zorder - 1)
    ax.plot(xs, ys, color=color, lw=lw, marker=marker, ms=ms,
            markerfacecolor=C["bg"], markeredgecolor=color,
            markeredgewidth=1.8, label=label,
            solid_capstyle="round", zorder=zorder)


def lbox(ax, x, y, text, color, fs=8.5, va="bottom", ha="center", **kw):
    """ボックス付きラベルアノテーション"""
    ax.text(x, y, text, ha=ha, va=va, fontsize=fs,
            color=color, fontweight="bold", zorder=6,
            bbox=dict(boxstyle="round,pad=0.3", facecolor=C["bg"],
                      edgecolor=color, alpha=0.80, linewidth=0.9),
            **kw)


def gradient_bars(ax, xs, hs, color_hi, color_lo=None, width=0.55,
                  bottoms=None, edgecolor=None, **kw):
    """グラデーション風バー（明→暗）"""
    if color_lo is None:
        r, g, b = int(color_hi[1:3], 16)/255, int(color_hi[3:5], 16)/255, int(color_hi[5:7], 16)/255
        color_lo = f"#{int(r*60):02x}{int(g*60):02x}{int(b*60):02x}"
    n = len(hs)
    if bottoms is None:
        bottoms = [0] * n
    cmap = LinearSegmentedColormap.from_list("g", [color_lo, color_hi])
    bars = []
    for i, (x, h, b) in enumerate(zip(xs, hs, bottoms)):
        col = cmap(i / max(n - 1, 1))
        bar = ax.bar(x, h, width, bottom=b, color=col,
                     edgecolor=edgecolor or C["bg"],
                     linewidth=0.8, zorder=2, **kw)
        bars.append(bar)
    return bars


def to_img(fig, dpi=150):
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
        for k in keys[7:]:
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
    """① 日次売上トレンド（グロウ折れ線 + グラデ面）"""
    op     = [d for d in daily if d["定休"] != "休"]
    xs     = [d["日"] for d in op]
    total  = [d["総売上高"] / 10000 for d in op]
    lunch  = [d["昼食売上"]  / 10000 for d in op]
    dinner = [d["夕食売上"]  / 10000 for d in op]
    avg    = np.mean(total)
    max_i  = int(np.argmax(total))

    fig, ax = plt.subplots(figsize=(12, 5))
    dark_ax(ax, fig)

    # グラデ面
    ax.fill_between(xs, total,  alpha=0.14, color=C["c1"], zorder=1)
    ax.fill_between(xs, lunch,  alpha=0.09, color=C["c4"], zorder=1)
    ax.fill_between(xs, dinner, alpha=0.09, color=C["c2"], zorder=1)

    neon_line(ax, xs, total,  C["c1"], lw=2.4, ms=4.5, label="総売上高")
    neon_line(ax, xs, lunch,  C["c4"], lw=1.7, ms=3.5, marker="^", label="昼食売上")
    neon_line(ax, xs, dinner, C["c2"], lw=1.7, ms=3.5, marker="s", label="夕食売上")

    # 平均ライン（グロウ）
    for w, a in [(4, 0.06), (2, 0.12)]:
        ax.axhline(avg, color=C["c3"], lw=w, alpha=a, ls="--")
    ax.axhline(avg, color=C["c3"], lw=1, ls="--", alpha=0.65,
               label=f"平均 {avg:.0f}万円")

    # 最大値アノテーション
    lbox(ax, xs[max_i], total[max_i] + max(total) * 0.04,
         f"MAX {xs[max_i]}日  ¥{total[max_i]:.0f}万", C["c4"], fs=8.5)

    ax.set_xlabel("日", color=C["sub"], fontsize=9)
    ax.set_ylabel("売上（万円）", color=C["sub"], fontsize=9)
    ax.set_ylim(0, max(total) * 1.22)
    ax.set_xlim(min(xs) - 0.5, max(xs) + 0.5)
    ax.set_title("① 日次売上トレンド",
                 color=C["text"], fontsize=14, fontweight="bold", pad=14)
    ax.legend(facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9.5, framealpha=0.85,
              loc="upper left", borderpad=0.7)
    fig.tight_layout()
    return to_img(fig)


def chart_2(daily):
    """② 曜日別平均売上（昼夜グラデ積み上げ棒）"""
    order = ["火", "水", "木", "金", "土", "祝日"]
    DAY_COLORS = {
        "火": C["c1"], "水": C["c1"], "木": C["c1"],
        "金": C["c4"], "土": C["c3"], "祝日": C["c5"],
    }

    wmap_l = defaultdict(list)
    wmap_d = defaultdict(list)
    for d in daily:
        if d["定休"] == "休":
            continue
        key = "祝日" if (d["曜日"] == "日" or d["祝日"]) else d["曜日"]
        if key in order:
            wmap_l[key].append(d["昼食売上"]  / 10000)
            wmap_d[key].append(d["夕食売上"] / 10000)

    avgs_l = [np.mean(wmap_l[w]) if wmap_l[w] else 0 for w in order]
    avgs_d = [np.mean(wmap_d[w]) if wmap_d[w] else 0 for w in order]
    cnts   = [len(wmap_l[w]) for w in order]
    totals = [l + d for l, d in zip(avgs_l, avgs_d)]
    top    = max(totals) if totals else 1

    fig, ax = plt.subplots(figsize=(9, 5.5))
    dark_ax(ax, fig)

    bar_w = 0.52
    x_pos = np.arange(len(order))

    for i, (w, al, ad) in enumerate(zip(order, avgs_l, avgs_d)):
        co = DAY_COLORS[w]
        # 昼（明るい）
        ax.bar(i, al, bar_w, color=co, alpha=0.95,
               edgecolor=C["bg"], linewidth=0.8, zorder=2)
        # 夜（暗い）
        ax.bar(i, ad, bar_w, bottom=al, color=co, alpha=0.45,
               edgecolor=C["bg"], linewidth=0.8, zorder=2)

    for i, (t, n) in enumerate(zip(totals, cnts)):
        if t > 0:
            lbox(ax, i, t + top * 0.025,
                 f"¥{t:.0f}万\nn={n}", C["text"], fs=8)

    ax.set_xticks(x_pos)
    ax.set_xticklabels(order, fontsize=11)
    ax.set_ylabel("平均売上（万円）", color=C["sub"], fontsize=9)
    ax.yaxis.label.set_color(C["sub"])
    ax.set_ylim(0, top * 1.32)

    p1 = mpatches.Patch(color=C["c1"],  alpha=0.95, label="昼食売上（濃色）")
    p2 = mpatches.Patch(color=C["text"], alpha=0.40, label="夕食売上（淡色）")
    ax.legend(handles=[p1, p2], facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9, framealpha=0.85, loc="upper left")
    ax.set_title("② 曜日別 平均売上",
                 color=C["text"], fontsize=14, fontweight="bold", pad=14)
    fig.tight_layout()
    return to_img(fig)


def chart_3(daily):
    """③ 支払方法別構成（ドーナツ・グロウリング・ラベル直接表示）"""
    lbls   = ["現金", "JCB", "千葉銀行", "アクアコイン", "PayPay", "ふるさと納税", "売掛金"]
    vals   = [sum(d[k] for d in daily) for k in lbls]
    colors = [C["c1"], C["c3"], C["c4"], C["c2"], C["c5"], C["c6"], C["c7"]]
    total  = sum(vals)

    inner_lbls = [
        f"{l}\n¥{v/10000:.0f}万\n{v/total*100:.1f}%"
        for l, v in zip(lbls, vals)
    ]

    fig, ax = plt.subplots(figsize=(10, 6.5))
    fig.patch.set_facecolor(C["bg"])
    ax.set_facecolor(C["bg"])
    ax.set_aspect("equal")

    # 外側グロウリング
    ax.pie(vals, colors=[(*matplotlib.colors.to_rgb(co), 0.18) for co in colors],
           startangle=90, counterclock=False,
           wedgeprops=dict(width=0.56, edgecolor="none"),
           radius=1.08)

    # メインリング（ラベルをリング内部に配置）
    ws, texts = ax.pie(
        vals, colors=colors, startangle=90,
        labels=inner_lbls, labeldistance=0.74,
        wedgeprops=dict(width=0.52, edgecolor=C["bg"], linewidth=2),
        counterclock=False
    )
    for t in texts:
        t.set_color("#03071e")   # 濃い紺色
        t.set_fontsize(7.5)
        t.set_fontweight("bold")
        t.set_horizontalalignment("center")

    # センター
    ax.add_patch(plt.Circle((0, 0), 0.40, color=C["card"], zorder=10))
    for r, txt, co, fs in [
        (0.12, "総売上", "white", 11),   # 白字・+2pt
        (-0.16, f"¥{total/10000:.0f}万", C["c1"], 13),
    ]:
        ax.text(0, r, txt, ha="center", va="center",
                fontsize=fs, color=co, fontweight="bold", zorder=11)

    ax.set_title("③ 支払方法別 売上構成",
                 color=C["text"], fontsize=14, fontweight="bold", pad=14)
    fig.tight_layout()
    return to_img(fig)


def chart_4(daily):
    """④ カテゴリ別売上構成（円グラフ・ラベル直接表示）"""
    cats   = ["FOOD", "DRINK", "売店", "その他"]
    vals   = [sum(d[k] for d in daily) for k in cats]
    colors = [C["c4"], C["c1"], C["c3"], C["c2"]]
    total  = sum(vals)

    outer_lbls = [
        f"{c}\n¥{v/10000:.0f}万\n({v/total*100:.1f}%)"
        for c, v in zip(cats, vals)
    ]
    explode = [0.04 if v == max(vals) else 0 for v in vals]

    fig, ax = plt.subplots(figsize=(8, 6))
    fig.patch.set_facecolor(C["bg"])
    ax.set_facecolor(C["bg"])
    ax.set_aspect("equal")

    ws, texts = ax.pie(
        vals, colors=colors, startangle=90,
        labels=outer_lbls, labeldistance=0.65,
        explode=explode,
        wedgeprops=dict(edgecolor=C["bg"], linewidth=2.5),
        counterclock=False
    )
    for t in texts:
        t.set_color("white")
        t.set_fontsize(9.5)
        t.set_fontweight("bold")
        t.set_horizontalalignment("center")

    ax.set_title("④ カテゴリ別 売上構成",
                 color=C["text"], fontsize=14, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_5(daily):
    """⑤ 昼食 vs 夕食（細バー・二重軸・バー内数値・客単価上表示）"""
    op     = [d for d in daily if d["定休"] != "休"]
    l_amt  = sum(d["昼食売上"]  for d in op)
    d_amt  = sum(d["夕食売上"]  for d in op)
    l_pax  = sum(d["昼食客数"] for d in op)
    d_pax  = sum(d["夕食客数"] for d in op)
    l_unit = l_amt / l_pax if l_pax else 0
    d_unit = d_amt / d_pax if d_pax else 0

    fig, ax1 = plt.subplots(figsize=(6, 5.5))
    dark_ax(ax1, fig)
    ax2 = ax1.twinx()
    dark_ax(ax2)
    ax2.set_facecolor("none")
    ax2.spines["top"].set_visible(False)
    ax2.spines["left"].set_visible(False)

    x = np.array([0.0, 1.0])
    w = 0.18

    amts = [l_amt / 10000, d_amt / 10000]
    paxs = [l_pax, d_pax]
    clrs = [C["c4"], C["c1"]]

    # 売上バー（左軸）
    for xi, (v, co) in enumerate(zip(amts, clrs)):
        ax1.bar(xi - w / 2 - 0.01, v, w, color=co,
                edgecolor=C["bg"], linewidth=0.8, zorder=3)
        ax1.text(xi - w / 2 - 0.01, v * 0.5, f"¥{v:.0f}万",
                 ha="center", va="center", fontsize=8.5,
                 color="white", fontweight="bold", rotation=90, zorder=4)

    # 客数バー（右軸）
    for xi, (v, co) in enumerate(zip(paxs, clrs)):
        ax2.bar(xi + w / 2 + 0.01, v, w, color=co,
                alpha=0.38, edgecolor=co, linewidth=1.2, zorder=3)
        ax2.text(xi + w / 2 + 0.01, v * 0.5, f"{v:,}名",
                 ha="center", va="center", fontsize=8.5,
                 color="white", fontweight="bold", rotation=90, zorder=4)

    # 客単価をバー上にボックス表示
    for xi, (u, co) in enumerate(zip([l_unit, d_unit], clrs)):
        lbox(ax1, xi - w / 2 - 0.01,
             max(amts) * 1.06,
             f"客単価\n¥{u:,.0f}", co, fs=8.5)

    ax1.set_xticks(x)
    ax1.set_xticklabels(["昼食", "夕食"], fontsize=13, fontweight="bold",
                        color=C["text"])
    ax1.tick_params(axis="x", length=0)
    ax1.set_ylabel("売上金額（万円）", color=C["sub"], fontsize=9)
    ax1.yaxis.label.set_color(C["sub"])
    ax1.set_ylim(0, max(amts) * 1.48)
    ax1.set_xlim(-0.55, 1.55)

    ax2.set_ylabel("来客数（名）", color=C["sub"], fontsize=9)
    ax2.yaxis.label.set_color(C["sub"])
    ax2.tick_params(colors=C["text"])
    ax2.set_ylim(0, max(paxs) * 1.48)
    ax2.grid(False)

    ax1.text(0.5, 1.02,
             "濃色バー = 売上（左軸）　淡色バー = 客数（右軸）",
             ha="center", transform=ax1.transAxes,
             fontsize=8, color=C["sub"])
    ax1.set_title("⑤ 昼食 vs 夕食 比較",
                  color=C["text"], fontsize=14, fontweight="bold", pad=20)
    fig.tight_layout()
    return to_img(fig)


def chart_6(daily):
    """⑥ 天気別売上分布（昼天気→昼食売上、夜天気→夕食売上・散布点オーバーレイ）"""
    weather_colors = {"晴": C["c6"], "曇": C["sub"], "雨": C["c1"]}
    weather_order  = ["晴", "曇", "雨"]

    def make_groups(t_key, s_key):
        g = {"晴": [], "曇": [], "雨": []}
        for d in daily:
            if d["定休"] == "休":
                continue
            t = d.get(t_key, "")
            if t in g:
                g[t].append(d[s_key] / 10000)
        return g

    def draw_box(ax, groups, title, ylabel):
        dark_ax(ax)
        lbls = [k for k in weather_order if groups[k]]
        data = [groups[k] for k in lbls]
        bp = ax.boxplot(data, tick_labels=lbls, patch_artist=True,
                        medianprops=dict(color=C["c3"], lw=2.8),
                        whiskerprops=dict(color=C["text"], lw=1.2, ls="--"),
                        capprops=dict(color=C["text"], lw=1.5),
                        flierprops=dict(markerfacecolor=C["c5"], markersize=6,
                                        markeredgecolor=C["c5"], alpha=0.7),
                        boxprops=dict(linewidth=1.5))
        for patch, k in zip(bp["boxes"], lbls):
            patch.set_facecolor(weather_colors[k])
            patch.set_alpha(0.30)

        # 散布点オーバーレイ（jitter）
        np.random.seed(42)
        for i, (k, pts) in enumerate(zip(lbls, data), 1):
            jitter = np.random.uniform(-0.12, 0.12, len(pts))
            ax.scatter(np.full(len(pts), i) + jitter, pts,
                       color=weather_colors[k], s=28, alpha=0.65,
                       edgecolors=C["bg"], linewidths=0.6, zorder=4)

        # n数・平均ラベル
        all_vals = [v for g in data for v in g]
        y_min = min(all_vals); y_range = max(all_vals) - y_min
        for i, (k, pts) in enumerate(zip(lbls, data), 1):
            ax.text(i, y_min - y_range * 0.10,
                    f"n={len(pts)}日\nμ={np.mean(pts):.0f}万",
                    ha="center", va="top", fontsize=7.5,
                    color=weather_colors[k], fontweight="bold")

        ax.set_ylabel(ylabel, color=C["sub"], fontsize=9)
        ax.yaxis.label.set_color(C["sub"])
        ax.set_title(title, color=C["text"], fontsize=11, fontweight="bold")

    g_hiru = make_groups("昼天気", "昼食売上")
    g_yoru = make_groups("夜天気", "夕食売上")

    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(11, 5.5))
    fig.patch.set_facecolor(C["bg"])
    draw_box(ax1, g_hiru, "昼天気別 昼食売上分布", "昼食売上（万円）")
    draw_box(ax2, g_yoru, "夜天気別 夕食売上分布", "夕食売上（万円）")

    fig.suptitle("⑥ 天気別 売上分布",
                 color=C["text"], fontsize=14, fontweight="bold", y=1.02)
    fig.tight_layout()
    return to_img(fig)


def chart_7(shohin):
    """⑦ FOOD / DRINK 売れ筋ランキング Top10（売上金額バー・数量をテキスト表示）"""
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

    fig, (ax_f, ax_d) = plt.subplots(1, 2, figsize=(13, 5.5))
    fig.patch.set_facecolor(C["bg"])

    def _hbar_qty(ax, names, amts, qtys, cm_name, color_title, title, unit):
        dark_ax(ax)
        n = len(names)
        grad = [getattr(plt.cm, cm_name)(0.35 + 0.65 * i / max(n - 1, 1))
                for i in range(n)]
        bars = ax.barh(names, amts, color=grad, edgecolor=C["bg"])
        max_a = max(amts) if amts else 1
        for bar, a, q in zip(bars, amts, qtys):
            ax.text(bar.get_width() + max_a * 0.015,
                    bar.get_y() + bar.get_height() / 2,
                    f"¥{a:.1f}万  {q:,}{unit}",
                    va="center", fontsize=8, color=C["text"])
        ax.set_xlabel("売上金額（万円）", color=C["text"])
        ax.xaxis.label.set_color(C["text"])
        ax.set_xlim(0, max_a * 1.6)
        ax.set_title(title, color=color_title, fontsize=11, fontweight="bold")

    _hbar_qty(ax_f, f_names, f_amts, f_qtys, "YlOrRd", C["c4"],
              "FOOD 売上金額 Top10", "個")
    _hbar_qty(ax_d, d_names, d_amts, d_qtys, "Blues",  C["c1"],
              "DRINK 売上金額 Top10", "杯")

    fig.suptitle("⑦ FOOD / DRINK 売れ筋ランキング Top10",
                 color=C["text"], fontsize=14, fontweight="bold")
    fig.tight_layout()
    return to_img(fig)


def chart_9(daily):
    """⑨ 客単価 日次推移（グロウ折れ線・エリア間シェード）"""
    op = [d for d in daily
          if d["定休"] != "休" and d["昼食客数"] > 0 and d["夕食客数"] > 0]
    days   = [d["日"] for d in op]
    l_unit = [d["昼食売上"] / d["昼食客数"] for d in op]
    d_unit = [d["夕食売上"] / d["夕食客数"] for d in op]
    l_avg  = np.mean(l_unit)
    d_avg  = np.mean(d_unit)

    fig, ax = plt.subplots(figsize=(12, 5))
    dark_ax(ax, fig)

    ax.fill_between(days, l_unit, d_unit,
                    where=[l > d for l, d in zip(l_unit, d_unit)],
                    alpha=0.08, color=C["c4"])
    ax.fill_between(days, l_unit, d_unit,
                    where=[d >= l for l, d in zip(l_unit, d_unit)],
                    alpha=0.08, color=C["c1"])

    neon_line(ax, days, l_unit, C["c4"], lw=2.2, ms=5, label="昼食 客単価")
    neon_line(ax, days, d_unit, C["c1"], lw=2.2, ms=5, marker="s", label="夕食 客単価")

    # 平均ライン
    for color, avg, label in [(C["c4"], l_avg, f"昼平均 ¥{l_avg:,.0f}"),
                               (C["c1"], d_avg, f"夜平均 ¥{d_avg:,.0f}")]:
        for w, a in [(3, 0.06), (1.5, 0.12)]:
            ax.axhline(avg, color=color, lw=w, alpha=a, ls=":")
        ax.axhline(avg, color=color, lw=0.9, ls=":", alpha=0.6, label=label)

    ax.yaxis.set_major_formatter(plt.FuncFormatter(lambda v, _: f"¥{v:,.0f}"))
    ax.set_xlabel("日", color=C["sub"], fontsize=9)
    ax.set_ylabel("客単価（円）", color=C["sub"], fontsize=9)
    ax.set_xlim(min(days) - 0.5, max(days) + 0.5)
    ax.set_title("⑨ 客単価 日次推移",
                 color=C["text"], fontsize=14, fontweight="bold", pad=14)
    ax.legend(facecolor=C["card"], edgecolor=C["grid"],
              labelcolor=C["text"], fontsize=9.5, framealpha=0.85,
              loc="upper left", borderpad=0.7)
    fig.tight_layout()
    return to_img(fig)


# ══════════════════════════════════════════════════════════════
# レポートExcel 書き出し
# ══════════════════════════════════════════════════════════════
CHART_META = [
    ("① 日次売上トレンド",        (12, 5.0), "折れ線 + 面（グロウ）"),
    ("② 曜日別 平均売上",         ( 9, 5.5), "積み上げ棒グラフ"),
    ("③ 支払方法別 構成",         (10, 6.5), "ドーナツグラフ"),
    ("④ カテゴリ別 構成",         ( 8, 6.0), "円グラフ"),
    ("⑤ 昼食 vs 夕食",           ( 6, 5.5), "二重軸グラフ"),
    ("⑥ 天気別 売上分布",         (11, 5.5), "箱ひげ図 + 散布点"),
    ("⑦ FOOD・DRINK ランキング",  (13, 5.5), "横棒グラフ"),
    ("⑨ 客単価 日次推移",         (12, 5.0), "折れ線グラフ（グロウ）"),
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
    ws.column_dimensions["D"].width = 26

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
