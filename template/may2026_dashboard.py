#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
2026年5月 営業日報 ダッシュボード PowerPoint 生成（3Dグラフ版）
フォント全+5pt / 商品名・曜日・昼食夕食は白 / 3Dスタイル
"""
import io, warnings
warnings.filterwarnings("ignore")
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from mpl_toolkits.mplot3d import Axes3D
from collections import defaultdict
from pptx import Presentation
from pptx.util import Cm, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import openpyxl

import matplotlib.font_manager as fm
_avail = {f.name for f in fm.fontManager.ttflist}
JP = next((f for f in ["Yu Gothic","Meiryo","MS Gothic"] if f in _avail), "sans-serif")
plt.rcParams.update({"font.family": JP, "axes.unicode_minus": False})

# ══════════════════════════════════════════════════════════════════════════════
# [CONFIG] 基本設定
# ══════════════════════════════════════════════════════════════════════════════
TITLE    = "2026年5月 営業日報 ダッシュボード"
SUBTITLE = "集計期間: 2026年5月1日〜5月31日  ｜  稼働日: 27日 ／ 定休: 4日（月曜）"
OUTPUT   = "my_dashboard_2026年5月.pptx"

# [CONFIG] カラーパレット
COLORS = {
    "BG":   (3,   7,  30),
    "CARD": (8,  18,  55),
    "C1":   (0,  180, 255),
    "C2":   (155, 93, 229),
    "C3":   (0,  245, 160),
    "C4":   (255,159,  28),
    "C5":   (255, 50, 100),
    "C6":   (255,214,  10),
    "C7":   (0,  236, 236),
    "GRID": (15,  45,  74),
    "TEXT": (208,232, 255),
    "SUB":  (74, 106, 138),
}

# [CONFIG] KPI カード（3個・千円単位）
KPI_ITEMS = [
    {"label": "総売上高（税込）", "value": "20,989千円", "color": "C1"},
    {"label": "FOOD 売 上",       "value": "14,975千円", "color": "C3"},
    {"label": "来  客  数",       "value": "4,054 名",   "color": "C4"},
]

# ══════════════════════════════════════════════════════════════════════════════
# [CONFIG] データ読み込み
# ══════════════════════════════════════════════════════════════════════════════
SRC = r"../data/★営業日報2026年5月.xlsx"

def load_data():
    wb   = openpyxl.load_workbook(SRC, data_only=True)
    ws   = wb["データ"]
    rows = list(ws.iter_rows(min_row=1, max_row=54, values_only=True))
    COL0 = 5
    date_r  = rows[0][COL0:COL0+31]
    hol_r   = rows[1][COL0:COL0+31]
    wday_r  = rows[3][COL0:COL0+31]
    sales_r = rows[4][COL0:COL0+31]
    days, wdays, closed, daily = [], [], [], []
    for d, h, w, s in zip(date_r, hol_r, wday_r, sales_r):
        if d is not None:
            days.append(d.day)
            wdays.append(w or "")
            closed.append(h == "休")
            daily.append(s or 0)
    ws_i = wb["商品別"]
    food_s, drink_s = defaultdict(int), defaultdict(int)
    for r in ws_i.iter_rows(min_row=3, max_row=95, values_only=True):
        if r[4] and r[7]:  food_s[r[4]]  += (r[7]  or 0)
        if r[8] and r[11]: drink_s[r[8]] += (r[11] or 0)
    top_food  = sorted(food_s.items(),  key=lambda x: x[1], reverse=True)[:6]
    top_drink = sorted(drink_s.items(), key=lambda x: x[1], reverse=True)[:6]
    M = {
        "total":20_989_657, "cash":8_692_051,
        "jcb":3_806_445,    "chiba":5_643_149,
        "aqua":567_876,     "paypay":1_076_179, "kake":1_215_401,
        "food":14_975_700,  "drink":3_000_230,
        "baiten":230_260,   "sonota":2_783_467,
        "lunch_pax":2_520,  "din_pax":1_534,
        "lunch_amt":9_681_482, "din_amt":11_308_175,
    }
    return days, wdays, closed, daily, top_food, top_drink, M

days, wdays, closed, daily, top_food, top_drink, M = load_data()

# ══════════════════════════════════════════════════════════════════════════════
# ユーティリティ
# ══════════════════════════════════════════════════════════════════════════════
def c(key):    return tuple(v/255 for v in COLORS[key])
def _rgb(key): return RGBColor(*COLORS[key])

def buf(fig, dpi=160):
    b = io.BytesIO()
    fig.savefig(b, format="png", dpi=dpi, bbox_inches="tight",
                facecolor=fig.get_facecolor(), edgecolor="none")
    b.seek(0); plt.close(fig); return b

def dark_ax3d(ax, fig=None):
    """Axes3D 用ダークネオンスタイル"""
    card   = c("CARD")
    grid_c = c("GRID")
    if fig:
        fig.patch.set_facecolor(c("BG"))
    for pane in [ax.xaxis.pane, ax.yaxis.pane, ax.zaxis.pane]:
        pane.fill = True
        pane.set_facecolor((*card, 0.88))
        pane.set_edgecolor(grid_c)
    ax.tick_params(axis='x', colors='white', labelsize=13)
    ax.tick_params(axis='y', colors='white', labelsize=13)
    ax.tick_params(axis='z', colors='white', labelsize=13)
    ax.xaxis.label.set_color('white')
    ax.yaxis.label.set_color('white')
    ax.zaxis.label.set_color('white')

def dark_ax(ax, fig=None, bg_key="CARD"):
    bg = c(bg_key)
    if fig: fig.patch.set_facecolor(bg)
    ax.set_facecolor(bg)
    ax.tick_params(colors="white", labelsize=13)
    for sp in ax.spines.values(): sp.set_edgecolor(c("GRID"))
    ax.grid(axis="y", color=c("GRID"), ls="--", lw=0.4, alpha=0.7)

def glow(ax, x, y, col, lw=2.0, label=None):
    for w, a in [(12,.025),(6,.07),(3,.18),(lw,1.)]:
        kw = dict(color=col, lw=w, alpha=a, solid_capstyle="round")
        if w == lw and label: kw["label"] = label
        ax.plot(x, y, **kw)

def glow3d(ax, xi_list, z_list, col, lw=2.2, y=0.225):
    xs = [xi + 0.31 for xi in xi_list]
    ys = [y] * len(xi_list)
    for w, a in [(12,.02),(6,.06),(3,.15),(lw,1.)]:
        ax.plot3D(xs, ys, z_list, color=col, lw=w, alpha=a)

def neon_bar3d(ax, xi, yi, val, dx=0.62, dy=0.42, col=None, alpha=0.88):
    if col is None: col = c("C1")
    ax.bar3d(xi, yi, 0, dx, dy, max(val, 1),
             color=col, alpha=alpha,
             edgecolor=(*c("BG"), 0.5), linewidth=0.35, shade=True)

def z_万(ax):
    ax.zaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v, _: f"{int(v/10000)}万"))

# ══════════════════════════════════════════════════════════════════════════════
# [CONFIG] チャート定義（6枚・3Dスタイル）
# ══════════════════════════════════════════════════════════════════════════════

def chart_1():
    """① 日次売上トレンド — 3D棒グラフ＋ネオングロウ折れ線"""
    fig = plt.figure(figsize=(9.5, 4.8))
    ax  = fig.add_subplot(111, projection='3d')
    dark_ax3d(ax, fig)

    max_s = max(daily) if daily else 1
    dx, dy = 0.62, 0.42

    for xi, (s, cl) in enumerate(zip(daily, closed)):
        if cl:
            col = (*c("GRID"), 0.45)
        else:
            ratio = s / max_s
            col = tuple(c("C1")[j]*ratio + c("C2")[j]*(1-ratio) for j in range(3))
        neon_bar3d(ax, xi, 0, s, dx=dx, dy=dy, col=col)

    ox_i = [xi for xi, cl in enumerate(closed) if not cl]
    oy   = [daily[xi] for xi in ox_i]
    glow3d(ax, ox_i, oy, c("C1"), lw=2.2, y=dy/2)

    avg = np.mean(oy) if oy else 0
    ax.plot3D([0, len(days)], [dy/2, dy/2], [avg, avg],
              color=c("C4"), lw=1.8, ls="--", alpha=0.9)
    ax.text(len(days)+0.3, dy/2, avg*1.05,
            f"平均 {avg/10000:.0f}万",
            color=c("C4"), fontsize=12, fontweight='bold')

    xt_idx = list(range(0, len(days), 5))
    ax.set_xticks([i + dx/2 for i in xt_idx])
    ax.set_xticklabels([f"{days[i]}日" for i in xt_idx],
                       fontsize=11, color='white')
    ax.yaxis.set_visible(False)
    z_万(ax)
    ax.set_zlabel("売上高", color='white', fontsize=12, labelpad=8)
    ax.set_xlim(-0.5, len(days)+0.5)
    ax.set_ylim(-0.3, 1.5)
    ax.set_zlim(0, max_s*1.35)
    ax.view_init(elev=22, azim=-65)
    ax.set_title("日次売上トレンド（2026年5月）",
                 color=c("TEXT"), fontsize=15, fontweight="bold", pad=12)
    fig.tight_layout(pad=0.5)
    return buf(fig)

def chart_2():
    """② 曜日別平均売上 — 3D棒グラフ"""
    order = ["月","火","水","木","金","土","日"]
    wmap  = defaultdict(list)
    for s, w, cl in zip(daily, wdays, closed):
        if not cl and w: wmap[w].append(s)
    avgs = [np.mean(wmap[w]) if wmap[w] else 0 for w in order]
    cnts = [len(wmap[w]) for w in order]
    pal  = {
        "月": c("GRID"), "火": c("C1"), "水": c("C1"),
        "木": c("C1"),   "金": c("C4"), "土": c("C3"), "日": c("C3"),
    }

    fig = plt.figure(figsize=(5.0, 4.8))
    ax  = fig.add_subplot(111, projection='3d')
    dark_ax3d(ax, fig)

    dx, dy = 0.6, 0.42
    max_a  = max(avgs) if avgs else 1

    for xi, (w, avg, n) in enumerate(zip(order, avgs, cnts)):
        if avg > 0:
            neon_bar3d(ax, xi, 0, avg, dx=dx, dy=dy, col=pal[w])
            ax.text(xi+dx/2, dy+0.12, avg*1.07,
                    f"{avg/10000:.1f}万\nn={n}",
                    ha='center', va='bottom', fontsize=11,
                    color=c("TEXT"), fontweight='bold')

    ax.set_xticks([i+dx/2 for i in range(len(order))])
    ax.set_xticklabels(order, fontsize=14, color='white')  # 曜日は白
    ax.yaxis.set_visible(False)
    z_万(ax)
    ax.set_zlabel("平均売上", color='white', fontsize=12, labelpad=8)
    ax.set_xlim(-0.5, len(order)+0.5)
    ax.set_ylim(-0.3, 1.5)
    ax.set_zlim(0, max_a*1.5)
    ax.view_init(elev=25, azim=-55)
    ax.set_title("曜日別 平均売上",
                 color=c("TEXT"), fontsize=15, fontweight="bold", pad=12)
    fig.tight_layout(pad=0.5)
    return buf(fig)

def chart_3():
    """③ 支払方法別 売上構成 — グロウドーナツ"""
    labels = ["現金","JCB","千葉銀行","アクアコイン","PayPay","売掛金"]
    vals   = [M["cash"],M["jcb"],M["chiba"],M["aqua"],M["paypay"],M["kake"]]
    cols   = [c("C1"),c("C3"),c("C4"),c("C2"),c("C5"),c("C7")]

    fig, ax = plt.subplots(figsize=(5.0, 4.8))
    fig.patch.set_facecolor(c("CARD"))
    ax.set_facecolor(c("CARD"))
    ax.set_aspect("equal")

    # Shadow layer
    ax.pie(vals, colors=[(0,0,0,0.3)]*len(vals), startangle=90,
           counterclock=False,
           wedgeprops=dict(width=0.5, edgecolor='none'),
           radius=1.06, center=(0.04,-0.07))

    _, _, ats = ax.pie(
        vals, colors=cols, startangle=90,
        autopct=lambda p: f"{p:.0f}%" if p > 4 else "",
        wedgeprops=dict(width=0.52, edgecolor=c("BG"), linewidth=2.5),
        pctdistance=0.76, counterclock=False
    )
    for at in ats:
        at.set_color("white"); at.set_fontsize(13); at.set_fontweight("bold")

    ax.add_patch(plt.Circle((0,0), 1.03, fill=False,
                             color=c("C1"), lw=2.5, alpha=0.5))
    ax.add_patch(plt.Circle((0,0), 1.07, fill=False,
                             color=c("C1"), lw=7, alpha=0.1))
    ax.add_patch(plt.Circle((0,0), 0.38, color=c("CARD"), zorder=10))
    ax.text(0, 0.12, "総売上高", ha="center", va="center",
            fontsize=11, color=c("SUB"))
    ax.text(0,-0.15, f"¥{M['total']/10000:.0f}万",
            ha="center", va="center",
            fontsize=13, color=c("C1"), fontweight="bold")

    patches = [mpatches.Patch(color=co, label=l)
               for co, l in zip(cols, labels)]
    ax.legend(handles=patches, loc="lower center",
              bbox_to_anchor=(0.5,-0.22), ncol=3,
              fontsize=11, facecolor=c("BG"), edgecolor=c("GRID"),
              labelcolor="white", framealpha=0.9)
    ax.set_title("支払方法別 構成",
                 color=c("TEXT"), fontsize=15, fontweight="bold", pad=8)
    fig.tight_layout(pad=0.3)
    return buf(fig)

def chart_4():
    """④ カテゴリ別 売上金額 — 3D縦棒グラフ"""
    labels = ["FOOD","DRINK","売店","その他"]
    vals   = [M["food"],M["drink"],M["baiten"],M["sonota"]]
    cols   = [c("C4"),c("C1"),c("C3"),c("C2")]
    total  = sum(vals)

    fig = plt.figure(figsize=(5.5, 4.8))
    ax  = fig.add_subplot(111, projection='3d')
    dark_ax3d(ax, fig)

    dx, dy = 0.6, 0.42
    max_v  = max(vals)

    for xi, (lbl, v, col) in enumerate(zip(labels, vals, cols)):
        neon_bar3d(ax, xi, 0, v, dx=dx, dy=dy, col=col)
        ax.text(xi+dx/2, dy+0.12, v*1.07,
                f"¥{v/10000:.0f}万\n({v/total*100:.0f}%)",
                ha='center', va='bottom', fontsize=12,
                color=c("TEXT"), fontweight='bold')

    ax.set_xticks([i+dx/2 for i in range(len(labels))])
    ax.set_xticklabels(labels, fontsize=13, color='white')
    ax.yaxis.set_visible(False)
    z_万(ax)
    ax.set_zlabel("売上金額", color='white', fontsize=12, labelpad=8)
    ax.set_xlim(-0.5, len(labels)+0.5)
    ax.set_ylim(-0.3, 1.5)
    ax.set_zlim(0, max_v*1.45)
    ax.view_init(elev=25, azim=-55)
    ax.set_title("カテゴリ別 売上金額",
                 color=c("TEXT"), fontsize=15, fontweight="bold", pad=12)
    fig.tight_layout(pad=0.5)
    return buf(fig)

def chart_5():
    """⑤ 昼食 vs 夕食 — 3Dグループ棒グラフ（売上・客数）"""
    amts = [M["lunch_amt"], M["din_amt"]]
    pax  = [M["lunch_pax"], M["din_pax"]]
    cols = [c("C4"), c("C1")]
    lbl  = ["昼食", "夕食"]

    fig = plt.figure(figsize=(6.5, 4.8))
    ax1 = fig.add_subplot(121, projection='3d')
    ax2 = fig.add_subplot(122, projection='3d')
    dark_ax3d(ax1, fig)
    dark_ax3d(ax2)

    dx, dy = 0.55, 0.42

    for xi, (a, col) in enumerate(zip(amts, cols)):
        neon_bar3d(ax1, xi, 0, a, dx=dx, dy=dy, col=col)
        ax1.text(xi+dx/2, dy+0.1, a*1.07,
                 f"{a/10000:.0f}万",
                 ha='center', va='bottom', fontsize=13,
                 color=c("TEXT"), fontweight='bold')
    ax1.set_xticks([i+dx/2 for i in range(2)])
    ax1.set_xticklabels(lbl, fontsize=14, color='white')  # 昼食夕食は白
    ax1.yaxis.set_visible(False)
    ax1.zaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v,_: f"{int(v/10000)}万"))
    ax1.set_xlim(-0.5, 2.0); ax1.set_ylim(-0.3, 1.2)
    ax1.set_zlim(0, max(amts)*1.45)
    ax1.view_init(elev=22, azim=-55)
    ax1.set_title("売上金額", color=c("TEXT"),
                  fontsize=14, fontweight='bold', pad=8)

    for xi, (p, col) in enumerate(zip(pax, cols)):
        neon_bar3d(ax2, xi, 0, p, dx=dx, dy=dy, col=col)
        ax2.text(xi+dx/2, dy+0.1, p*1.07,
                 f"{p:,}名",
                 ha='center', va='bottom', fontsize=13,
                 color=c("TEXT"), fontweight='bold')
    ax2.set_xticks([i+dx/2 for i in range(2)])
    ax2.set_xticklabels(lbl, fontsize=14, color='white')  # 昼食夕食は白
    ax2.yaxis.set_visible(False)
    ax2.zaxis.set_major_formatter(
        mticker.FuncFormatter(lambda v,_: f"{int(v):,}名"))
    ax2.set_xlim(-0.5, 2.0); ax2.set_ylim(-0.3, 1.2)
    ax2.set_zlim(0, max(pax)*1.55)
    ax2.view_init(elev=22, azim=-55)
    ax2.set_title("来客数", color=c("TEXT"),
                  fontsize=14, fontweight='bold', pad=8)

    fig.suptitle("昼食 vs 夕食 比較",
                 color=c("TEXT"), fontsize=15, fontweight="bold")
    fig.tight_layout(pad=0.5)
    return buf(fig)

def chart_6():
    """⑥ 売れ筋商品ランキング — FOOD/DRINK Top6 3D棒グラフ"""
    fig = plt.figure(figsize=(8.5, 4.8))
    ax1 = fig.add_subplot(121, projection='3d')
    ax2 = fig.add_subplot(122, projection='3d')
    dark_ax3d(ax1, fig)
    dark_ax3d(ax2)

    def _rank3d(ax, items, base_col, title):
        names = [x[0][:6] for x in items]   # 6文字以内
        vals  = [x[1] for x in items]
        max_v = max(vals) if vals else 1
        dx, dy = 0.65, 0.42

        for xi, (n, v) in enumerate(zip(names, vals)):
            ratio = 1.0 - 0.55 * xi / max(len(names)-1, 1)
            col = tuple(
                base_col[j]*ratio + c("CARD")[j]*(1-ratio)
                for j in range(3)
            )
            neon_bar3d(ax, xi, 0, v, dx=dx, dy=dy, col=col)
            ax.text(xi+dx/2, dy+0.1, v*1.07,
                    f"{v/10000:.1f}万",
                    ha='center', va='bottom', fontsize=11,
                    color=c("TEXT"), fontweight='bold')

        ax.set_xticks([i+dx/2 for i in range(len(names))])
        ax.set_xticklabels(names, fontsize=10, color='white')  # 商品名は白
        ax.yaxis.set_visible(False)
        z_万(ax)
        ax.set_zlabel("売上", color='white', fontsize=11, labelpad=6)
        ax.set_xlim(-0.5, len(names)+0.5)
        ax.set_ylim(-0.3, 1.5)
        ax.set_zlim(0, max_v*1.45)
        ax.view_init(elev=22, azim=-55)
        ax.set_title(title, color=c("TEXT"),
                     fontsize=13, fontweight="bold", pad=8)

    _rank3d(ax1, top_food,  c("C4"), "FOOD Top6")
    _rank3d(ax2, top_drink, c("C1"), "DRINK Top6")
    fig.tight_layout(pad=0.5)
    return buf(fig)

# ══════════════════════════════════════════════════════════════════════════════
# グラフ生成
# ══════════════════════════════════════════════════════════════════════════════
print("チャートを生成中...")
charts = [chart_1(), chart_2(), chart_3(), chart_4(), chart_5(), chart_6()]
print(f"  → {len(charts)}チャート完了")

# ══════════════════════════════════════════════════════════════════════════════
# PowerPoint 組み立て
# ══════════════════════════════════════════════════════════════════════════════
print("PowerPoint を組み立て中...")
prs = Presentation()
prs.slide_width  = Cm(33.87)
prs.slide_height = Cm(19.05)
slide = prs.slides.add_slide(prs.slide_layouts[6])

def _shape(l,t,w,h,fill=None,border=None,bw=0.75,g1=None,g2=None,ang=135):
    sh = slide.shapes.add_shape(1,Cm(l),Cm(t),Cm(w),Cm(h))
    if g1 and g2:
        sh.fill.gradient(); sh.fill.gradient_angle = ang
        sh.fill.gradient_stops[0].position = 0.0
        sh.fill.gradient_stops[0].color.rgb = RGBColor(*g1)
        sh.fill.gradient_stops[1].position = 1.0
        sh.fill.gradient_stops[1].color.rgb = RGBColor(*g2)
    elif fill:
        sh.fill.solid(); sh.fill.fore_color.rgb = RGBColor(*fill)
    else:
        sh.fill.background()
    if border: sh.line.color.rgb = RGBColor(*border); sh.line.width = Pt(bw)
    else: sh.line.fill.background()
    return sh

def _txt(text,l,t,w,h,size=9,bold=False,col="TEXT",
         align=PP_ALIGN.LEFT,italic=False):
    tb = slide.shapes.add_textbox(Cm(l),Cm(t),Cm(w),Cm(h))
    tb.text_frame.word_wrap = True
    p = tb.text_frame.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = text
    r.font.size=Pt(size); r.font.bold=bold; r.font.italic=italic
    r.font.color.rgb = RGBColor(*COLORS[col])

def _img(stream,l,t,w,h):
    slide.shapes.add_picture(stream,Cm(l),Cm(t),Cm(w),Cm(h))

# 背景・ヘッダー
_shape(0,0,33.87,19.05, g1=COLORS["BG"],g2=(6,14,38),ang=130)
_shape(0,0,33.87,2.20,  g1=(4,12,48),g2=(8,22,60),ang=0)
_shape(0,2.18,33.87,0.05, g1=(0,60,150),g2=COLORS["C1"],ang=0)

_txt(f"◈  {TITLE}", 0.4,0.10,20,1.2, size=23, bold=True, col="C1")   # +5pt
_txt(SUBTITLE,       0.4,1.32,20,0.75, size=13, col="SUB")             # +5pt

# KPI カード（3枚・幅3.85cm）
kx0=22.0; kw=3.85; kg=0.12
for i, k in enumerate(KPI_ITEMS):
    lx = kx0 + i*(kw+kg)
    _shape(lx,0.12,kw,2.0, g1=(6,18,50),g2=(4,12,38),
           border=COLORS[k["color"]],bw=0.75)
    _txt(k["label"], lx+0.12,0.18,kw-0.24,0.65,
         size=12, col="SUB", align=PP_ALIGN.CENTER)                     # +5pt
    _txt(k["value"], lx+0.12,0.82,kw-0.24,1.0,
         size=16, bold=True, col=k["color"], align=PP_ALIGN.CENTER)     # +5pt

# チャート配置（3×2）
R1  = 2.80
R2  = 11.05
RH  = 7.75

row1 = [(0.20,14.70),(15.10,8.60),(23.90, 9.77)]
row2 = [(0.20, 8.90),(9.30,12.40),(21.90,11.77)]
lbls1 = ["① 日次売上トレンド","③ 支払方法別構成","④ カテゴリ別売上"]
lbls2 = ["② 曜日別 平均売上", "⑤ 昼食 vs 夕食",  "⑥ 商品ランキング"]

for (lx,lw), im, lb in zip(row1, [charts[0],charts[2],charts[3]], lbls1):
    _shape(lx,R1,lw,RH, g1=(5,11,30),g2=(3,9,25),border=COLORS["C1"],bw=0.6)
    _img(im,lx,R1,lw,RH)
    _txt(lb, lx+0.15,R1-0.44,lw,0.4, size=12, bold=True, col="C7")    # +5pt

for (lx,lw), im, lb in zip(row2, [charts[1],charts[4],charts[5]], lbls2):
    _shape(lx,R2,lw,RH, g1=(5,11,30),g2=(3,9,25),border=COLORS["C1"],bw=0.6)
    _img(im,lx,R2,lw,RH)
    _txt(lb, lx+0.15,R2-0.44,lw,0.4, size=12, bold=True, col="C7")    # +5pt

_txt("データソース: ★営業日報2026年5月.xlsx  ／  Generated by may2026_dashboard.py",
     0.3,18.72,33.0,0.33, size=12, col="SUB", italic=True)             # +5pt

prs.save(OUTPUT)
print(f"完了: {OUTPUT}")
