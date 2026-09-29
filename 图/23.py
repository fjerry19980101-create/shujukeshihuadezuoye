# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:15:27 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""23 柱形折线图 复现（尺寸/坐标/配色/字号均为参考图像素级实测值）"""
import os
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\23 柱形折线图.xlsx"
RAW = pd.read_excel(XLSX, sheet_name="23 柱形折线图", header=None)
hdr = next(i for i, r in RAW.iterrows() if any(str(v).strip() == "年份" for v in r.values))
df = RAW.iloc[hdr + 1:, 1:4].copy()
df.columns = ["年份", "销售量", "同比"]
df = df.dropna(subset=["年份"]).reset_index(drop=True)
df["销售量"] = pd.to_numeric(df["销售量"], errors="coerce")
df["同比"] = pd.to_numeric(df["同比"], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_BAR = "#0070C0"       # 销售量（柱）
C_LINE = "#E74E69"      # 同比（折线 + 圆点）
C_TITLE = C_SUB = "#FFFFFF"
C_VAL = C_PCT = C_XLAB = "#F2F2F2"
C_LEG = "#B2B2BC"       # 图例文字
C_NOTE = "#D9D9D9"      # 脚注

FS_TITLE, FS_SUB = 34.6, 21.1
FS_VAL, FS_PCT, FS_XLAB, FS_LEG, FS_NOTE = 14.9, 14.2, 14.8, 15.0, 15.4

W, H, DPI = 1088, 756, 100            # 画布 1088 × 756 px

X0, PITCH = 159.5, 153.3              # 第 1 组中心 / 组间距
BARW = 48.5                           # 柱宽
BASE, K = 635.0, 13.44                # 数值轴基线像素行 / 每像素单位数
LINE_C, LINE_K = 319.05, 2.6224       # 折线纵向映射：y = LINE_C - 同比(%) * LINE_K
LINE_LW, DOT_MS, DOT_R = 2.5, 9.4, 6.5    # 折线粗细 / 圆点直径 / 圆点半径(px)
GAP_VAL, GAP_PCT = 22.5, 26.3         # 柱值标签、同比标签的基线偏移

LEG_BAR = (720, 189, 53, 14)          # 图例：销售量色块 (x, y, w, h)
LEG_LN = (861, 916, 888.5, 195.5)     # 图例：同比（线 x0, x1, 圆点 x, 中心 y）


def yp(v):
    """销售量 -> 像素行"""
    return BASE - v / K


def yline(pct):
    """同比（0.27 等）-> 像素行"""
    return LINE_C - pct * 100 * LINE_K


# ---------- 3. 像素坐标系画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                     # y 向下，与图片行号一致
ax.axis("off")

n = len(df)
centers = [X0 + PITCH * i for i in range(n)]

# ---------- 4. 柱形 ----------
for i, r in df.iterrows():
    ax.add_patch(Rectangle((centers[i] - BARW / 2, yp(r["销售量"])), BARW,
                           BASE - yp(r["销售量"]), facecolor=C_BAR,
                           edgecolor="none", zorder=2))

# ---------- 5. 同比折线 + 圆点 ----------
ax.plot(centers, [yline(p) for p in df["同比"]], color=C_LINE, lw=LINE_LW,
        marker="o", ms=DOT_MS, mfc=C_LINE, mec=C_LINE,
        solid_capstyle="round", zorder=4)

# ---------- 6. 柱顶销售量标签（避让折线圆点） ----------
for i, r in df.iterrows():
    y_act, y_dot = yp(r["销售量"]), yline(r["同比"])
    lab_bottom = y_act - GAP_VAL
    if lab_bottom - 15 <= y_dot + DOT_R:      # 常规位置会压到圆点 -> 改放圆点下方
        lab_bottom = y_dot + DOT_R + 20.5
    ax.text(centers[i], lab_bottom + 1, "%d" % r["销售量"], ha="center",
            va="baseline", fontsize=FS_VAL, color=C_VAL, zorder=5)

# ---------- 7. 同比百分比标签 ----------
for i, r in df.iterrows():
    ax.text(centers[i], yline(r["同比"]) - GAP_PCT, "%.0f%%" % (r["同比"] * 100),
            ha="center", va="baseline", fontsize=FS_PCT, color=C_PCT, zorder=5)

# ---------- 8. 年份标签 ----------
for i, r in df.iterrows():
    ax.text(centers[i], 673, str(r["年份"]), ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_XLAB, zorder=5)

# ---------- 9. 图例（右上：柱色块 + 折线圆点） ----------
ax.add_patch(Rectangle((LEG_BAR[0], LEG_BAR[1]), LEG_BAR[2], LEG_BAR[3],
                       facecolor=C_BAR, edgecolor="none", zorder=5))
ax.plot([LEG_LN[0], LEG_LN[1]], [LEG_LN[3]] * 2, color=C_LINE, lw=LINE_LW,
        solid_capstyle="butt", zorder=5)
ax.plot([LEG_LN[2]], [LEG_LN[3]], marker="o", ms=DOT_MS, mfc=C_LINE,
        mec=C_LINE, ls="none", zorder=6)
ax.text(779.5, 201, "销售量", ha="left", va="baseline", fontsize=FS_LEG,
        color=C_LEG, zorder=5)
ax.text(924.5, 201, "同比", ha="left", va="baseline", fontsize=FS_LEG,
        color=C_LEG, zorder=5)

# ---------- 10. 标题 / 副标题 / 脚注 ----------
fig.text(0.06103, 1 - 78 / 756, "近六年销售量及增长率",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(0.06287, 1 - 123.5 / 756, "平台销量近6年持续增长，但近两年增长率有所放缓",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(0.05914, 1 - 724 / 756, "*注：数据来源于公司销售系统",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

# ---------- 11. 输出（保存到指定目录） ----------
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "23.png"), dpi=DPI, facecolor=BG)
plt.show()
