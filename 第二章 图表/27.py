# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 08:55:49 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""27 簇状柱形折线图 复现（尺寸/坐标/配色/字号均为参考图像素级实测值）"""
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\27 簇状柱形折线图.xlsx"
raw = pd.read_excel(XLSX, sheet_name="27 簇状柱形折线图", header=None)
hdr = next(i for i, r in raw.iterrows() if any(str(v).strip() == "区域" for v in r.values))
df = raw.iloc[hdr + 1:, 1:5].copy()
df.columns = ["区域", "销量2022", "销量2021", "同比去年"]
df = df.dropna(subset=["区域"]).reset_index(drop=True)
for c in ["销量2022", "销量2021", "同比去年"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_2022 = "#0070C0"      # 2022 柱（蓝）
C_2021 = "#E74E69"      # 2021 柱（粉）
C_LINE = "#FFC000"      # 同比折线 + 圆点（金）
C_AXIS = "#454866"      # 底部轴线
C_TEXT = "#F2F2F2"      # 标题/副标题/各类标签
C_NOTE = "#D9D9D9"      # 脚注

FS_TITLE, FS_SUB = 34.5, 23.1
FS_VAL, FS_PCT, FS_XLAB, FS_LEG, FS_NOTE = 14.3, 14.3, 16.3, 14.8, 15.4

W, H, DPI = 1088, 792, 100         # 画布 1088 × 792 px

X0, PITCH = 165.25, 153.1          # 第 1 组中心 / 组间距
BARW, GAP = 34.5, 9.5              # 柱宽 / 组内两柱间隙
BASE, YMIN, K = 660.0, 100.0, 10.915   # 基线像素行 / 数值轴下限 / 单位每像素
GAP_VAL, GAP_PCT = 19.8, 26.3      # 数值标签、百分比标签的基线偏移
LINE_C, LINE_K = 235.5, 3.3333     # 折线纵向映射：y = LINE_C + (22 - pct)*100*LINE_K
AXIS_Y, AXIS_X0, AXIS_X1 = 669.5, 89, 1006
LEGEND = [(892, 525, 100, 37, C_2022, "2022"),
          (903, 580, 92, 37, C_2021, "2021")]


def yp(v):
    """销量 -> 像素行"""
    return BASE - (v - YMIN) / K


def yln(pct):
    """同比（0.16 等）-> 像素行"""
    return LINE_C + (0.22 - pct) * 100 * LINE_K


# ---------- 3. 像素坐标系画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                  # y 向下，与图片行号一致
ax.axis("off")

n = len(df)
centers = [X0 + PITCH * i for i in range(n)]

# ---------- 4. 簇状柱形 ----------
for i, r in df.iterrows():
    c = centers[i]
    for x0, val, color in ((c - BARW - GAP / 2, r["销量2022"], C_2022),
                           (c + GAP / 2, r["销量2021"], C_2021)):
        ax.add_patch(Rectangle((x0, yp(val)), BARW, BASE - yp(val),
                               facecolor=color, edgecolor="none", zorder=2))

# ---------- 5. 同比折线 + 圆点 ----------
ax.plot(centers, [yln(p) for p in df["同比去年"]], color=C_LINE, lw=2.5,
        marker="o", ms=9.4, mfc=C_LINE, mec=C_LINE, solid_capstyle="round",
        zorder=5)

# ---------- 6. 底部轴线 ----------
ax.plot([AXIS_X0, AXIS_X1], [AXIS_Y, AXIS_Y], color=C_AXIS, lw=1.44,
        solid_capstyle="butt", zorder=1)

# ---------- 7. 数值标签 ----------
for i, r in df.iterrows():
    c = centers[i]
    for x, val in ((c - BARW - GAP / 2 + BARW / 2, r["销量2022"]),
                   (c + GAP / 2 + BARW / 2, r["销量2021"])):
        ax.text(x, yp(val) - GAP_VAL, "%d" % val, ha="center", va="baseline",
                fontsize=FS_VAL, color=C_TEXT, zorder=6)
    ax.text(c, yln(r["同比去年"]) - GAP_PCT, "%.0f%%" % (r["同比去年"] * 100),
            ha="center", va="baseline", fontsize=FS_PCT, color=C_TEXT, zorder=6)

# ---------- 8. 区域标签 ----------
for i, r in df.iterrows():
    ax.text(centers[i], 706, str(r["区域"]), ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_TEXT, zorder=6)

# ---------- 9. 图例（深底 + 彩色描边小框，压在柱子上） ----------
for x0, y0, w, h, ec, txt in LEGEND:
    ax.add_patch(Rectangle((x0, y0), w, h, facecolor=BG, edgecolor=ec,
                           lw=1.44, zorder=7))
    ax.text(x0 + w / 2, y0 + h / 2 + 2.5, txt, ha="center", va="center",
            fontsize=FS_LEG, color=C_TEXT, zorder=8)

# ---------- 10. 标题 / 副标题 / 脚注 ----------
fig.text(0.05184, 1 - 78 / 792, "上半年各月商品销量同比去年情况",
         fontsize=FS_TITLE, fontweight="bold", color=C_TEXT, ha="left", va="baseline")
fig.text(0.05349, 1 - 125.5 / 792, "2022年相比于2021年销量都有提升，半年整体提升17%",
         fontsize=FS_SUB, color=C_TEXT, ha="left", va="baseline")
fig.text(0.04444, 1 - 756 / 792, "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

# ---------- 11. 输出 ----------
fig.savefig("27.png", dpi=DPI, facecolor=BG)
plt.show()
