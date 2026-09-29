# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 09:31:25 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import numpy as np
import pandas as pd

# ============================================================
# 0. 数据读取
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\13 对比折线图.xlsx"

# 兜底数据
MONTHS = ["1月", "2月", "3月", "4月", "5月", "6月"]
V21 = [1686, 1345, 1934, 1658, 1865, 1936]
V22 = [1385, 1846, 1654, 1936, 2564, 2236]
TITLE = "2022年上半年各月同比去年销量"
SUBTITLE = "上半年同比去年增长明显，5月份同比增长最多，增长近40%"
FOOTNOTE = "*注：数据来源于公司销售系统，统计日期截至2022.06.30"

def load_data(path):
    """从 xlsx 读取月份、2021年、2022年三列数据与文案。"""
    try:
        xl = pd.ExcelFile(path)
        for sheet in xl.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            # 找表头
            hdr = None
            for i, row in raw.iterrows():
                cells = [str(v).strip() for v in row.values if pd.notna(v)]
                if any("2021" in c for c in cells) and any("2022" in c for c in cells):
                    hdr = int(i); break
            if hdr is None:
                continue
            df = raw.iloc[hdr+1:, :].copy()
            df.columns = [str(c) if pd.notna(c) else f"c{j}" for j, c in enumerate(raw.iloc[hdr, :])]
            # 定位月份、2021、2022列
            mcol = next((c for c in df.columns if "月" in str(c) or "月份" in str(c)), None)
            y21 = next((c for c in df.columns if "2021" in str(c)), None)
            y22 = next((c for c in df.columns if "2022" in str(c)), None)
            if mcol and y21 and y22:
                ms = df[mcol].astype(str).tolist()[:6]
                v21 = [int(v) if pd.notna(v) else 0 for v in df[y21].tolist()[:6]]
                v22 = [int(v) if pd.notna(v) else 0 for v in df[y22].tolist()[:6]]
                return ms, v21, v22
    except Exception as e:
        print("[warn] 读取 xlsx 失败，使用内置数据:", e)
    return MONTHS, V21, V22

_months, _v21, _v22 = load_data(XLSX)
MONTHS = _months
V21 = _v21
V22 = _v22

# ============================================================
# 1. 画布与颜色
# ============================================================
W, H = 1088, 792
DPI = 100
BG = "#1A1E43"            # 背景
GRID = "#454866"          # 虚线网格
PINK = "#E74E69"          # 2021年 折线/标记
BLUE = "#0070C0"          # 2022年 折线/标记

# 字体
FP_REG = fm.FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
FP_BOLD = fm.FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc", weight="bold")

# 字号（实测标定）
FS_TITLE = 38.4           # 标题 2022年上半年各月同比去年销量
FS_SUB = 23.0             # 副标题
FS_LABEL = 15.4           # Y轴刻度
FS_XLABEL = 16.5          # X轴月份
FS_DATALABEL = 15.0       # 数据标签
FS_LEGEND = 15.3          # 图例
FS_NOTE = 15.3            # 脚注

# ============================================================
# 2. 坐标系：像素坐标，y 向下
# ============================================================
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ============================================================
# 3. 绘图区实测参数
# ============================================================
# 网格线：x 142..1026，7 条水平虚线
GRID_LEFT = 142.0
GRID_RIGHT = 1026.0
GRID_Y = {
    3000: 284.0,
    2500: 348.0,
    2000: 413.0,
    1500: 478.0,
    1000: 542.0,
    500: 607.0,
    0: 672.0,
}
# 数值 -> y 像素（线性回归：y = 672.25 - 0.129357 * v）
A0, B0 = 672.25, -0.129357

def yp(v):
    return A0 + B0 * v

# X 轴月份位置（标签中心）
X_CENTER = [215.0, 362.5, 510.0, 657.5, 805.0, 952.5]

# ============================================================
# 4. 画网格线与 Y 轴刻度
# ============================================================
# 虚线样式：实测 16 px on / 6 px off
for val, y in GRID_Y.items():
    ax.plot([GRID_LEFT, GRID_RIGHT], [y, y], color=GRID, lw=1.0,
            linestyle=(0, (16.0 / 1.0, 6.0 / 1.0)), solid_capstyle="butt")
    # 左侧 Y 轴标签
    label = str(int(val))
    # 右对齐到 x≈124，baseline ≈ 墨迹底行 - 3
    ax.text(124.0, y - 3.0, label, fontsize=FS_LABEL, color="white",
            ha="right", va="baseline", fontproperties=FP_REG)

# ============================================================
# 5. 画折线与数据点
# ============================================================
# 数据点圆盘直径 11px => ms≈11（像素下 marker 直径）
# 线宽 4px
X = np.array(X_CENTER, dtype=float)
Y21 = np.array([yp(v) for v in V21])
Y22 = np.array([yp(v) for v in V22])

for (x, y, col, vals) in [(X, Y21, PINK, V21), (X, Y22, BLUE, V22)]:
    ax.plot(x, y, color=col, lw=4.0, solid_capstyle="round",
            solid_joinstyle="round", zorder=3)
    ax.plot(x, y, "o", color=col, markersize=11.0, markeredgecolor=col,
            markeredgewidth=0, zorder=4)

# ============================================================
# 6. 数据标签（白字，放在本列较高点上方 / 较低点下方）
# ============================================================
# 规则：逐列比较 2021/2022，较高点的标签在点上方，较低点的标签在点下方
# 上方标签：标签底行 ≈ 点 y - 28.5；下方标签：标签顶行 ≈ 点 y + 30.5
# 数字 baseline ≈ 墨迹底行 + 1（纯数字无下伸部）
for k, (x, v21, v22) in enumerate(zip(X, V21, V22)):
    y21 = yp(v21); y22 = yp(v22)
    # 2021
    if v21 >= v22:
        # 标签在上方
        label_y = y21 - 28.5 - 14.0  # 标签中心偏上：墨迹高 15，中心≈点上方 ~35
        va = "bottom"
    else:
        label_y = y21 + 30.5 + 7.0
        va = "top"
    ax.text(x, label_y, str(v21), fontsize=FS_DATALABEL, color="white",
            ha="center", va=va, fontproperties=FP_REG, zorder=5)
    # 2022
    if v22 >= v21:
        label_y = y22 - 28.5 - 14.0
        va = "bottom"
    else:
        label_y = y22 + 30.5 + 7.0
        va = "top"
    ax.text(x, label_y, str(v22), fontsize=FS_DATALABEL, color="white",
            ha="center", va=va, fontproperties=FP_REG, zorder=5)

# ============================================================
# 7. X 轴月份标签
# ============================================================
for x, m in zip(X, MONTHS):
    ax.text(x, 710.0, m, fontsize=FS_XLABEL, color="white",
            ha="center", va="baseline", fontproperties=FP_REG)

# ============================================================
# 8. 图例（—●— 2021年   —●— 2022年）
# ============================================================
LEGEND_Y = 236.0      # 图例中心线
LEGEND_LINE_L = 28.0  # 单侧线长
MARKER_SIZE = 11.0

# 2021 图例（位于 x713..769 区域，文字 x≈778..846）
legend_2021_x = 740.0
ax.plot([legend_2021_x - LEGEND_LINE_L, legend_2021_x + LEGEND_LINE_L],
        [LEGEND_Y, LEGEND_Y], color=PINK, lw=4.0, solid_capstyle="butt", zorder=3)
ax.plot(legend_2021_x, LEGEND_Y, "o", color=PINK, markersize=MARKER_SIZE,
        markeredgecolor=PINK, markeredgewidth=0, zorder=4)
ax.text(legend_2021_x + LEGEND_LINE_L + 9.0, LEGEND_Y, "2021年",
        fontsize=FS_LEGEND, color="white", ha="left", va="center",
        fontproperties=FP_REG)

# 2022 图例（文字 x≈984..1080，推测左侧色块中心 x≈954）
legend_2022_x = 954.0
ax.plot([legend_2022_x - LEGEND_LINE_L, legend_2022_x + LEGEND_LINE_L],
        [LEGEND_Y, LEGEND_Y], color=BLUE, lw=4.0, solid_capstyle="butt", zorder=3)
ax.plot(legend_2022_x, LEGEND_Y, "o", color=BLUE, markersize=MARKER_SIZE,
        markeredgecolor=BLUE, markeredgewidth=0, zorder=4)
ax.text(legend_2022_x + LEGEND_LINE_L + 9.0, LEGEND_Y, "2022年",
        fontsize=FS_LEGEND, color="white", ha="left", va="center",
        fontproperties=FP_REG)

# ============================================================
# 9. 标题 / 副标题 / 脚注
# ============================================================
# 标题 bbox x73..839 y42..92，baseline ≈ 墨迹底行 - 0.13*高
ax.text(73.0, 87.0, TITLE, fontsize=FS_TITLE, color="white",
        ha="left", va="baseline", fontproperties=FP_BOLD)
# 副标题 bbox x71..920 y117..147，baseline ≈ 146
ax.text(71.0, 145.0, SUBTITLE, fontsize=FS_SUB, color="white",
        ha="left", va="baseline", fontproperties=FP_REG)
# 脚注 bbox x71..613 y736..756，baseline ≈ 751
ax.text(71.0, 752.0, FOOTNOTE, fontsize=FS_NOTE, color="#D9D9D9",
        ha="left", va="baseline", fontproperties=FP_REG)

# ============================================================
# 10. 输出
# ============================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "13 对比折线图.png")
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.close(fig)
print("saved ->", OUT)
