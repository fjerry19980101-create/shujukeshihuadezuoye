# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 09:52:00 2026

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
# 0. 数据读取（读不到 xlsx 时用内置兜底值）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\12 菱形走势图.xlsx"

MONTHS = ["1月", "2月", "3月", "4月", "5月", "6月", "7月", "8月"]
RATES = [0.536, 0.498, 0.527, 0.708, 0.609, 0.496, 0.586, 0.704]
TITLE = "2022年1-8月公司计划完成率"
SUBTITLE = "公司整体完成率55%，4月和8月超过70%，2月和6月较低未过半"
FOOTNOTE = "*注：数据来源于公司销售系统，统计日期截至2022.08.31"

def load_data(path):
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None)
        hdr = None
        for i, row in raw.iterrows():
            cells = [str(v).strip() for v in row.values if pd.notna(v)]
            if any("月份" in c for c in cells) and any("完成率" in c for c in cells):
                hdr = int(i); break
        if hdr is None:
            return None
        body = raw.iloc[hdr + 1:]
        ms, vs = [], []
        for _, row in body.iterrows():
            vals = [v for v in row.values if pd.notna(v)]
            if len(vals) < 2:
                continue
            ms.append(str(vals[0]))
            vs.append(float(vals[1]))
        if len(ms) >= 2:
            return ms, vs
    except Exception as e:
        print("[warn] 读取 xlsx 失败，使用内置数据:", e)
    return None

_d = load_data(XLSX)
if _d:
    MONTHS, RATES = _d
print("months =", MONTHS)
print("rates  =", RATES)

# ============================================================
# 1. 画布 / 颜色 / 字体（全部为参考图实测值）
# ============================================================
W, H = 1088, 758
DPI = 100
BG = "#1A1E43"        # 背景
PINK = "#E74E69"      # 茎与菱形
AXIS = "#717389"      # 底部基线
GRAY_TXT = "#D9D9D9"  # 脚注

FP_REG = fm.FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
FP_BOLD = fm.FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc", weight="bold")

# 字号（渲染标定）
FS_TITLE = 38.4      # 2022年1-8月公司计划完成率（粗体）目标墨迹 696x51
FS_SUB = 23.0        # 副标题 目标 909x31
FS_DATALABEL = 15.0  # 53.60% 目标 69x16
FS_MONTH = 16.5      # 1月 目标 30x20
FS_NOTE = 15.3       # 脚注 目标 541x21

# ============================================================
# 2. 像素坐标系画布（y 向下）
# ============================================================
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ============================================================
# 3. 实测几何
# ============================================================
# 茎/菱形中心 x（8 根，等距 117.7）
X_CENTERS = [146.5, 264.0, 382.0, 499.5, 617.5, 735.0, 852.5, 970.5]
# 菱形顶点 y = 612.76 - 517.17 * rate（最小二乘，8 点残差 < 0.6px）
#   即 0% 在 y=624（基线），100% 在 y=106.8；菱形高 22px，中心 = 顶点 + 10.5
TOP_A, TOP_K = 612.76, 517.17
BASE_Y = 624.0          # 基线（数值 0%）
DIAM_HALF = 10.5        # 菱形半高（22px）
STEM_W = 2.0            # 茎宽 2px
BASE_X0, BASE_X1 = 88.0, 1029.0  # 基线 x 范围

def dia_top(rate):
    return TOP_A - TOP_K * rate

def dia_center(rate):
    return dia_top(rate) + DIAM_HALF

# ============================================================
# 4. 基线（先画，茎压在上面）
# ============================================================
ax.plot([BASE_X0, BASE_X1], [BASE_Y, BASE_Y], color=AXIS, lw=2.0,
        solid_capstyle="butt", zorder=2)

# ============================================================
# 5. 茎 + 空心菱形
# ============================================================
for x, r in zip(X_CENTERS, RATES):
    cy = dia_center(r)
    # 茎：2px，从菱形中心到基线（上半段被菱形盖住）
    ax.plot([x, x], [cy, BASE_Y], color=PINK, lw=STEM_W,
            solid_capstyle="butt", zorder=3)
    # 空心菱形：22px 见方，面填充背景色（挡住茎）、边为粉色 2px
    ms = 2 * DIAM_HALF * 72.0 / DPI          # 22px -> 15.84pt
    mew = 2.0 * 72.0 / DPI                   # 2px -> 1.44pt
    ax.plot(x, cy, "D", markersize=ms,
            markerfacecolor=BG, markeredgecolor=PINK,
            markeredgewidth=mew, zorder=4)

# ============================================================
# 6. 百分比标签（白字，菱形顶点上方 21.4px 处为墨迹底行）
# ============================================================
for x, r in zip(X_CENTERS, RATES):
    ty = dia_top(r)
    baseline = ty - 21.4 + 1.0     # 数字 baseline ≈ 墨迹底行 + 1
    ax.text(x, baseline, "%.2f%%" % (r * 100), fontsize=FS_DATALABEL,
            color="white", ha="center", va="baseline",
            fontproperties=FP_REG, zorder=5)

# ============================================================
# 7. X 轴月份标签（中心 = 茎中心 - 2，baseline = 墨迹底行 - 3）
# ============================================================
for x, m in zip(X_CENTERS, MONTHS):
    ax.text(x - 2.0, 660.0, m, fontsize=FS_MONTH, color="white",
            ha="center", va="baseline", fontproperties=FP_REG)

# ============================================================
# 8. 标题 / 副标题 / 脚注
# ============================================================
# 标题墨迹 x76..771 y44..94，baseline ≈ 底行 - 0.13*51 ≈ 87
ax.text(73.0, 87.0, TITLE, fontsize=FS_TITLE, color="white",
        ha="left", va="baseline", fontproperties=FP_BOLD)
# 副标题墨迹 x74..982 y119..149，baseline ≈ 146
ax.text(73.0, 146.0, SUBTITLE, fontsize=FS_SUB, color="white",
        ha="left", va="baseline", fontproperties=FP_REG)
# 脚注墨迹 x79..619 y700..720，baseline ≈ 717
ax.text(77.0, 717.0, FOOTNOTE, fontsize=FS_NOTE, color=GRAY_TXT,
        ha="left", va="baseline", fontproperties=FP_REG)

# ============================================================
# 9. 输出
# ============================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "12 菱形走势图.png")
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.close(fig)
print("saved ->", OUT)
