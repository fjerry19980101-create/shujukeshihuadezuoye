# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 08:23:17 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-

import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.lines import Line2D

XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\30 对比滑珠图.xlsx"
SHEET = "30 对比滑珠图"

# ===================== 1. 读数据 =====================
raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
hdr = next(i for i, r in raw.iterrows()
           if any(str(v).strip() == "区域" for v in r.values))
df = raw.iloc[hdr + 1:, 1:6].copy()
df.columns = ["区域", "完成率2022", "完成率2021", "未完成占比", "辅助"]
df = df.dropna(subset=["区域"])
df["区域"] = df["区域"].astype(str).str.strip()
for c in ["完成率2022", "完成率2021"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df = (df.dropna(subset=["完成率2022", "完成率2021"])
        .sort_values("完成率2022", ascending=False)   # 华南在最上
        .reset_index(drop=True))

# ===================== 2. 样式常量（按原图取样） =====================
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG       = "#1A1E43"   # 画布底色（深藏蓝）
BLUE     = "#0070C0"   # 2022：进度条 + 圆点
BASE_BAR = "#717389"   # 滑道底条
GRAY_DOT = "#A6A6A6"   # 2021：圆点
C_TITLE  = "#FFFFFF"   # 标题
C_SUB    = "#F2F2F2"   # 副标题
C_LABEL  = "#C8CAD6"   # 左侧区域名 / 图例文字
C_NOTE   = "#B3B4C0"   # 脚注

BAR_LW  = 13           # 条带粗细 ≈18 px @dpi=100
DOT_MS  = 18           # 圆点直径 ≈25 px（含白描边）
LAB_DY  = 0.28         # 数值标签相对条带中心上移量（数据坐标）

# ===================== 3. 画布与坐标系 =====================
fig = plt.figure(figsize=(10.88, 7.56), dpi=100, facecolor=BG)
# 条带区精确对齐原图：[163→864] px 水平、[205→676] px 垂直
ax = fig.add_axes([163 / 1088, (756 - 676) / 756,
                   (864 - 163) / 1088, (676 - 205) / 756])
ax.set_facecolor(BG)

n = len(df)
ax.set_xlim(-0.013, 1.013)      # 留出圆头条带的端部
ax.set_ylim(-0.5, n - 0.5)
ax.set_xticks([])
for s in ax.spines.values():
    s.set_visible(False)

# ===================== 4. 逐行绘制滑珠 =====================
for i, row in df.iterrows():
    y = n - 1 - i                       # 第一行（华南）在最上方
    v22, v21 = row["完成率2022"], row["完成率2021"]

    # 4.1 滑道：整条 0%~100%
    ax.plot([0, 1], [y, y], color=BASE_BAR, lw=BAR_LW,
            solid_capstyle="round", zorder=2)
    # 4.2 2022 进度条：0 → 完成率
    ax.plot([0, v22], [y, y], color=BLUE, lw=BAR_LW,
            solid_capstyle="round", zorder=3)
    # 4.3 2021 灰点
    ax.plot([v21], [y], marker="o", ms=DOT_MS, mfc=GRAY_DOT, mec="white",
            mew=2, ls="none", zorder=4)
    # 4.4 2022 蓝点（压在灰点之上）
    ax.plot([v22], [y], marker="o", ms=DOT_MS, mfc=BLUE, mec="white",
            mew=2, ls="none", zorder=5)
    # 4.5 数值标签（只在蓝点上方）
    ax.text(v22, y + LAB_DY, f"{v22:.0%}", ha="center", va="bottom",
            fontsize=15, fontweight="bold", color="#FFFFFF", zorder=6)

# 左侧区域名
ax.set_yticks(range(n))
ax.set_yticklabels(list(df["区域"])[::-1], fontsize=15, color=C_LABEL)
ax.tick_params(axis="y", length=0, pad=12)

# ===================== 5. 图例 =====================
handles = [
    Line2D([], [], marker="o", ls="none", ms=14, mfc=BLUE,
           mec="white", mew=1.8, label="2022完成率"),
    Line2D([], [], marker="o", ls="none", ms=14, mfc=GRAY_DOT,
           mec="white", mew=1.8, label="2021完成率"),
]
leg = ax.legend(handles=handles, loc="center left",
                bbox_to_anchor=(-0.095, 1.009), ncol=2, frameon=False,
                handlelength=1.0, handletextpad=0.5, columnspacing=1.6,
                borderpad=0, borderaxespad=0, fontsize=14)
for t in leg.get_texts():
    t.set_color(C_LABEL)

# ===================== 6. 标题 / 副标题 / 脚注 =====================
fig.text(0.0625, 0.9418,
         "2022年上半年销量目标达成率同比去年情况",
         fontsize=26, fontweight="bold", color=C_TITLE,
         ha="left", va="top")

fig.text(0.0625, 0.8720,
         "华南完成率最高达到86%，华东最低35%，其中华南和华东不及2021年",
         fontsize=17, color=C_SUB, ha="left", va="top")

fig.text(0.0625, 0.0622,
         "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=13, color=C_NOTE, ha="left", va="center")

# ===================== 7. 输出 =====================
fig.savefig("对比滑珠图.png", dpi=100, facecolor=BG)
plt.show()
