# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:32:21 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""20 南丁格尔图（PPT 版）复现（尺寸/角度/配色/字号均为参考图像素级实测值）"""
import os
import math
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\20 南丁格尔（PPT）.xlsx"
RAW = pd.read_excel(XLSX, sheet_name="20 南丁格尔（PPT）", header=None)
hdr = next(i for i, r in RAW.iterrows() if any(str(v).strip() == "部门" for v in r.values))
df = RAW.iloc[hdr + 1:, 1:3].copy()
df.columns = ["部门", "占比"]
df = df.dropna(subset=["部门"]).reset_index(drop=True)
df["部门"] = df["部门"].astype(str).str.strip()
df["占比"] = pd.to_numeric(df["占比"], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_TEXT = "#F2F2F2"      # 标题 / 副标题 / 脚注

# 花瓣颜色（按数据行顺序：销售部 / 采购部 / 工程部 / 财务部 / 行政部 / 人力部）
PETAL_COLORS = ["#E74E69", "#F5C353", "#11A7AD", "#37A2DA", "#0070C0", "#7030A0"]
# 每片花瓣的引出线：(圆点, 拐点, 末端)  —— 全部实测像素坐标
LEADERS = [
    ((818, 322), (855, 274), (1013, 274)),
    ((756, 628), (782, 675), (938, 675)),
    ((371, 636), (332, 686), (177, 686)),
    ((443, 368), (410, 324), (208, 324)),
    ((342, 456), (310, 503), (150, 503)),
    ((518, 329), (489, 283), (333, 283)),
]
FS_TITLE, FS_SUB, FS_NOTE = 34.6, 23.1, 15.4

W, H, DPI = 1090, 792, 100             # 画布 1090 × 792 px

CX, CY = 531.0, 472.0                  # 花瓣公共圆心
A0 = 0.0                               # 第 1 片花瓣从正上方开始，顺时针
R_SCALE = 10.56                        # 半径比例：r = R_SCALE × 占比(%)
DOT_R = 5.0                            # 引出圆点半径
LINE_LW = 2.2                          # 引出线粗细（≈3px）


def pol(r, adeg):
    """极坐标（正上方 0°、顺时针为正）-> 画布像素坐标"""
    a = math.radians(adeg)
    return CX + r * math.sin(a), CY - r * math.cos(a)


total = float(df["占比"].sum())

# ---------- 3. 画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                      # y 向下，与图片行号一致
ax.axis("off")

# ---------- 4. 花瓣（角度按占比分 360°，半径按占比线性） ----------
ang = A0
for i, row in df.iterrows():
    v = float(row["占比"])
    span = 360.0 * v / total
    r = R_SCALE * v * 100
    color = PETAL_COLORS[i % len(PETAL_COLORS)]
    ax.add_patch(Wedge((CX, CY), r, ang - 90, ang + span - 90,
                       facecolor=color, edgecolor="none", zorder=2))
    # 引出线：圆点 -> 拐点 -> 末端
    dot, bend, end = LEADERS[i % len(LEADERS)]
    ax.plot([dot[0], bend[0], end[0]], [dot[1], bend[1], end[1]], color=color,
            lw=LINE_LW, solid_capstyle="butt", solid_joinstyle="miter", zorder=4)
    ax.add_patch(Circle(dot, DOT_R, facecolor=color, edgecolor="none", zorder=5))
    ang += span

# ---------- 5. 标题 / 副标题 / 脚注 ----------
fig.text(0.05817, 1 - 84.5 / 792, "2021年各部门人数分布",
         fontsize=FS_TITLE, fontweight="bold", color=C_TEXT,
         ha="left", va="baseline")
fig.text(0.05798, 1 - 141 / 792, "公司总人数1664，销售部人数最多451，占比29.2%",
         fontsize=FS_SUB, color=C_TEXT, ha="left", va="baseline")
fig.text(0.04894, 1 - 756 / 792, "*注：数据来源于公司人力资源系统，统计日期截至2022.01.01",
         fontsize=FS_NOTE, color=C_TEXT, ha="left", va="baseline")

# ---------- 6. 输出（保存到指定目录） ----------
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "20_南丁格尔图_PPT.png"), dpi=DPI, facecolor=BG)
plt.show()
