# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 13:14:07 2026

@author: fjwsm
"""

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import Rectangle
import numpy as np
import pandas as pd

# ============================================================
# 0. 数据读取
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\1 渐变柱形图.xlsx"

# 兜底数据（与 xlsx 一致）
REGIONS = ["华北", "华南", "东北", "西北", "西南", "华东"]
SALES = [2354, 1902, 3524, 2698, 2896, 2563]
TITLE = "3月各区域销量分布"
SUBTITLE = "东北销量最多占比总销量的22%，华南销量最低"
FOOTNOTE = "*注：数据来源于公司销售系统，统计日期截至2022.03.31"


def load_data(path):
    """从 xlsx 读取区域与销售量两列（表头行含 '区域' '销售'）。"""
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None)
        hdr, rcol, vcol = None, None, None
        for i, row in raw.iterrows():
            cells = [(j, str(v).strip()) for j, v in enumerate(row.values) if pd.notna(v)]
            tags = [c for _, c in cells]
            if any("区域" in c for c in tags) and any("销售" in c for c in tags):
                hdr = int(i)
                rcol = next(j for j, c in cells if "区域" in c)
                vcol = next(j for j, c in cells if "销售" in c)
                break
        if hdr is None:
            raise ValueError("未找到表头（区域/销售量）")
        df = raw.iloc[hdr + 1:, [rcol, vcol]].dropna()
        rs = [str(m).strip() for m in df[rcol].tolist()]
        vs = [int(float(v)) for v in df[vcol].tolist()]
        return rs, vs
    except Exception as e:
        print("[warn] 读取 xlsx 失败，使用内置数据:", e)
    return REGIONS, SALES


REGIONS, SALES = load_data(XLSX)

# ============================================================
# 1. 画布与颜色
# ============================================================
W, H = 1088, 794
DPI = 100
BG = "#1A1E43"          # 背景
GRID = "#404361"        # 网格虚线
WHITE = "#FFFFFF"       # 标题 / 副标题 / 柱值标签
SOFT = "#F2F2F2"        # Y 轴刻度
CAT = "#BBBCC5"         # 类目标签 (187,188,197)
GREY = "#D9D9D9"        # 脚注
C_TOP = (0, 112, 192)   # 柱顶色
C_BOT = (0, 176, 240)   # 柱底色

# 字体
FP_REG = fm.FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
FP_BOLD = fm.FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc")

# 字号（pt，宽度对标实测标定）
FS_TITLE = 38.4     # 标题（粗体），宽度 454px 精确命中
FS_SUB = 26.87      # 副标题，宽度 785px 标定
FS_Y = 13.2         # Y 轴刻度，宽度 42px 精确命中
FS_VAL = 15.0       # 柱值标签，宽度 47px 标定
FS_CAT = 16.5       # 类目标签，宽度 45px 精确命中
FS_NOTE = 15.36     # 脚注，宽度 541px 标定

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
# 3. 映射 / 网格虚线（在柱后）/ Y 轴刻度
# ============================================================
A0, B0 = 671.0, -0.10425


def yp(v):
    return A0 + B0 * v


GRID_Y = [254, 358, 462, 567, 671]      # 4000 / 3000 / 2000 / 1000 / 0
for gy in GRID_Y:
    ax.plot([157, 979], [gy, gy], color=GRID, lw=1.44,            # 2px
            linestyle=(4.0, (8.0, 3.0)),                          # 首段 8px 通，其后 16/6
            dash_capstyle="butt", antialiased=False, zorder=1)

# Y 轴刻度标签（右对齐 x=134.5，墨迹垂直居中于网格线 -> baseline = 线心 + 7.6）
for val, gy in zip([4000, 3000, 2000, 1000, 0], GRID_Y):
    ax.text(134.5, gy + 7.6, str(val), fontsize=FS_Y, color=SOFT,
            ha="right", va="baseline", fontproperties=FP_REG, zorder=3)

# ============================================================
# 4. 渐变柱（逐像素行切片，二次 ease-out）
# ============================================================
# 每根柱：(墨迹起始列, 像素宽) —— 与参考图逐列实测一致
BAR_X = [(204, 44), (342, 44), (480, 43), (618, 43), (756, 43), (894, 43)]
BAR_BOT = 670.0     # 柱底边（墨迹最底行 669，紧贴 0 网格线上沿）

for (x0, bw), v in zip(BAR_X, SALES):
    top = yp(v)                        # 真实柱顶（分数 px）
    r0 = int(np.floor(top))            # 墨迹顶行
    for r in range(r0, int(BAR_BOT)):
        t = (r + 0.5 - top) / (BAR_BOT - top)
        f = 1.0 - (1.0 - t) ** 2       # 二次 ease-out
        col = tuple(round(a + (b - a) * f) / 255.0 for a, b in zip(C_TOP, C_BOT))
        ax.add_patch(Rectangle((x0, r), bw, 1.0, facecolor=col,
                               edgecolor="none", antialiased=False, zorder=2))

# ============================================================
# 5. 柱值标签（白字，baseline = 柱顶 - 21.6）与类目标签
# ============================================================
for (x0, bw), v, name in zip(BAR_X, SALES, REGIONS):
    cx = x0 + (bw - 1) / 2.0           # 墨迹中心
    ax.text(cx, yp(v) - 21.6, str(v), fontsize=FS_VAL, color=WHITE,
            ha="center", va="baseline", fontproperties=FP_REG, zorder=3)
    ax.text(cx, 707.0, name, fontsize=FS_CAT, color=CAT,
            ha="center", va="baseline", fontproperties=FP_REG, zorder=3)

# ============================================================
# 6. 标题 / 副标题 / 脚注
# ============================================================
# 标题 bbox x89..542 y63..109，baseline = 墨迹底行 - 0.13*高
ax.text(89.0, 102.9, TITLE, fontsize=FS_TITLE, color=WHITE,
        ha="left", va="baseline", fontproperties=FP_BOLD)
# 副标题 bbox x86..870 y140..175，baseline = 墨迹底行 - 3
ax.text(86.0, 172.0, SUBTITLE, fontsize=FS_SUB, color=WHITE,
        ha="left", va="baseline", fontproperties=FP_REG)
# 脚注 bbox x66..606 y740..760，baseline = 757
ax.text(66.0, 757.0, FOOTNOTE, fontsize=FS_NOTE, color=GREY,
        ha="left", va="baseline", fontproperties=FP_REG)

# ============================================================
# 7. 输出
# ============================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "1 渐变柱形图.png")
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.close(fig)
print("saved ->", OUT)
