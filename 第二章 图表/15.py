# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 09:49:15 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-

import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Ellipse, Rectangle
import pandas as pd

# ============================================================
# 1. 参数区（全部来自参考图实测）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\15 水球图.xlsx"
SHEET = "15 水球图"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "15 水球图.png")

W, H = 1076, 780            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝
WATER = "#0070C0"           # 水色 / 环色（Office 标准蓝，实测）
FONT = "Microsoft YaHei"    # 微软雅黑

# 椭圆几何（中心线半径，实测）
OUTER_CX, OUTER_CY = 509.5, 442.5   # 外环
OUTER_A, OUTER_B = 233.0, 226.0
OUTER_LW = 4.0 / DPI * 72           # 4px -> 2.88pt
BALL_CX, BALL_CY = 511.5, 442.5     # 内球
BALL_A, BALL_B = 214.5, 207.5
BALL_LW = 5.0 / DPI * 72            # 5px -> 3.6pt

# 文字实测参数 -----------------------------------------------
TITLE = "2022年上半年目标完成率"
# 原图副标题设置了字符间距（约 0.65px/字，19 字累计 +12px），
# matplotlib 无原生字距 -> 按参考图实测分三段锚定段首
SUBTITLE = "截至6月30日销售目标总体完成率达到65%"
SUB_SEGMENTS = [
    (75.0, "截至6月30日"),                 # 墨迹 76-283
    (290.0, "销售目标总体完成率达到"),     # 墨迹 291-697
    (703.0, "65%"),                        # 墨迹 705-775
]
NOTE = "*注：数据来源于公司销售系统"

FS_TITLE, FS_SUB, FS_PCT, FS_NOTE = 38.5, 26.5, 65.0, 15.4

# (x_px, baseline_px)；baseline 由墨迹底行反推：
#   粗体标题 ≈ 底行-5.5；CJK 副标题 ≈ 底行-4.5；脚注 ≈ 底行-2；数字 ≈ 底行+0.5
X_TITLE, Y_TITLE_BASE = 74.0, 90 - 5.5     # 标题墨迹 78-685 / 40-90
Y_SUB_BASE = 146 - 4.5                     # 副标题墨迹 / 111-146
X_NOTE, Y_NOTE_BASE = 62.0, 749 - 2.0      # 脚注墨迹 64-346 / 730-749
X_PCT, Y_PCT_BASE = 428.0, 461.5           # 65% 墨迹 434-608 / 391-462


# ============================================================
# 2. 读取数据
# ============================================================
def load_pct():
    """从 xlsx 读完成率（找含「完成率」的表头，其下一行即数值）。"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    for i, r in raw.iterrows():
        for j, v in enumerate(r.values):
            if str(v).strip() == "完成率":
                return float(raw.iloc[i + 1, j])
    raise ValueError("未找到「完成率」数据")


# ============================================================
# 3. 绘图
# ============================================================
def draw():
    pct = load_pct()

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    # ---- 水域：矩形 clip 到内球椭圆（水面 = 球底向上 pct x 球高）----
    water_y = BALL_CY + BALL_B - pct * 2 * BALL_B   # 数据驱动；实测 379，差 1px
    ball = Ellipse((BALL_CX, BALL_CY), 2 * BALL_A, 2 * BALL_B,
                   fill=False, lw=0, zorder=2)       # 仅作为 clip 路径
    ax.add_patch(ball)
    water = Rectangle((BALL_CX - BALL_A, water_y),
                      2 * BALL_A, BALL_CY + BALL_B - water_y + 10,
                      facecolor=WATER, edgecolor="none", zorder=2)
    ax.add_patch(water)
    water.set_clip_path(ball)

    # ---- 内球描边（水面以上可见的弧线；水下与水同色融合）----
    ax.add_patch(Ellipse((BALL_CX, BALL_CY), 2 * BALL_A, 2 * BALL_B,
                         fill=False, edgecolor=WATER, lw=BALL_LW, zorder=3))

    # ---- 外椭圆环 ----
    ax.add_patch(Ellipse((OUTER_CX, OUTER_CY), 2 * OUTER_A, 2 * OUTER_B,
                         fill=False, edgecolor=WATER, lw=OUTER_LW, zorder=3))

    # ---- 百分比大字（白，坐落水面处）----
    fig.text(X_PCT / W, (H - Y_PCT_BASE) / H, "%.0f%%" % (pct * 100),
             fontsize=FS_PCT, color="#FFFFFF", fontfamily=FONT,
             ha="left", va="baseline", zorder=4)

    # ---- 标题 / 副标题（三段锚定）/ 脚注 ----
    # fig.text 的 y 从画布底部起算，须用 (H - baseline_px)/H 换算
    fig.text(X_TITLE / W, (H - Y_TITLE_BASE) / H, TITLE, fontsize=FS_TITLE,
             color="#FFFFFF", fontweight="bold", fontfamily=FONT,
             ha="left", va="baseline")
    for x_seg, s in SUB_SEGMENTS:
        fig.text(x_seg / W, (H - Y_SUB_BASE) / H, s, fontsize=FS_SUB,
                 color="#FFFFFF", fontfamily=FONT, ha="left", va="baseline")
    fig.text(X_NOTE / W, (H - Y_NOTE_BASE) / H, NOTE, fontsize=FS_NOTE,
             color="#D9D9D9", fontfamily=FONT, ha="left", va="baseline")

    os.makedirs(OUT_DIR, exist_ok=True)
    fig.savefig(OUT_PNG, dpi=DPI, facecolor=BG)
    plt.close(fig)
    print("saved:", OUT_PNG)


if __name__ == "__main__":
    draw()
