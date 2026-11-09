# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:19:26 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""22 仪表盘图 复现（尺寸/角度/配色/字号均为参考图像素级实测值）"""
import os
import math
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\22 仪表盘图.xlsx"
RAW = pd.read_excel(XLSX, sheet_name="22 仪表盘图", header=None)

# B 列 = 刻度值(50..150)；C 列 = 每段对应的角度值(27，最后一段 90)；H 列 = 指针数值
_col_b = pd.to_numeric(RAW.iloc[:, 1], errors="coerce").dropna()
_col_c = pd.to_numeric(RAW.iloc[:, 2], errors="coerce").dropna()
_col_h = pd.to_numeric(RAW.iloc[:, 7], errors="coerce").dropna()
VALS = [int(v) for v in _col_b if 50 <= v <= 150]
SIZES = [float(v) for v in _col_c if v > 0]
if len(SIZES) != len(VALS):                       # 兜底：按 27/90 的规律补
    SIZES = [90.0 if v == 150 else 27.0 for v in VALS]
TICKS = list(zip(VALS, SIZES))
NEEDLE_VAL = float(_col_h.iloc[0]) if len(_col_h) else 76.0

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色（也是圆环内的暗色带 / 分隔线颜色）
PANEL = "#FFFFFF"       # 中央白色底板
C_OUT = "#0070C0"       # 外圈（固定蓝）
C_TEAL = "#087E8B"      # 低区段
C_GOLD = "#FFC54D"      # 中区段
C_RED = "#F94C66"       # 高区段
C_NEEDLE = "#F2F2F2"    # 指针
C_TICK = "#F2F2F2"      # 刻度数字
C_TITLE = "#FFFFFF"

FS_TITLE, FS_TICK = 34.6, 14.7

W, H, DPI = 1088, 792, 100            # 画布 1088 × 792 px

CX, CY = 563.0, 484.5                 # 圆心
PANEL_XY, PANEL_WH = (305, 226), (517, 518)
R_HOLE = 181.5                        # 内孔半径
R_SEG_IN, R_SEG_OUT = 207.5, 231.5    # 彩色区段带
R_OUT_IN, R_OUT_OUT = 234.5, 257.5    # 外圈蓝带
R_LABEL = 195.5                       # 刻度数字所在半径
A_START, A_STEP = 225.0, 2.7          # 刻度 50 在 225°，每 1 个单位 2.7°（顺时针）
NEEDLE_LEN = 178.0                    # 指针长度


def seg_color(v):
    """区段配色：按段起始刻度值分档（150 段留空）"""
    if v < 80:
        return C_TEAL
    if v < 120:
        return C_GOLD
    if v < 150:
        return C_RED
    return None


def pol(r, adeg):
    """极坐标(角度为「正上方起顺时针」) -> 画布像素坐标"""
    a = math.radians(adeg)
    return CX + r * math.sin(a), CY - r * math.cos(a)


# ---------- 3. 画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                     # y 向下，与图片行号一致
ax.axis("off")

# ---------- 4. 白色底板 ----------
ax.add_patch(Rectangle(PANEL_XY, PANEL_WH[0], PANEL_WH[1], facecolor=PANEL,
                       edgecolor="none", zorder=1))

# ---------- 5. 圆环暗底（内孔到外圈之间全部铺底色，留出分隔线） ----------
ax.add_patch(Wedge((CX, CY), R_OUT_OUT, 0, 360, width=R_OUT_OUT - R_HOLE,
                   facecolor=BG, edgecolor="none", zorder=2))

# ---------- 6. 区段切片（彩色带）与外圈蓝带 ----------
for v, size in TICKS:
    a0 = A_START + (v - 50) * A_STEP
    a1 = a0 + size
    col = seg_color(v)
    if col:
        ax.add_patch(Wedge((CX, CY), R_SEG_OUT, a0 - 90, a1 - 90,
                           width=R_SEG_OUT - R_SEG_IN, facecolor=col,
                           edgecolor=BG, lw=1.44, zorder=3))
    ax.add_patch(Wedge((CX, CY), R_OUT_OUT, a0 - 90, a1 - 90,
                       width=R_OUT_OUT - R_OUT_IN, facecolor=C_OUT,
                       edgecolor=BG, lw=1.44, zorder=4))

# ---------- 7. 刻度数字 ----------
for v, _ in TICKS:
    x, y = pol(R_LABEL, A_START + (v - 50) * A_STEP)
    ax.text(x, y + 2.5, str(v), ha="center", va="center",
            fontsize=FS_TICK, color=C_TICK, zorder=5)

# ---------- 8. 指针 ----------
nd = A_START + (NEEDLE_VAL - 50) * A_STEP
tx, ty = pol(NEEDLE_LEN, nd)
ax.plot([CX, tx], [CY, ty], color=C_NEEDLE, lw=2.0,
        solid_capstyle="butt", zorder=6)

# ---------- 9. 标题（居中于画布） ----------
fig.text(0.5, 1 - 105 / 792, "2022年6月29公司整体运营指数良好",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE,
         ha="center", va="baseline")

# ---------- 10. 输出（保存到指定目录） ----------
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "22.png"), dpi=DPI, facecolor=BG)
plt.show()
