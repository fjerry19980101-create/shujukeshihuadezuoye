# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 10:58:42 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import pandas as pd

# ============================================================
# 1. 参数区（全部来自参考图实测）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\8 数值百分比.xlsx"
SHEET = "8 数值百分比"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "8 数值百分比.png")

W, H = 1088, 756            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝
C_DARK = "#09387E"          # 深蓝（销量段）
C_LIGHT = "#82ADD7"         # 浅蓝（余量段）
C_RED = "#9B3D4F"           # 红（同比段）
C_TEXT = "#F2F2F2"          # 类目/数值/百分比文字
C_NOTE = "#D9D9D9"          # 脚注文字
FONT = "Microsoft YaHei"

# 数值轴（6 个深蓝段右缘最小二乘 K=0.148303；X0 含 +1.5px 掩膜对齐，残差 <=0.21px）
X0, K = 138.72, 0.148303    # px / 件
V_MAX = 4321.0              # 深蓝+浅蓝总长对应值（源表 销量+占位1 恒定）
RED_UNITS = 1080.25         # 红色段对应值（源表「占位2」恒定列）

# 条带行几何（6 行中心 LSQ：cy = CY0 + PITCH x i，残差 <=0.4px）
BAR_H = 59.0
CY0, PITCH = 251.14, 76.743

# 文字（fs 渲染标定）
FS_TITLE, FS_SUB = 34.6, 21.2
FS_CAT, FS_VAL, FS_PCT, FS_NOTE = 16.0, 17.0, 16.5, 15.4
X_TITLE, Y_TITLE_BASE = 69.5, 96.5      # 标题墨迹 73-714 / 56-101，左对齐
X_SUB, Y_SUB_BASE = 69.5, 140.5         # 副标题墨迹 71-1002 / 116-143
X_CAT_C = 99.5                          # 类目名中心 x
VAL_RIGHT_OFF = 19.0                    # 数值右缘(advance) = 深蓝右缘 - 19
PCT_CX_OFF = -1.6                       # 百分比中心 = 红段几何中心 - 1.6（实测 858）
Y_VAL_OFF = 9.2                         # 数值/百分比 baseline = 条带中心 + 9.2
Y_CAT_OFF = 6.7                         # 类目名 baseline = 条带中心 + 6.7
X_MID_OFF = -1.0                        # 浅蓝/红分界与红段整体左移 1px（渲染取整对齐）
X_NOTE, Y_NOTE_BASE = 60.5, 717.5       # 脚注墨迹 62-599 / 700-719

TITLE = "2021年各区域销量及同比情况"
SUBTITLE = "各区域商品销量同比去年均有下降，其中华南下降最多，同比下降20.8%"
NOTE = "*注：数据来源于公司销售系统，统计日期截至2022.01.01"


# ============================================================
# 2. 读取数据
# ============================================================
def load_data():
    """读数据：[(区域, 销量, 红段值, 同比), ...]，按图上行序（= xlsx 倒序）"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    hdr = next(i for i, r in raw.iterrows()
               if any(str(v).strip() == "区域" for v in r.values))
    sub = raw.iloc[hdr + 1:, 1:6].dropna(how="all")
    sub.columns = ["区域", "销量", "占位1", "占位2", "同比"]
    rows = []
    for _, r in sub.iterrows():
        rows.append((str(r["区域"]).strip(),
                     float(r["销量"]),
                     float(r["占位2"]),
                     float(r["同比"])))
    rows.reverse()                      # Excel 行序自下而上
    return rows


def x_of(v):
    """数值 -> 画布 x 像素"""
    return X0 + v * K


# ============================================================
# 3. 绘图
# ============================================================
def draw():
    rows = load_data()
    v_max = V_MAX                       # 深蓝+浅蓝总长（= max(销量+占位1)）

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    for i, (name, sales, red_units, yoy) in enumerate(rows):
        cy = CY0 + PITCH * i
        ytop = cy - BAR_H / 2
        x_right = x_of(sales)           # 深蓝段右缘
        x_mid = x_of(v_max) + X_MID_OFF  # 浅蓝/红分界（含 -1px 渲染取整对齐）
        x_red1 = x_mid + red_units * K   # 红段右缘

        # ---- 三段条带 ----
        ax.add_patch(Rectangle((X0, ytop), x_right - X0, BAR_H,
                               facecolor=C_DARK, edgecolor="none", zorder=2))
        ax.add_patch(Rectangle((x_right, ytop), x_mid - x_right, BAR_H,
                               facecolor=C_LIGHT, edgecolor="none", zorder=2))
        ax.add_patch(Rectangle((x_mid, ytop), x_red1 - x_mid, BAR_H,
                               facecolor=C_RED, edgecolor="none", zorder=2))

        # ---- 类目名（居中于 x=99.5，与条带垂直居中）----
        fig.text(X_CAT_C / W, (H - (cy + Y_CAT_OFF)) / H, name,
                 fontsize=FS_CAT, color=C_TEXT, fontfamily=FONT,
                 ha="center", va="baseline", zorder=3)

        # ---- 数值（白字，右缘 = 深蓝右缘 - 19）----
        fig.text((x_right - VAL_RIGHT_OFF) / W, (H - (cy + Y_VAL_OFF)) / H,
                 "%.0f" % sales,
                 fontsize=FS_VAL, color=C_TEXT, fontfamily=FONT,
                 ha="right", va="baseline", zorder=3)

        # ---- 同比百分比（白字，中心 = 红段中心 - 1.6）----
        fig.text(((x_mid + x_red1) / 2 + PCT_CX_OFF) / W,
                 (H - (cy + Y_VAL_OFF)) / H,
                 "%.1f%%" % (yoy * 100),
                 fontsize=FS_PCT, color=C_TEXT, fontfamily=FONT,
                 ha="center", va="baseline", zorder=3)

    # ---- 标题（左对齐，白色粗体）----
    fig.text(X_TITLE / W, (H - Y_TITLE_BASE) / H, TITLE,
             fontsize=FS_TITLE, color="#FFFFFF", fontweight="bold",
             fontfamily=FONT, ha="left", va="baseline", zorder=4)

    # ---- 副标题（左对齐，白色）----
    fig.text(X_SUB / W, (H - Y_SUB_BASE) / H, SUBTITLE,
             fontsize=FS_SUB, color="#FFFFFF", fontfamily=FONT,
             ha="left", va="baseline", zorder=4)

    # ---- 脚注（左对齐，#D9D9D9）----
    fig.text(X_NOTE / W, (H - Y_NOTE_BASE) / H, NOTE,
             fontsize=FS_NOTE, color=C_NOTE, fontfamily=FONT,
             ha="left", va="baseline", zorder=3)

    os.makedirs(OUT_DIR, exist_ok=True)
    fig.savefig(OUT_PNG, dpi=DPI, facecolor=BG)
    plt.close(fig)
    print("saved:", OUT_PNG)


if __name__ == "__main__":
    draw()
