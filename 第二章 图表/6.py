# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 11:48:24 2026

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
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\6 蝴蝶图.xlsx"
SHEET = "6 蝴蝶图"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "6 蝴蝶图.png")

W, H = 1088, 756            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝
C_BLUE = "#0070C0"          # 左条（2022）
C_RED = "#E74E69"           # 右条（2021）
C_YEAR_R = "#F94C66"        # 年份「2021」文字（比条带红亮）
C_CAT = "#F2F2F2"           # 类目名
C_VAL_L = "#F2F2F2"         # 左条数值
C_VAL_R = "#D9D9D9"         # 右条数值（实测比左侧暗一档）
C_NOTE = "#D9D9D9"          # 脚注
FONT = "Microsoft YaHei"

# 左条映射（5 个蓝条左缘 LSQ：x_left = XL - KB x 2022销量，残差 <=0.28px）
XL, KB = 482.61, 0.128955
# 右条映射（5 个红条右缘 LSQ：x_right = XR + KR x 2021销量，残差 <=0.15px；XR 含 +1px 掩膜对齐）
XR, KR = 637.93, 0.128787

# 条带行几何（5 行中心实测，等距 90.5）
BAR_H = 36.5
CY0, PITCH = 260.5, 90.5

# 文字（fs 渲染标定）
FS_TITLE, FS_SUB = 34.6, 21.2
FS_CAT, FS_VAL, FS_YEAR, FS_NOTE = 15.8, 15.0, 18.5, 15.4
X_TITLE, Y_TITLE_BASE = 70.5, 86.0      # 标题墨迹 74-811 / 46-91，左对齐
X_SUB, Y_SUB_BASE = 70.0, 131.5         # 副标题墨迹 73-862 / 106-133
X_CAT_C = 561.5                         # 类目名中心 x（中缝居中）
VAL_IN_L, VAL_IN_R = 18.5, 18.0         # 数值距条外端（左条 ha=left / 右条 ha=right）
Y_VAL_OFF = 8.2                         # 数值 baseline = 条带中心 + 8.2
Y_CAT_OFF = 6.7                         # 类目名 baseline = 条带中心 + 6.7
Y_YEAR_BASE = 223.0                     # 年份标签 baseline（墨迹 204-221）
X_YEAR_L, X_YEAR_R = 403.0, 683.0      # 年份标签中心 x（实测）
X_NOTE, Y_NOTE_BASE = 68.5, 718.5       # 脚注墨迹 70-611 / 701-720

TITLE = "2022年上半年各区域对比去年销量"
SUBTITLE = "2022年整体销量高于2021年，只有东北区域较2021有所下降"
NOTE = "*注：数据来源于公司销售系统，统计日期截至2022.06.30"


# ============================================================
# 2. 读取数据
# ============================================================
def load_data():
    """读数据：[(区域, 2022销量, 2021销量), ...]，按图上行序（= xlsx 倒序）"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    hdr = next(i for i, r in raw.iterrows()
               if any(str(v).strip() == "区域" for v in r.values))
    sub = raw.iloc[hdr + 1:, 1:6].dropna(how="all")
    sub.columns = ["区域", "占位1", "销量2022", "占位2", "销量2021"]
    rows = []
    for _, r in sub.iterrows():
        rows.append((str(r["区域"]).strip(),
                     float(r["销量2022"]),
                     float(r["销量2021"])))
    rows.reverse()                      # Excel 行序自下而上
    return rows


# ============================================================
# 3. 绘图
# ============================================================
def draw():
    rows = load_data()

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    for i, (name, v22, v21) in enumerate(rows):
        cy = CY0 + PITCH * i
        ytop = cy - BAR_H / 2

        # ---- 左条（2022，蓝）：右缘恒定中线 XL ----
        x_left = XL - v22 * KB
        ax.add_patch(Rectangle((x_left, ytop), v22 * KB, BAR_H,
                               facecolor=C_BLUE, edgecolor="none", zorder=2))
        # ---- 右条（2021，红）：左缘恒定 XR ----
        ax.add_patch(Rectangle((XR, ytop), v21 * KR, BAR_H,
                               facecolor=C_RED, edgecolor="none", zorder=2))

        # ---- 类目名（中缝居中，与条带垂直居中）----
        fig.text(X_CAT_C / W, (H - (cy + Y_CAT_OFF)) / H, name,
                 fontsize=FS_CAT, color=C_CAT, fontfamily=FONT,
                 ha="center", va="baseline", zorder=3)

        # ---- 左条数值（靠条左缘内侧）----
        fig.text((x_left + VAL_IN_L) / W, (H - (cy + Y_VAL_OFF)) / H,
                 "%.0f" % v22,
                 fontsize=FS_VAL, color=C_VAL_L, fontfamily=FONT,
                 ha="left", va="baseline", zorder=3)

        # ---- 右条数值（靠条右缘内侧）----
        fig.text((XR + v21 * KR - VAL_IN_R) / W, (H - (cy + Y_VAL_OFF)) / H,
                 "%.0f" % v21,
                 fontsize=FS_VAL, color=C_VAL_R, fontfamily=FONT,
                 ha="right", va="baseline", zorder=3)

    # ---- 年份标签（各居中于实测位置）----
    fig.text(X_YEAR_L / W, (H - Y_YEAR_BASE) / H, "2022",
             fontsize=FS_YEAR, color=C_BLUE, fontfamily=FONT,
             ha="center", va="baseline", zorder=3)
    fig.text(X_YEAR_R / W, (H - Y_YEAR_BASE) / H, "2021",
             fontsize=FS_YEAR, color=C_YEAR_R, fontfamily=FONT,
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
