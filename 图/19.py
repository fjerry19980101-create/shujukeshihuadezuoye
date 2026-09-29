# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:52:02 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle
import pandas as pd

# ============================================================
# 1. 参数区（全部来自实测）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\19 南丁格尔圆饼图.xlsx"
SHEET = "19 南丁格尔圆饼图"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "19 南丁格尔圆饼图.png")

W, H = 1088, 792            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝

FONT = "Microsoft YaHei"    # 微软雅黑

# 实测色值（按扇区顺时针顺序）
COLORS = {
    "销售部": "#E74E69",   # 玫红
    "采购部": "#F5C353",   # 金黄
    "工程部": "#11A7AD",   # 青绿
    "财务部": "#0070C0",   # 深蓝
    "行政部": "#37A2DA",   # 浅蓝
    "人力部": "#7030A0",   # 紫
}
ORDER = ["销售部", "采购部", "工程部", "财务部", "行政部", "人力部"]

# 圆心与半径（实测）
CX, CY, R = 540.85, 465.57, 276.50

# 引出线实测参数 --------------------------------------------------
# 左侧三条：只有一条水平线，从外侧指向饼图边缘（无圆点、无折角）
# (颜色key, 水平线 y, 线左端 x, 线右端 x)
LEFT_LINES = [
    ("行政部", 321.0, 207, 475),   # 浅蓝
    ("财务部", 501.0, 148, 445),   # 深蓝
    ("工程部", 683.0, 175, 330),   # 青绿（左端 175 起）
]
# 右侧两条：水平线 + 折角 + 斜线 + 彩色圆点
# (颜色key, 水平线 y, 水平线左端x[折角], 水平线右端x, 圆点中心x, 圆点中心y, 圆点半径)
RIGHT_LINES = [
    ("销售部", 271.0, 856, 1011, 813.5, 320.5, 5.0),   # 玫红
    ("采购部", 673.0, 781,  936, 751.0, 637.0, 5.0),   # 金黄
]
LW = 2.6                     # 引出线线宽(px)，实测约 3px

# 文字实测参数 ---------------------------------------------------
TITLE = "2021年各部门人数分布"
SUBTITLE = "公司总人数1664，销售部人数最多451，占比29.2%"
NOTE = "*注：数据来源于公司人力资源系统，统计日期截至2022.01.01"

X_TITLE, X_SUB, X_NOTE = 72 / W, 70 / W, 73 / W
# fig.text 的 y 从画布「底部」起算，故用 (H - 像素y) / H 换算；
# baseline 相对墨迹底行：标题(粗体) 底行-5、副标题/脚注(CJK) 底行-3
Y_TITLE_BASE = (H - (74 - 5)) / H    # 标题墨迹底行 74
Y_SUB_BASE = (H - (127 - 3)) / H     # 副标题墨迹底行 127
Y_NOTE_BASE = (H - (757 - 3)) / H    # 脚注墨迹底行 757

FS_TITLE, FS_SUB, FS_NOTE = 34.5, 23.1, 15.4


# ============================================================
# 2. 读取数据
# ============================================================
def load_data():
    """从 xlsx 读取部门与人数占比；表头行用「含'部门'」定位，鲁棒。"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    hdr = next(i for i, r in raw.iterrows()
               if any(str(v).strip() == "部门" for v in r.values))
    sub = raw.iloc[hdr + 1:, 1:3].dropna()
    sub.columns = ["部门", "占比"]
    sub["部门"] = sub["部门"].astype(str).str.strip()
    sub["占比"] = sub["占比"].astype(float)
    # 按数值降序（最大值在正上方开始，顺时针）
    sub = sub.sort_values("占比", ascending=False).reset_index(drop=True)
    return sub


# ============================================================
# 3. 绘图
# ============================================================
def draw():
    df = load_data()

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 用像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    # ---- 饼图扇区 ----
    # Wedge 的 theta 以 x 轴正向为 0、数学逆时针为正；本图 y 轴已翻转(ylim=(H,0))，
    # 实测映射为：theta=0→屏幕右、theta=90→屏幕下、theta=180→屏幕左、theta=270→屏幕上。
    # 于是「12 点方向为 0°、顺时针递增」的屏幕角 phi 对应 theta = (270 + phi) % 360。
    start_deg = 0.0
    for _, row in df.iterrows():
        name, val = row["部门"], row["占比"]
        span = val * 360.0
        d1, d2 = start_deg, start_deg + span
        t1, t2 = (270.0 + d1) % 360.0, (270.0 + d2) % 360.0
        col = COLORS.get(name, "#888888")
        ax.add_patch(Wedge((CX, CY), R, t1, t2,
                           facecolor=col, edgecolor="none", zorder=2))
        start_deg = d2

    # ---- 左侧引出线（水平，无圆点）----
    for key, y, x0, x1 in LEFT_LINES:
        ax.plot([x0, x1], [y, y], color=COLORS[key], lw=LW,
                solid_capstyle="butt", zorder=1)

    # ---- 右侧引出线（水平 + 折角斜线 + 圆点）----
    for key, y, x_elbow, x_end, dx, dy, dr in RIGHT_LINES:
        col = COLORS[key]
        # 水平段
        ax.plot([x_elbow, x_end], [y, y], color=col, lw=LW,
                solid_capstyle="butt", zorder=1)
        # 斜线段：折角 -> 圆点
        ax.plot([x_elbow, dx], [y, dy], color=col, lw=LW,
                solid_capstyle="butt", zorder=1)
        # 圆点（像素坐标系下 Circle 即正圆）
        ax.add_patch(Circle((dx, dy), dr, facecolor=col, edgecolor="none", zorder=3))

    # ---- 标题 / 副标题 / 脚注 ----
    fig.text(X_TITLE, Y_TITLE_BASE, TITLE, fontsize=FS_TITLE, color="#FFFFFF",
             fontweight="bold", fontfamily=FONT, ha="left", va="baseline")
    fig.text(X_SUB, Y_SUB_BASE, SUBTITLE, fontsize=FS_SUB, color="#FFFFFF",
             fontfamily=FONT, ha="left", va="baseline")
    fig.text(X_NOTE, Y_NOTE_BASE, NOTE, fontsize=FS_NOTE, color="#D9D9D9",
             fontfamily=FONT, ha="left", va="baseline")

    os.makedirs(OUT_DIR, exist_ok=True)
    fig.savefig(OUT_PNG, dpi=DPI, facecolor=BG)
    plt.close(fig)
    print("saved:", OUT_PNG)


if __name__ == "__main__":
    draw()
