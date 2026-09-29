# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 10:43:29 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import datetime as dt
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle
import pandas as pd

# ============================================================
# 1. 参数区（全部来自参考图实测）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\10 甘特图.xlsx"
SHEET = "10 甘特图"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "10 甘特图.png")

W, H = 1157, 756            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝
C_LIGHT = "#00A3E6"         # 浅蓝（已完成部分）
C_DARK = "#0070C0"          # 深蓝（剩余部分）
C_GRID = "#717389"          # 虚线网格
C_DATE = "#F2F2F2"          # 日期刻度文字
C_TASK = "#B2B2BC"          # 任务名文字
FONT = "Microsoft YaHei"

# 时间轴（14 个条带端点最小二乘：x = X0 + K x 天数，残差 <=0.94px）
D0 = dt.datetime(2022, 3, 1)
X0, K = 191.72, 8.1813      # px / 天

# 网格虚线：3/1 起每 15 天一条，共 8 条
GRID_STEP_DAYS = 15
GRID_N = 8
GRID_Y0, GRID_Y1 = 170.0, 711.0
GRID_LW = 2.0 / DPI * 72                    # 2px -> 1.44pt
# 注意：matplotlib 的 dash 序列(pt)会被线宽缩放：渲染px = spec_pt x (DPI/72) x lw_pt
# 此处 lw=1.44pt -> 缩放因子 2.0，故 spec 取 8.0/3.0pt，渲染为 16px on / 6px off
GRID_DASH = (0, (8.0, 3.0))

# 条带行几何
BAR_H = 38.5                                # 条高 px
BAR_CY0 = 208.0                             # 第 1 条中心 y
BAR_PITCH = 77.5                            # 行距

# 文字（fs 渲染标定）
FS_TITLE, FS_DATE, FS_TASK, FS_PCT = 30.8, 15.0, 17.0, 15.0
X_TITLE, Y_TITLE_BASE = 578.5, 57 - 5.5     # 标题墨迹 273-884 / 16-57，居中
Y_DATE_BASE = 147.0                         # 日期墨迹 131-148（渲染标定下移 1.5）
X_TASK, Y_TASK_BASE0 = 81.0, 215.0          # 任务名墨迹左缘 82；首行 baseline（标定下移 1.5）
Y_PCT_OFF = 7.5                             # 百分比 baseline = 条带中心 + 7.5


# ============================================================
# 2. 读取数据
# ============================================================
def load_data():
    """读甘特图数据：[(项目名称, 开始日期, 项目天数, 完成度, 进行天数), ...]"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    hdr = next(i for i, r in raw.iterrows()
               if any(str(v).strip() == "项目名称" for v in r.values))
    sub = raw.iloc[hdr + 1:, 1:6].dropna(how="all")
    sub.columns = ["名称", "开始", "天数", "完成度", "进行天数"]
    out = []
    for _, r in sub.iterrows():
        out.append((str(r["名称"]).strip(),
                    pd.to_datetime(r["开始"]).to_pydatetime(),
                    float(r["天数"]),
                    float(r["完成度"]),
                    float(r["进行天数"])))
    return out


def x_of(d):
    """日期 -> 画布 x 像素"""
    return X0 + (d - D0).days * K


# ============================================================
# 3. 绘图
# ============================================================
def draw():
    data = load_data()

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    # ---- 竖直虚线网格（每 15 天一条，共 8 条）----
    for i in range(GRID_N):
        gx = X0 + (GRID_STEP_DAYS * i) * K
        ax.plot([gx, gx], [GRID_Y0, GRID_Y1], color=C_GRID, lw=GRID_LW,
                linestyle=GRID_DASH, zorder=1)

    # ---- 甘特条带：深色整条 + 浅色已完成段 ----
    for i, (name, start, days, pct, done_days) in enumerate(data):
        cy = BAR_CY0 + BAR_PITCH * i
        bx0 = x_of(start)
        bw = days * K                       # 条带宽 = 项目天数
        bw_done = done_days * K             # 已完成段宽 = 进行天数
        ytop = cy - BAR_H / 2
        ax.add_patch(Rectangle((bx0, ytop), bw, BAR_H,
                               facecolor=C_DARK, edgecolor="none", zorder=2))
        ax.add_patch(Rectangle((bx0, ytop), bw_done, BAR_H,
                               facecolor=C_LIGHT, edgecolor="none", zorder=3))

    # ---- 日期刻度（网格线顶端，居中于网格线）----
    for i in range(GRID_N):
        d = D0 + dt.timedelta(days=GRID_STEP_DAYS * i)
        gx = X0 + GRID_STEP_DAYS * i * K
        fig.text(gx / W, (H - Y_DATE_BASE) / H, "%d/%d/%d" % (d.year, d.month, d.day),
                 fontsize=FS_DATE, color=C_DATE, fontfamily=FONT,
                 ha="center", va="baseline", zorder=3)

    # ---- 任务名（左对齐，与条带垂直居中）----
    for i, (name, *_rest) in enumerate(data):
        cy = BAR_CY0 + BAR_PITCH * i
        fig.text(X_TASK / W, (H - (cy + (Y_TASK_BASE0 - BAR_CY0))) / H, name,
                 fontsize=FS_TASK, color=C_TASK, fontfamily=FONT,
                 ha="left", va="baseline", zorder=3)

    # ---- 百分比（白色，居中于整条条带）----
    for i, (name, start, days, pct, done_days) in enumerate(data):
        cy = BAR_CY0 + BAR_PITCH * i
        cx = x_of(start) + days * K / 2
        fig.text(cx / W, (H - (cy + Y_PCT_OFF)) / H, "%.0f%%" % (pct * 100),
                 fontsize=FS_PCT, color="#FFFFFF", fontfamily=FONT,
                 ha="center", va="baseline", zorder=4)

    # ---- 标题（画布居中，白色粗体）----
    fig.text(X_TITLE / W, (H - Y_TITLE_BASE) / H, "2022年化妆品类目采购项目进度",
             fontsize=FS_TITLE, color="#FFFFFF", fontweight="bold",
             fontfamily=FONT, ha="center", va="baseline", zorder=4)

    os.makedirs(OUT_DIR, exist_ok=True)
    fig.savefig(OUT_PNG, dpi=DPI, facecolor=BG)
    plt.close(fig)
    print("saved:", OUT_PNG)


if __name__ == "__main__":
    draw()
