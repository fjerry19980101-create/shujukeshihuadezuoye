# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 09:03:58 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""
复现：2022年上半年龄分布（玉玦图 / 深色底同心圆弧条形图）

参考图：1091 x 759 px 深色背景，4 条同心圆弧从 12 点方向顺时针展开，
弧的角跨度 = 占比 x 720°（xlsx 中「占位」列 = 0.5 - 占比，说明满圆对应 0.5）。
数据源：第二章 图表(后15) / 17 玉玦图.xlsx

所有坐标、颜色、字号均由参考图像素实测反推（非目测）：
  - 圆心 (550.2, 441.3)，最外环外半径 229.3（对红色外缘最小二乘圆拟合，残差 0.26px）
  - 环向外每个环 R_out 递减 37.75px，带宽恒为 32px（环间空隙 ≈ 5.8px）
  - 弧的屏幕角 phi 以 12 点为 0°、顺时针为正；像素坐标系(y向下)里
    Wedge 的 theta = (270 + phi)，此映射经 4 象限采样实测确认
  - 标题 38.5pt 粗体 / 副标题 26.5pt / 类目标签 18.5pt / 百分比标签 15pt / 脚注 15.4pt
  - 两条灰色(#A6A6A6)引出线起点恰为「蓝弧/青弧的中角 x 中半径」点（实测吻合至 0.5px）
"""

import os
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge
import pandas as pd

# ============================================================
# 1. 参数区（全部来自参考图实测）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\17 玉玦图.xlsx"
SHEET = "17 玉玦图"
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT_PNG = os.path.join(OUT_DIR, "17 玉玦图.png")

W, H = 1091, 759            # 画布像素尺寸
DPI = 100
BG = "#1A1E43"              # 背景深蓝
FONT = "Microsoft YaHei"    # 微软雅黑

# 实测色值（按 类目 -> 颜色）
COLORS = {
    "[20,30)": "#E74E69",   # 玫红（最外环）
    "[30,40)": "#F5C353",   # 金黄
    "[40,50)": "#11A7AD",   # 青绿
    ">=50":    "#37A2DA",   # 浅蓝（最内环）
}

# 圆心与环几何（实测）
CX, CY = 550.2, 441.3
R_OUT0 = 229.3              # 最外环外半径
PITCH = 37.75               # 相邻环外半径差
BAND = 32.0                 # 环带宽（径向厚度）
FULL = 0.5                  # 满圆对应的占比（占位列 = 0.5 - 占比）

# 引出线（实测）：仅 12.5% 与 20.8% 两条，灰色 #A6A6A6
# 线宽：REF 逐行占 2px（含端面）→ 2.0px = 1.44pt @100dpi（逐行扫描标定）
# 终点为实测像素；起点由「弧中角 x 中半径」几何计算（见 draw()）
LEADERS = [
    {"cat": ">=50",    "end": (641.0, 441.5)},                  # 蓝弧：终点在弧端面
    {"cat": "[40,50)", "end": (637.0, 565.5),                  # 青弧：斜线终点 + 向左短横脚
     "foot_x0": 623.5, "foot_y": 567.0},                       # 脚占 y566-567 两行 → 中心 567.0
]
LEADER_COLOR = "#A6A6A6"
LEADER_LW = 1.44             # pt；1.44pt x 100dpi = 2.0px（与 REF 线宽一致）

# 文字实测参数 ---------------------------------------------------
TITLE = "2022年上半年龄分布"
SUBTITLE = "公司平均年龄32.5， 23-30员工比例最高"
NOTE = "*注：数据来源于公司人力资源系统，统计日期截至2022.06.30"

FS_TITLE, FS_SUB, FS_CAT, FS_PCT, FS_NOTE = 38.5, 26.5, 18.5, 15.0, 15.4

# 各文字的 (x_px, baseline_px)；baseline 由墨迹底行反推
#   粗体标题 ≈ 底行-5.5；CJK 副标题/脚注 ≈ 底行-3.5；数字标签 ≈ 底行+0.5
#   （x/baseline 均经「渲染-参考」逐元素 bbox diff 迭代校准，误差 ≤3px）
X_TITLE, Y_TITLE_BASE = 73.0, 91 - 5.5      # 标题墨迹 77-577 / 41-91
X_SUB, Y_SUB_BASE = 74.0, 146 - 4.5         # 副标题墨迹 75-735 / 111-146
X_NOTE, Y_NOTE_BASE = 69.0, 722 - 2.0       # 脚注墨迹 71-655 / 704-722

# 类目标签：左对齐，baseline 逐个实测；>=50 单独 17pt（18.5pt 墨迹偏宽 5px）
CAT_LABEL_POS = {
    "[20,30)": (458.0, 235.0, 18.5),   # 墨迹 462-540 / 216-240（括号下探 3.5px）
    "[30,40)": (458.0, 275.0, 18.5),   # 墨迹 462-540 / 256-280
    "[40,50)": (458.0, 315.0, 18.5),   # 墨迹 462-540 / 296-320
    ">=50":    (458.0, 357.0, 17.0),   # 墨迹 461-516 / 338-355（数字底=基线）
}
# 百分比标签：左对齐（墨迹左缘-1），baseline = 数字墨迹底行 +0.5
PCT_LABEL_POS = {
    "[20,30)": (310.0, 419.5),   # 37.5%  墨迹 313-368 / 404-419
    "[30,40)": (415.0, 575.5),   # 29.2%  墨迹 416-472 / 559-575
    "[40,50)": (557.0, 575.5),   # 20.8%  墨迹 558-614 / 559-575
    ">=50":    (613.0, 469.5),   # 12.5%  墨迹 614-669 / 453-468
}


# ============================================================
# 2. 读取数据
# ============================================================
def load_data():
    """从 xlsx 读取年龄组与占比；表头行用「含'年龄'」定位，鲁棒。
    返回按占比降序（外环->内环）的列表 [(类目, 占比), ...]"""
    raw = pd.read_excel(XLSX, sheet_name=SHEET, header=None)
    hdr = next(i for i, r in raw.iterrows()
               if any(str(v).strip() == "年龄" for v in r.values))
    sub = raw.iloc[hdr + 1:, 1:3].dropna()
    sub.columns = ["年龄", "占比"]
    sub["年龄"] = sub["年龄"].astype(str).str.strip()
    sub["占比"] = sub["占比"].astype(float)
    sub = sub.sort_values("占比", ascending=False).reset_index(drop=True)
    return list(sub.itertuples(index=False, name=None))


# ============================================================
# 3. 绘图
# ============================================================
def arc_point(pct, ring_idx, frac=0.5):
    """弧上「中角 x 中半径」处的像素坐标（引出线起点）。
    frac: 取弧 span 的比例位置（0.5 = 中角）"""
    r_out = R_OUT0 - PITCH * ring_idx
    r_mid = r_out - BAND / 2.0
    phi = pct / FULL * 360.0 * frac        # 屏幕角：12点=0，顺时针
    t = math.radians(phi)
    return (CX + r_mid * math.sin(t), CY - r_mid * math.cos(t))


def draw():
    data = load_data()

    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    # 像素坐标系：数据坐标 == 画布像素，y 轴向下（与图片行号一致）
    ax = fig.add_axes([0, 0, 1, 1])
    ax.set_facecolor(BG)
    ax.set_xlim(0, W)
    ax.set_ylim(H, 0)
    ax.axis("off")

    # ---- 同心弧（玉玦）----
    # 第 i 环（i=0 最外）：R_out = 229.3 - 37.75i，带宽 32
    # 屏幕角 phi 从 0 顺时针扫 span = 占比/0.5 x 360；
    # 像素坐标系(y 向下)下 Wedge 的 theta = 270 + phi（4 象限实测映射），
    # Wedge 自 theta1 逆时针画到 theta2 == 屏幕上自 12 点顺时针扫。
    for i, (cat, pct) in enumerate(data):
        r_out = R_OUT0 - PITCH * i
        r_in = r_out - BAND
        span = pct / FULL * 360.0
        t1, t2 = 270.0, 270.0 + span
        ax.add_patch(Wedge((CX, CY), r_out, t1, t2, width=r_out - r_in,
                           facecolor=COLORS.get(cat, "#888888"),
                           edgecolor="none", zorder=2))

    # ---- 灰色引出线（起点 = 弧中角x中半径，终点/短横脚为实测）----
    ring_idx = {cat: i for i, (cat, _) in enumerate(data)}
    for ld in LEADERS:
        cat = ld["cat"]
        i = ring_idx[cat]
        pct = dict(data)[cat]
        x0, y0 = arc_point(pct, i, frac=0.5)
        x1, y1 = ld["end"]
        ax.plot([x0, x1], [y0, y1], color=LEADER_COLOR, lw=LEADER_LW,
                solid_capstyle="butt", zorder=4)
        if "foot_x0" in ld:      # 青弧引出线底部的向左短横脚
            ax.plot([ld["foot_x0"], x1], [ld["foot_y"], ld["foot_y"]],
                    color=LEADER_COLOR, lw=LEADER_LW,
                    solid_capstyle="butt", zorder=4)

    # ---- 类目标签（白，左对齐；>=50 单独 17pt）----
    for cat, (x_px, y_base, fs) in CAT_LABEL_POS.items():
        fig.text(x_px / W, (H - y_base) / H, cat, fontsize=fs,
                 color="#FFFFFF", fontfamily=FONT, ha="left", va="baseline")

    # ---- 百分比标签（白，数字保留 1 位小数）----
    for cat, (x_px, y_base) in PCT_LABEL_POS.items():
        pct = dict(data)[cat]
        s = "%.1f%%" % (pct * 100)
        fig.text(x_px / W, (H - y_base) / H, s, fontsize=FS_PCT,
                 color="#FFFFFF", fontfamily=FONT, ha="left", va="baseline")

    # ---- 标题 / 副标题 / 脚注 ----
    # fig.text 的 y 从画布底部起算，须用 (H - baseline_px)/H 换算
    fig.text(X_TITLE / W, (H - Y_TITLE_BASE) / H, TITLE, fontsize=FS_TITLE,
             color="#FFFFFF", fontweight="bold", fontfamily=FONT,
             ha="left", va="baseline")
    fig.text(X_SUB / W, (H - Y_SUB_BASE) / H, SUBTITLE, fontsize=FS_SUB,
             color="#FFFFFF", fontfamily=FONT, ha="left", va="baseline")
    fig.text(X_NOTE / W, (H - Y_NOTE_BASE) / H, NOTE, fontsize=FS_NOTE,
             color="#D9D9D9", fontfamily=FONT, ha="left", va="baseline")

    os.makedirs(OUT_DIR, exist_ok=True)
    fig.savefig(OUT_PNG, dpi=DPI, facecolor=BG)
    plt.close(fig)
    print("saved:", OUT_PNG)


if __name__ == "__main__":
    draw()
