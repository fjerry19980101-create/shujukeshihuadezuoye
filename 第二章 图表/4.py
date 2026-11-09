# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 12:43:39 2026

@author: fjwsm
"""

import os
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------------- 画布与颜色 ----------------
W, H, DPI = 1088, 792, 100
BG     = "#1A1E43"   # 背景深蓝
WHITE  = "#F2F2F2"   # 类目 / Y刻度 / 气泡数字
NOTE_C = "#D9D9D9"   # 脚注
GRID_C = "#454866"   # 虚线网格

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------- 数据 ----------------
# (类目, 销量, 柱色, 柱x0, 柱w, 类目ink_x0)
BARS = [
    ("口红", 9221, "#7BBDD5", 182, 53, 189),
    ("面膜", 5102, "#49A098", 289, 53, 293),
    ("隔离", 6571, "#E66B4C", 396, 53, 401),
    ("防晒", 5760, "#FFC000", 503, 53, 508),
    ("精华", 6321, "#0097E0", 609, 54, 614),
    ("面霜", 8612, "#0070C0", 716, 54, 721),
    ("眼影", 2645, "#4A5BD1", 823, 54, 829),
    ("气垫", 5321, "#464CAC", 930, 54, 935),
]

# 气泡(与 BARS 逐位对应): (气泡x0, 顶y, 宽, 矩形高,
#   尾巴底边x0, 底边x1, 尖点x, 尖点y, 白字ink_x0, 白字ink_y0)
# 矩形底边 y = 顶y+矩形高-0.5; 尾巴为直角三角形(一边近垂直一边斜),
# 尖点指向柱顶中心方向, 逐气泡实测硬编码
BUBS = [
    (183, 219, 59, 34, 194, 206, 204.0,  265.5, 192, 230),
    (287, 388, 59, 34, 321, 334, 320.5,  433.5, 296, 399),
    (392, 332, 59, 34, 403, 416, 418.0,  378.5, 401, 343),
    (503, 367, 59, 34, 537, 549, 534.5,  407.5, 511, 378),
    (609, 348, 60, 34, 621, 633, 635.5,  387.5, 619, 359),
    (716, 248, 60, 33, 727, 740, 744.0,  291.5, 725, 259),
    (823, 496, 59, 34, 834, 847, 848.5,  542.5, 833, 507),
    (934, 383, 59, 34, 944, 957, 959.0,  426.5, 943, 394),
]

# 柱顶标定公式: top = 662.6 - 0.042578 * value (残差 <= 0.7px)
def top_of(v):
    return 662.6 - 0.042578 * v

BASE_Y   = 662.6          # 0 值基线(与最下一条网格线重合)
GRID_YS  = [236.5, 321.5, 406.5, 492.5, 577.5, 662.6]
GRID_X   = (155, 1009)    # 网格 x 范围

# Y 刻度 (ha=left, ink 左上角实测): (ink_x0, ink_y0, 文本)
YTICKS = [
    (81,  230, "10000"),
    (92,  315, "8000"),
    (92,  400, "6000"),
    (92,  485, "4000"),
    (92,  570, "2000"),
    (128, 655, "0"),
]
CAT_Y = 680               # 类目 ink 顶 y (统一)

# 字号(二分标定): 标题 34.61 粗 / 副标题 23.09 / 脚注 15.36
# Y刻度 14.81 / 类目 15.4 / 气泡数字 13.89
FS_TITLE, FS_SUB, FS_NOTE = 34.61, 23.09, 15.36
FS_YT, FS_CAT, FS_BUB     = 14.81, 15.4, 13.89

# ---------------- 画布 ----------------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)
fig.patch.set_facecolor(BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ---------------- 虚线网格 (16on/6off, 在柱下层) ----------------
for gy in GRID_YS:
    ax.plot(list(GRID_X), [gy, gy], color=GRID_C, lw=2,
            linestyle=(0, (8, 3)), solid_capstyle="butt", zorder=1)

# ---------------- 柱 + 气泡标注 + 气泡数字 + 类目 ----------------
for (cat, v, c, x0, bw, cat_x0), (bx0, bytop, bw2, bh2,
        tx0, tx1, tipx, tipy, txt_x0, txt_y0) in zip(BARS, BUBS):
    # 柱体
    t = top_of(v)
    ax.add_patch(Rectangle((x0, t), bw, BASE_Y - t,
                           facecolor=c, edgecolor="none", zorder=2))
    # 气泡矩形(与柱同色)
    ax.add_patch(Rectangle((bx0, bytop), bw2, bh2,
                           facecolor=c, edgecolor="none", zorder=3))
    # 气泡尾巴(直角三角形, 底边贴矩形底, 尖点向下)
    tail_y = bytop + bh2 - 0.5
    ax.add_patch(Polygon([(tx0, tail_y), (tx1, tail_y), (tipx, tipy)],
                         closed=True, facecolor=c, edgecolor="none", zorder=3))
    # 气泡内白色数字 (dx=1 dy=1 补偿)
    ax.text(txt_x0 - 1, txt_y0 - 1, str(v), fontsize=FS_BUB, color=WHITE,
            ha="left", va="top", zorder=5)
    # 类目 (dx=2 dy=-1 补偿)
    ax.text(cat_x0 - 2, CAT_Y + 1, cat, fontsize=FS_CAT, color=WHITE,
            ha="left", va="top", zorder=5)

# ---------------- Y 刻度 (dx=2 dy=1 补偿) ----------------
for x0, y0, s in YTICKS:
    ax.text(x0 - 2, y0 - 1, s, fontsize=FS_YT, color=WHITE,
            ha="left", va="top", zorder=5)

# ---------------- 标题 / 副标题 / 脚注 ----------------
ax.text(71, 57, "2021年商品销量情况",
        fontsize=FS_TITLE, color="#FFFFFF", fontweight="bold",
        ha="left", va="top", zorder=5)
ax.text(70, 135, "口红销量最好达9221，是眼影最低值2645近3.5倍",
        fontsize=FS_SUB, color="#FFFFFF",
        ha="left", va="top", zorder=5)
ax.text(62, 742, "*注：数据来源于公司销售系统，统计日期截至2022.08.31",
        fontsize=FS_NOTE, color=NOTE_C,
        ha="left", va="top", zorder=5)

# ---------------- 保存 ----------------
out_path = os.path.join(OUT_DIR, "4 标注柱形图.png")
fig.savefig(out_path, facecolor=BG, dpi=DPI)
print("saved:", out_path)
plt.show()
