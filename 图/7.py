# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 11:48:27 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------------- 画布与颜色 ----------------
W, H, DPI = 904, 506, 100
BG     = "#1A1E43"   # 背景深蓝
BLUE   = "#0070C0"   # 左条 2022年
RED    = "#E74E69"   # 右条 2021年
WHITE  = "#FFFFFF"   # 标题/副标题/区名/百分比(本图全部纯白)
NOTE_C = "#D9D9D9"   # 脚注

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------- 数据 ----------------
# (区域, 2022率, 2021率, 蓝条x0, 红条x1, 条y0, 条高,
#  蓝缝rel[距条左远端], 红缝rel[距条左中轴], 左%ink_x0, 右%ink_x0, 区名ink_x0)
DATA = [
    ("华东", 0.36, 0.42, 139, 766, 190, 26,
     [24, 70, 116, 162, 208], [24, 70, 116, 162, 208, 254], 90, 776, 417),
    ("西北", 0.31, 0.26, 170, 667, 232, 27,
     [21, 67, 113, 159],      [24, 70, 116],               90, 776, 415),
    ("东北", 0.18, 0.19, 250, 624, 275, 26,
     [21, 67],                [24, 70, 116],               94, 781, 415),
    ("华北", 0.13, 0.12, 280, 582, 317, 26,
     [21, 67],                [24, 70],                    94, 781, 417),
    ("华南", 0.09, 0.05, 305, 539, 359, 26,
     [24],                    [24],                        101, 777, 417),
]
BLUE_X1 = 359    # 蓝条中轴端(右侧对齐)
RED_X0  = 508    # 红条中轴端(左侧对齐)
BAR_H   = 26     # 条高
GAP_W_B = 1      # 蓝条缝宽
GAP_W_R = 2      # 红条缝宽

# 字号(二分标定): 标题 26.94 粗 / 副标题 19.31 / 百分比与区名 14.45 / 脚注 15.44
FS_TITLE, FS_SUB, FS_TXT, FS_NOTE = 26.94, 19.31, 14.45, 15.44

# ---------------- 画布 ----------------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)
fig.patch.set_facecolor(BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ---------------- 条 + 分段缝 + 两侧文字 ----------------
for (name, v22, v21, bx0, rx1, y0, bh, bseam, rseam, lx0, rx0, nx0) in DATA:
    # 左蓝条: 从 bx0 到中轴 359
    ax.add_patch(Rectangle((bx0, y0), BLUE_X1 - bx0, bh,
                           facecolor=BLUE, edgecolor="none", zorder=2))
    # 分段缝(距条左远端 rel, 露背景)
    for rel in bseam:
        ax.add_patch(Rectangle((bx0 + rel, y0), GAP_W_B, bh,
                               facecolor=BG, edgecolor="none", zorder=3))
    # 右红条: 从中轴 508 到 rx1
    ax.add_patch(Rectangle((RED_X0, y0), rx1 - RED_X0, bh,
                           facecolor=RED, edgecolor="none", zorder=2))
    for rel in rseam:
        ax.add_patch(Rectangle((RED_X0 + rel, y0), GAP_W_R, bh,
                               facecolor=BG, edgecolor="none", zorder=3))
    # 左侧百分比(ha=left, ink x0 已按右对齐结果硬编码; dx=1 dy=1 补偿)
    ax.text(lx0 - 1, y0 + 6.5, "%d%%" % round(v22 * 100), fontsize=FS_TXT,
            color=WHITE, ha="left", va="top", zorder=5)
    # 右侧百分比
    ax.text(rx0 - 1, y0 + 6.5, "%d%%" % round(v21 * 100), fontsize=FS_TXT,
            color=WHITE, ha="left", va="top", zorder=5)
    # 区名(中轴两列之间, ha=left)
    ax.text(nx0 - 1, y0 + 6.5, name, fontsize=FS_TXT,
            color=WHITE, ha="left", va="top", zorder=5)

# ---------------- 标题 / 副标题 / 脚注 ----------------
ax.text(55.5, 33.5, "2022年第一季度销售目标完成情况",
        fontsize=FS_TITLE, color=WHITE, fontweight="bold",
        ha="left", va="top", zorder=5)
ax.text(54.5, 79.5, "华东区域完成率最高达到36%，但是相比去年的42%有所下降",
        fontsize=FS_SUB, color=WHITE,
        ha="left", va="top", zorder=5)
ax.text(34.5, 452.5, "*注：数据来源于公司销售系统，统计日期截至2022.03.31",
        fontsize=FS_NOTE, color=NOTE_C,
        ha="left", va="top", zorder=5)

# ---------------- 保存 ----------------
out_path = os.path.join(OUT_DIR, "7 蝴蝶图.png")
fig.savefig(out_path, facecolor=BG, dpi=DPI)
print("saved:", out_path)
plt.show()
