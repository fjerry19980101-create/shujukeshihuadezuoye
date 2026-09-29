# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:46:18 2026

@author: fjwsm
"""

import os
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge

mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

W, H, DPI = 1088, 792, 100
BG = "#1A1E43"
CX, CY = 488.90, 453.80      # 共同圆心（像素坐标）
LINE_W = 15.0                # 弧线宽 px

# (部门, 人数, 颜色, 中线半径px, 屏幕起始角°(自12点顺时针), 顺时针扫角°)
DEPTS = [
    ("销售部", 451, "#E74E69", 247.96, 88.85, 192.97),
    ("采购部", 326, "#F5C353", 224.10, 88.73, 138.65),
    ("工程部", 293, "#11A7AD", 200.26, 88.59, 124.09),
    ("财务部", 238, "#37A2DA", 176.43, 88.40,  99.95),
    ("行政部", 226, "#0070C0", 152.67, 88.17,  94.38),
    ("人力部", 130, "#7030A0", 128.90, 87.85,  52.11),
]

fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)            # 像素坐标系，y 向下
ax.axis("off")

# ---- 同心跑道弧 ----
# y 轴翻转坐标系中 Wedge 的数据角 θ 对应屏幕角 -θ：
# 屏幕上"从 start_deg 顺时针扫 sweep 度" == 数据角从 -start_deg 逆时针扫 sweep 度
for name, val, color, r_mid, start_deg, sweep in DEPTS:
    ax.add_patch(Wedge((CX, CY), r_mid + LINE_W / 2, -start_deg, -start_deg + sweep,
                       width=LINE_W, facecolor=color, edgecolor="none"))

# ---- 左侧标签（部门名 + 人数），右对齐，与各弧端面垂直居中（行距 24px）----
for i, (name, val, color, r_mid, start_deg, sweep) in enumerate(DEPTS):
    ax.text(484.5, 214.5 + 24.0 * i, "%s %d" % (name, val),
            fontsize=13.0, color="#F2F2F2",
            ha="right", va="baseline", zorder=5)

# ---- 标题 / 副标题 / 脚注 ----
ax.text(70, 89, "2022年上半年各部门人数",
        fontsize=38.5, color="#FFFFFF", fontweight="bold",
        ha="left", va="baseline", zorder=5)
ax.text(71, 147.5, "公司总人数1664，销售部人数最多451， 占比27%",
        fontsize=26.6, color="#FFFFFF",
        ha="left", va="baseline", zorder=5)
ax.text(50, 753, "*注：数据来源于公司人力资源系统，统计日期截止2022.06.30",
        fontsize=15.4, color="#D9D9D9",
        ha="left", va="baseline", zorder=5)

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "18 跑道图.png"), dpi=DPI, facecolor=BG)
plt.show()
