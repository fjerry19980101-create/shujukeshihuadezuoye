# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 10:01:36 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-

import os

import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, Ellipse, PathPatch
from matplotlib.path import Path
from scipy.interpolate import PchipInterpolator

mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

W, H, DPI = 1082, 790, 100
BG = "#1A1E43"
BLUE = "#0070C0"

PCT = 65                       # 完成率 0.65 -> 65%

# ---- 几何（像素实测）----
WX, WY, WR = 565.3, 442.5, 229.4            # 水体圆
RX, RY = 562.0, 449.5                       # 外环中心
RMX, RMY = 251.0, 243.5                     # 外环中线半径（椭圆，素材比例 195:190）
RING_LW = 2.88                              # 环线宽 pt（= 4 px）

# ---- 波浪锚点 (x, y)：逐像素实测；文字遮挡区按两侧趋势平滑内插 ----
WAVE_PTS = [
    (336, 436), (344, 404), (348, 394), (352, 392), (360, 390), (364, 389),
    (372, 388), (392, 388), (400, 390), (412, 392), (424, 395), (436, 399),
    (448, 403), (460, 407), (472, 410), (484, 414),
    (500, 418.0), (516, 419.5), (532, 418.0), (548, 413.0), (564, 407.5),
    (580, 402.5), (596, 398.0),
    (608, 396.5), (620, 392.0), (632, 388.5), (644, 392.0), (656, 395.5),
    (668, 398.0),
    (676, 401), (688, 404.5), (700, 408), (712, 412), (724, 415), (736, 417),
    (750, 417.5), (764, 417), (772, 415), (780, 413), (786, 414), (792, 416),
]

fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)              # 像素坐标系，y 向下
ax.axis("off")

# 1) 水体圆
ax.add_patch(Circle((WX, WY), WR, facecolor=BLUE, edgecolor="none", zorder=1))

# 2) 波浪上方盖背景色多边形 -> 露出波浪以下的水体
xs = [p[0] for p in WAVE_PTS]
ys = [p[1] for p in WAVE_PTS]
pchip = PchipInterpolator(xs, ys)
xx = np.linspace(xs[0], xs[-1], 600)
yy = pchip(xx)
TOP_Y = 150.0                  # 多边形上边（低于外环顶缘，配合 zorder 不碰环）
verts = list(zip(xx, yy)) + [(xs[-1], TOP_Y), (xs[0], TOP_Y)]
codes = [Path.MOVETO] + [Path.LINETO] * (len(verts) - 1)
ax.add_patch(PathPatch(Path(verts, codes), facecolor=BG, edgecolor="none", zorder=2))

# 3) 外圈椭圆环
ax.add_patch(Ellipse((RX, RY), 2 * RMX, 2 * RMY, fill=False,
                     edgecolor=BLUE, linewidth=RING_LW, zorder=3))

# 4) 文字
ax.text(69, 93, "本科及以上学历员占比", fontsize=38.5, color="#FFFFFF",
        fontweight="bold", ha="left", va="baseline", zorder=4)
ax.text(70, 146.5, "6月30日最新统计数据本科及以上员工占比65%", fontsize=26.9,
        color="#FFFFFF", ha="left", va="baseline", zorder=4)
ax.text(575, 459, "%d%%" % PCT, fontsize=65.0, color="#FFFFFF",
        ha="center", va="baseline", zorder=4)
ax.text(69, 751.5, "*注：数据来源于公司人力资源系统", fontsize=15.3,
        color="#D9D9D9", ha="left", va="baseline", zorder=4)

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "16 波浪水球图.png"), dpi=DPI, facecolor=BG)
plt.show()
