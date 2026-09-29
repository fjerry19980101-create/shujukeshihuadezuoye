# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 12:45:14 2026

@author: fjwsm
"""
import os
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle, FancyBboxPatch

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------------- 画布与颜色 ----------------
W, H, DPI = 1088, 792, 100
BG     = "#1A1E43"   # 背景深蓝
GRID_C = "#454866"   # 虚线网格
STEM_C = "#4874CB"   # 竖线/圆点真色(气泡内被罩成 #3A73C8)
BUB_FC = "#026ABA"   # 气泡真色 (alpha 0.21 叠背景 = #152E5C)
BUB_AL = 0.21
WHITE  = "#F2F2F2"   # 刻度 / 类目 / 气泡数字
NOTE_C = "#D9D9D9"   # 脚注

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------- 数据 ----------------
# (类目, 销量, 柱x0, 柱宽, 类目ink_x0, 气泡x0, 气泡y0, 字ink_x0, 字ink_y0)
DATA = [
    ("口红", 653, 207, 23, 198, 197, 361, 232, 368),
    ("面膜", 523, 350, 23, 338, 344, 415, 377, 427),
    ("隔离", 648, 493, 24, 482, 486, 370, 518, 380),
    ("防晒", 856, 636, 24, 626, 633, 300, 661, 311),
    ("精华", 714, 779, 24, 768, 777, 350, 812, 359),
    ("面霜", 785, 922, 24, 911, 916, 327, 950, 337),
]

BAR_BOT   = 655     # 柱底 (0 网格线 y655.5 上缘)
CORNER_R  = 6       # 柱顶圆角半径 (实测顶行宽 = 柱宽-2r)
BUB_W, BUB_H, BUB_R = 79, 39, 4.5
DOT_R, STEM_W = 8.5, 3
DOT_DY    = 22      # 圆点圆心 = 气泡顶 + 22

# 柱顶标定公式: top = 656.02 - 0.3542 * value (残差 <= 0.5px)
def top_of(v):
    return 656.02 - 0.3542 * v

# 渐变表 (面霜柱中线逐行实测, 相对位置 -> RGB)
GT = [
    (0.000, 0x00, 0xF3, 0xFD), (0.047, 0x00, 0xF1, 0xFC),
    (0.095, 0x00, 0xEF, 0xFA), (0.142, 0x00, 0xEC, 0xF8),
    (0.190, 0x00, 0xE9, 0xF6), (0.237, 0x00, 0xE4, 0xF3),
    (0.285, 0x00, 0xDE, 0xEF), (0.332, 0x00, 0xD8, 0xEB),
    (0.380, 0x00, 0xD0, 0xE6), (0.427, 0x00, 0xC6, 0xE0),
    (0.474, 0x00, 0xBB, 0xD9), (0.522, 0x00, 0xB1, 0xD3),
    (0.569, 0x00, 0xA4, 0xCC), (0.617, 0x00, 0x99, 0xC6),
    (0.664, 0x00, 0x8C, 0xC0), (0.712, 0x00, 0x80, 0xBA),
    (0.759, 0x00, 0x74, 0xB5), (0.807, 0x00, 0x6A, 0xB0),
    (0.854, 0x00, 0x61, 0xAD), (0.901, 0x00, 0x58, 0xAB),
    (0.949, 0x00, 0x52, 0xA8), (1.000, 0x00, 0x4D, 0xA7),
]
GT_REL = [g[0] for g in GT]
GT_RGB = np.array([[g[1], g[2], g[3]] for g in GT], float)

# 网格: 7 条虚线 (16on/6off, 最下一条即 0 值线)
GRID_YS = [231.5, 301.5, 372.5, 443.5, 514.5, 585.5, 655.5]
GRID_X  = (147, 1005)

# Y 刻度 (ha=left, ink 左上角实测)
YTICKS = [
    (86,  224, "1200"), (86,  295, "1000"), (97, 366, "800"),
    (97,  436, "600"),  (96,  507, "400"),  (97, 578, "200"),
    (120, 649, "0"),
]
CAT_Y = 674    # 类目 ink 顶 y

# 字号(二分标定): 标题 38.30 粗 / 副标题 26.87 / 脚注 15.27
# Y刻度 14.30 / 类目 16.37 / 气泡数字 14.2
FS_TITLE, FS_SUB, FS_NOTE = 38.30, 26.87, 15.27
FS_YT, FS_CAT, FS_BUB     = 14.30, 16.37, 14.2

# ---------------- 画布 ----------------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)
fig.patch.set_facecolor(BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ---------------- 虚线网格 ----------------
for gy in GRID_YS:
    ax.plot(list(GRID_X), [gy, gy], color=GRID_C, lw=2,
            linestyle=(0, (8, 3)), solid_capstyle="butt", zorder=1)

# ---------------- 柱(逐行渐变+顶圆角) + 标注系统 + 类目 ----------------
for (cat, v, bx0, bw, cx, ux, uy, tx, ty) in DATA:
    bt = int(round(top_of(v)))
    cxm = bx0 + bw / 2.0
    # 逐行渐变柱 (圆角行按圆弧缩进)
    span = BAR_BOT - bt
    for y in range(bt, BAR_BOT):
        d = y - bt
        if d < CORNER_R:
            ind = CORNER_R - (CORNER_R ** 2 - (CORNER_R - d) ** 2) ** 0.5
        else:
            ind = 0.0
        rel = d / float(span)
        rgb = [float(np.interp(rel, GT_REL, GT_RGB[:, ch])) for ch in range(3)]
        col = "#%02X%02X%02X" % tuple(int(round(c)) for c in rgb)
        ax.add_patch(Rectangle((bx0 + ind, y), bw - 2 * ind, 1.0,
                               facecolor=col, edgecolor="none", zorder=2))
    # 竖线 (从圆点圆心到柱顶上方 1px; 气泡内段将被半透明气泡罩色)
    dot_cy = uy + DOT_DY
    ax.add_patch(Rectangle((cxm - STEM_W / 2.0, dot_cy), STEM_W,
                           (bt - 1) - dot_cy,
                           facecolor=STEM_C, edgecolor="none", zorder=3))
    # 圆点 (在气泡层之下, 被罩后呈 #3A73C8)
    ax.add_patch(Circle((cxm, dot_cy), DOT_R,
                        facecolor=STEM_C, edgecolor="none", zorder=3))
    # 半透明气泡 (罩住圆点/竖线/网格, 叠加等效 #152E5C)
    ax.add_patch(FancyBboxPatch((ux, uy), BUB_W, BUB_H,
                                boxstyle="round,pad=0,rounding_size=%g" % BUB_R,
                                facecolor=BUB_FC, alpha=BUB_AL,
                                edgecolor="none", zorder=4))
    # 气泡内白色销量 (dx=1 dy=1 补偿)
    ax.text(tx - 1, ty - 1, str(v), fontsize=FS_BUB, color=WHITE,
            ha="left", va="top", zorder=5)
    # 类目 (dx=2 dy=0 补偿)
    ax.text(cx - 2, CAT_Y, cat, fontsize=FS_CAT, color=WHITE,
            ha="left", va="top", zorder=5)

# ---------------- Y 刻度 (dx=2 dy=1 补偿) ----------------
for x0, y0, s in YTICKS:
    ax.text(x0 - 2, y0 - 1, s, fontsize=FS_YT, color=WHITE,
            ha="left", va="top", zorder=5)

# ---------------- 标题 / 副标题 / 脚注 ----------------
ax.text(76, 42, "3月商品销量对比",
        fontsize=FS_TITLE, color="#FFFFFF", fontweight="bold",
        ha="left", va="top", zorder=5)
ax.text(76, 119, "防晒销量最多，3月销量856；面膜最少，3月销量523",
        fontsize=FS_SUB, color="#FFFFFF",
        ha="left", va="top", zorder=5)
ax.text(67, 736, "*注：数据来源于公司销售系统，统计日期截至2022.03.31",
        fontsize=FS_NOTE, color=NOTE_C,
        ha="left", va="top", zorder=5)

# ---------------- 保存 ----------------
out_path = os.path.join(OUT_DIR, "3 渐变圆角柱形图.png")
fig.savefig(out_path, facecolor=BG, dpi=DPI)
print("saved:", out_path)
plt.show()
