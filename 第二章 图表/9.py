# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 10:58:41 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Polygon

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------------- 画布与颜色 ----------------
W, H, DPI = 1088, 828, 100
BG     = "#1A1E43"   # 背景深蓝
BLUE   = "#0070C0"   # 2021 柱 / 延伸线
LBLUE  = "#82ADD7"   # 2022 柱
WHITE  = "#F2F2F2"   # 数值 / 箭头 / 图例数字
GREY   = "#B2B2BC"   # 类目 / 图例汉字
NOTE_C = "#D9D9D9"   # 脚注
BASE_C = "#717389"   # 底部基线

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------- 数据 (商品, 2021销量, 2022销量, 差值, 2021柱x范围, 2022柱x范围) ----------------
DATA = [
    ("口红", 3568, 2569,  999, (116, 156), (169, 209)),
    ("面膜", 4135, 3241,  894, (302, 342), (355, 395)),
    ("隔离", 4436, 2965, 1471, (488, 529), (541, 583)),
    ("防晒", 4106, 3209,  897, (674, 715), (727, 768)),
    ("精华", 4936, 3541, 1395, (860, 901), (913, 954)),
]

# 柱顶标定公式: top = 694.78 - 0.07576 * value (残差 <= 0.7px)
def top_of(v):
    return 694.78 - 0.07576 * v

BASE_Y   = 692          # 柱底(基线上缘)
BASELINE = [70, 999]    # 基线 x 范围, y=693 处 2px

# 2021 柱顶延伸线: 逐组实测 y 与右端点
EXT_Y  = [423, 381, 358, 383, 321]
EXT_X1 = [193, 380, 565, 752, 938]

# 白色箭头: 三角头顶点 y / 中心 x / 杆底 y (逐组实测)
AR_CX = [188, 374.5, 561, 746.5, 933]
AR_TY = [425, 383, 360, 385, 322]
AR_BOT = [498, 448, 468, 450, 425]

# 数值锚点 (ha=center, va=top): 2021值(柱顶上方) / 2022值(柱内) / 差值(箭头右侧)
V21_POS = [(136, 387.5), (321, 344.5), (507.5, 322.5), (693.5, 346.5), (880, 284.5)]
V22_POS = [(189, 520.5), (374.5, 470.5), (560.5, 491.5), (747, 472.5), (933, 447.5)]
DIF_POS = [(219, 467.5), (405, 413.5), (594, 421.5), (778, 418.5), (976.5, 376.5)]

# 类目锚点 (ha=center, va=top)
CAT_POS = [(163, 712), (347.5, 711), (535, 711), (721, 711), (906.5, 711)]

FS_NUM  = 14.2    # 柱顶/柱内/差值数值
FS_CAT  = 16.3    # 类目与图例汉字
FS_LEG  = 15.8    # 图例数字

# ---------------- 画布 ----------------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)
fig.patch.set_facecolor(BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ---------------- 基线 ----------------
ax.plot(BASELINE, [693, 693], color=BASE_C, lw=2,
        solid_capstyle="butt", zorder=2)

# ---------------- 柱 / 延伸线 / 数值 / 类目 ----------------
for i, (cat, v1, v2, dv, (xa0, xa1), (xb0, xb1)) in enumerate(DATA):
    t1, t2 = top_of(v1), top_of(v2)
    # 2021 柱
    ax.add_patch(Rectangle((xa0, t1), xa1 - xa0 + 1, BASE_Y - t1,
                           facecolor=BLUE, edgecolor="none", zorder=2))
    # 2022 柱
    ax.add_patch(Rectangle((xb0, t2), xb1 - xb0 + 1, BASE_Y - t2,
                           facecolor=LBLUE, edgecolor="none", zorder=2))
    # 2021 柱顶延伸线
    ax.plot([xa1, EXT_X1[i]], [EXT_Y[i], EXT_Y[i]], color=BLUE, lw=1.2,
            solid_capstyle="butt", zorder=3)
    # 白色箭头 (三角头 + 竖杆)
    cx, ty = AR_CX[i], AR_TY[i]
    ax.add_patch(Polygon([(cx, ty), (cx - 7.5, ty + 15), (cx + 7.5, ty + 15)],
                         closed=True, facecolor=WHITE, edgecolor="none", zorder=4))
    ax.plot([cx, cx], [ty + 15, AR_BOT[i]], color=WHITE, lw=1.5,
            solid_capstyle="butt", zorder=4)
    # 数值: 2021(柱顶上方) / 2022(柱内上部) / 差值(箭头右侧)
    ax.text(*V21_POS[i], str(v1), fontsize=FS_NUM, color=WHITE,
            ha="center", va="top", zorder=5)
    ax.text(*V22_POS[i], str(v2), fontsize=FS_NUM, color=WHITE,
            ha="center", va="top", zorder=5)
    ax.text(*DIF_POS[i], str(dv), fontsize=FS_NUM, color=WHITE,
            ha="center", va="top", zorder=5)
    # 类目
    ax.text(*CAT_POS[i], cat, fontsize=FS_CAT, color=GREY,
            ha="center", va="top", zorder=5)

# ---------------- 图例 ----------------
ax.add_patch(Rectangle((73, 235), 16, 15, facecolor=BLUE, edgecolor="none", zorder=4))
ax.text(120.5, 235, "2021", fontsize=FS_LEG, color=WHITE,
        ha="center", va="top", zorder=5)
ax.text(168.5, 232, "销量", fontsize=FS_CAT, color=GREY,
        ha="center", va="top", zorder=5)
ax.add_patch(Rectangle((212, 235), 15, 15, facecolor=LBLUE, edgecolor="none", zorder=4))
ax.text(258.5, 235, "2022", fontsize=FS_LEG, color=WHITE,
        ha="center", va="top", zorder=5)
ax.text(307.5, 232, "销量", fontsize=FS_CAT, color=GREY,
        ha="center", va="top", zorder=5)

# ---------------- 标题 / 副标题 / 脚注 ----------------
ax.text(63.5, 55, "2022年商品对比去年销售情况",
        fontsize=38.3, color="#FFFFFF", fontweight="bold",
        ha="left", va="top", zorder=5)
ax.text(63.5, 119, "商品整体比去年销量有所下降，其中隔离下降最多，下降33%",
        fontsize=23.0, color="#FFFFFF",
        ha="left", va="top", zorder=5)
ax.text(51.5, 770, "*注：数据来源于公司销售系统，统计日期截至2022.01.01",
        fontsize=15.27, color=NOTE_C,
        ha="left", va="top", zorder=5)

# ---------------- 保存 ----------------
out_path = os.path.join(OUT_DIR, "9 对比柱形图.png")
fig.savefig(out_path, facecolor=BG, dpi=DPI)
print("saved:", out_path)
plt.show()
