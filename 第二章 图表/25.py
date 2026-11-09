# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:03:31 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""25 子弹图 复现（尺寸/坐标/配色/字号均为参考图像素级实测值）"""
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\25 子弹图.xlsx"
raw = pd.read_excel(XLSX, sheet_name="25 子弹图", header=None)
hdr = next(i for i, r in raw.iterrows() if any(str(v).strip() == "商品" for v in r.values))
df = raw.iloc[hdr + 1:, 1:7].copy()
df.columns = ["商品", "实际", "目标", "及格", "良好", "优秀"]
df = df.dropna(subset=["商品"]).reset_index(drop=True)
for c in ["实际", "目标", "及格", "良好", "优秀"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_PASS = "#82ADD7"      # 及格（浅蓝灰）
C_GOOD = "#0885BC"      # 良好（青）
C_EXC = "#0070C0"       # 优秀（中蓝）
C_ACT = "#0E5DFF"       # 实际（亮蓝，细柱）
C_TGT = "#FFC000"       # 目标（金色横线）
C_AXIS = "#454866"      # 底部轴线
C_TITLE = C_SUB = "#FFFFFF"
C_YLAB = "#F2F2F2"      # 左侧刻度数字
C_XLAB = C_LEG = "#B2B2BC"   # 商品名 / 图例文字
C_NOTE = "#D9D9D9"      # 脚注

FS_TITLE, FS_SUB = 34.6, 21.1
FS_YLAB, FS_XLAB, FS_LEG, FS_NOTE = 14.8, 16.3, 16.3, 15.4

W, H, DPI = 1091, 756, 100            # 画布 1091 × 756 px

X0, PITCH = 214.0, 149.7              # 第 1 组中心 / 组间距
BW_STACK, BW_ACT = 43.0, 25.0         # 背景堆叠柱宽 / 实际柱宽
BASE, K = 629.0, 2.9414               # 数值轴基线像素行 / 每像素单位数
TGT_LW = 4.0                          # 目标横线粗细（pt）
AXIS_X0, AXIS_X1, AXIS_Y = 140, 1038, 629.5

LEG_Y, LEG_W, LEG_H = 217, 16, 15     # 图例色块
LEGEND = [                            # (色块 x, 文字 x, 颜色, 文字, 是否目标横线)
    (187, 214, C_PASS, "及格", False),
    (274, 302, C_GOOD, "良好", False),
    (360, 385, C_EXC, "优秀", False),
    (446, 472, C_ACT, "实际", False),
    (533, 561, C_TGT, "目标", True),
]
YTICKS = [0, 200, 400, 600, 800, 1000, 1200]


def yp(v):
    """数值 -> 像素行"""
    return BASE - v / K


# ---------- 3. 像素坐标系画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                     # y 向下，与图片行号一致
ax.axis("off")

n = len(df)
centers = [X0 + PITCH * i for i in range(n)]

# ---------- 4. 底部轴线 ----------
ax.plot([AXIS_X0, AXIS_X1], [AXIS_Y, AXIS_Y], color=C_AXIS, lw=1.44,
        solid_capstyle="butt", zorder=1)

# ---------- 5. 逐商品绘制子弹图 ----------
for i, r in df.iterrows():
    cx = centers[i]
    # 5.1 背景堆叠柱：及格 -> 良好 -> 优秀（自下而上累加）
    cum = 0.0
    for val, color in ((r["及格"], C_PASS), (r["良好"], C_GOOD), (r["优秀"], C_EXC)):
        ax.add_patch(Rectangle((cx - BW_STACK / 2, yp(cum + val)), BW_STACK, val / K,
                               facecolor=color, edgecolor="none", zorder=2))
        cum += val
    # 5.2 实际值细柱（叠在背景柱之上）
    ax.add_patch(Rectangle((cx - BW_ACT / 2, yp(r["实际"])), BW_ACT,
                           r["实际"] / K, facecolor=C_ACT, edgecolor="none", zorder=3))
    # 5.3 目标横线（最上层，宽度同实际柱）
    ax.plot([cx - BW_ACT / 2, cx + BW_ACT / 2], [yp(r["目标"])] * 2, color=C_TGT,
            lw=TGT_LW, solid_capstyle="butt", zorder=4)

# ---------- 6. 左侧刻度 ----------
for v in YTICKS:
    ax.text(122.5, yp(v) + 2.0, str(v), ha="right", va="center",
            fontsize=FS_YLAB, color=C_YLAB, zorder=5)

# ---------- 7. 商品名 ----------
for i, r in df.iterrows():
    ax.text(centers[i], 665, str(r["商品"]), ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_XLAB, zorder=5)

# ---------- 8. 图例 ----------
for sx, tx, color, txt, is_dash in LEGEND:
    if is_dash:
        ax.add_patch(Rectangle((sx, 222), 15.5, 4.5, facecolor=color,
                               edgecolor="none", zorder=5))
    else:
        ax.add_patch(Rectangle((sx, LEG_Y), LEG_W, LEG_H, facecolor=color,
                               edgecolor="none", zorder=5))
    ax.text(tx, 231, txt, ha="left", va="baseline", fontsize=FS_LEG,
            color=C_LEG, zorder=5)

# ---------- 9. 标题 / 副标题 / 脚注 ----------
fig.text(0.06178, 1 - 74 / 756, "2022年上半年各商品销量完成情况",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(0.06269, 1 - 119.5 / 756, "防晒整体销量最好，达到856，面霜远超目标，超额完成30%",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(0.06173, 1 - 718 / 756, "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

# ---------- 10. 输出 ----------
fig.savefig("25.png", dpi=DPI, facecolor=BG)
plt.show()
