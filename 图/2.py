# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 12:45:16 2026

@author: fjwsm
"""
import os
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

plt.rcParams["font.sans-serif"] = ["Microsoft YaHei"]
plt.rcParams["axes.unicode_minus"] = False

# ---------------- 画布与颜色 ----------------
W, H, DPI = 1088, 814, 100
BG     = "#1A1E43"   # 背景深蓝
BLUE   = "#0070C0"   # 柱
YELLOW = "#FFC000"   # 均值线 / 平均值标签
WHITE  = "#F2F2F2"   # 数值 / 类目
BASE_C = "#86889A"   # 基线 (50% 透明白叠底色的等效色)
NOTE_C = "#D9D9D9"   # 脚注

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------- 数据 ----------------
# (类目, 销量, 柱x0, 数值ink_x0, 数值ink_y0, 类目ink_x0)
DATA = [
    ("华北", 2354, 121, 124, 355, 123),
    ("华南", 1902, 285, 289, 411, 287),
    ("东北", 3524, 449, 452, 213, 451),
    ("西北", 2698, 613, 616, 313, 615),
    ("西南", 2896, 777, 780, 289, 779),
    ("华东", 2563, 941, 947, 370, 943),
]
BAR_W  = 51     # 柱宽 (全部一致)
BASE_Y = 678    # 柱底 (基线上缘, 基线 y678..679)
CAT_Y  = 697    # 类目 ink 顶 y (6 块实测一致)

# 柱顶标定公式: top = 679.46 - 0.1221 * value (残差 <= 0.5px)
def top_of(v):
    return 679.46 - 0.1221 * v

MEAN_VAL = 2656.16666666667    # xlsx D列 =AVERAGE($C$3:$C$8), 显示 2656

# 均值线: 实测 y353..356 (4px 厚), x 144..967
# 线中心 355 ≈ top_of(2656.17) = 355.14, 与数据自洽
LINE_X0, LINE_X1, LINE_Y, LINE_H = 144, 967, 353, 4
# 黄标签 ink 左上角 (黄 mask 实测 x884..1016 y309..327)
YLAB_X, YLAB_Y = 884, 310

# 字号(二分标定): 标题 38.48 粗 / 副标题 26.87 / 脚注 15.36(与图4同款)
# 数值 14.36 / 类目 16.8 / 平均值标签 15.27
FS_TITLE, FS_SUB, FS_NOTE = 38.48, 26.87, 15.36
FS_VAL, FS_CAT, FS_YLAB   = 14.36, 16.8, 15.27

# ---------------- 画布 ----------------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI)
fig.patch.set_facecolor(BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ---------------- 基线 (2px, x64..1048) ----------------
ax.add_patch(Rectangle((64, 678), 1048 - 64 + 1, 2,
                       facecolor=BASE_C, edgecolor="none", zorder=2))

# ---------------- 柱 + 数值 + 类目 ----------------
for (cat, v, x0, vx, vy, cx) in DATA:
    t = top_of(v)
    ax.add_patch(Rectangle((x0, t), BAR_W, BASE_Y - t,
                           facecolor=BLUE, edgecolor="none", zorder=2))
    # 柱顶白色销量 (dx=1 dy=1 补偿; 华东 vy=370 叠柱顶, zorder 高于柱)
    ax.text(vx - 1, vy - 1, str(v), fontsize=FS_VAL, color=WHITE,
            ha="left", va="top", zorder=5)
    # 类目 (dx=0 dy=-1 补偿)
    ax.text(cx, CAT_Y + 1, cat, fontsize=FS_CAT, color=WHITE,
            ha="left", va="top", zorder=5)

# ---------------- 均值线 (画在柱上层, 实测 a[354,470] 纯黄穿过柱) ----------------
ax.add_patch(Rectangle((LINE_X0, LINE_Y), LINE_X1 - LINE_X0 + 1, LINE_H,
                       facecolor=YELLOW, edgecolor="none", zorder=3))
# 平均值标签 (dx=0 dy=-1 补偿)
ax.text(YLAB_X, YLAB_Y, "平均值：2656", fontsize=FS_YLAB, color=YELLOW,
        ha="left", va="top", zorder=5)

# ---------------- 标题 / 副标题 / 脚注 ----------------
ax.text(72, 48, "3月各区域销量分布",
        fontsize=FS_TITLE, color="#FFFFFF", fontweight="bold",
        ha="left", va="top", zorder=5)
ax.text(72, 127, "东北销量最多占比总销量的22%，华南销量最低",
        fontsize=FS_SUB, color="#FFFFFF",
        ha="left", va="top", zorder=5)
ax.text(66, 757, "*注：数据来源于公司销售系统，统计日期截至2022.03.31",
        fontsize=FS_NOTE, color=NOTE_C,
        ha="left", va="top", zorder=5)

# ---------------- 保存 ----------------
out_path = os.path.join(OUT_DIR, "2 带均值柱形图.png")
fig.savefig(out_path, facecolor=BG, dpi=DPI)
print("saved:", out_path)
plt.show()
