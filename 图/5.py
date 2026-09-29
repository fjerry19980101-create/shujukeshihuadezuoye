# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 12:45:16 2026

@author: fjwsm
"""
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 画布与颜色（实测） ----------
W, H = 1088, 756
BG      = "#1A1E43"   # 背景 (26,30,67)
C_BLUE  = "#0070C0"   # 销售额柱 (0,112,192)
C_RED   = "#E74E69"   # 利润额柱 (231,78,105)
C_AX    = "#454866"   # 轴线 (69,72,102)
C_TXT   = "#F2F2F2"   # 数值/刻度/图例文字 (242,242,242)
C_FOOT  = "#D9D9D9"   # 脚注 (217,217,217)
C_TITLE = "#FFFFFF"   # 标题/副标题 (255,255,255)
# 柱底亮边(第638行) = 80%柱色+20%白, 逐像素实测:
CAP_BLUE = (48/255, 138/255, 202/255)
CAP_RED  = (233/255, 110/255, 132/255)

FONT = "Microsoft YaHei"   # 微软雅黑（逐个 text 显式指定）

OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "5 层叠柱形图.png")

# ---------- 数据（来自 xlsx） ----------
cats   = ["2021Q1", "Q2", "Q3", "Q4", "2022Q1", "Q2"]
sales  = [3121, 4086, 4321, 4601, 4936, 4231]   # 销售额 -> 蓝柱
profit = [1020, 1421, 1502, 1623, 1781, 1432]   # 利润额 -> 红柱

# ---------- 几何参数（像素实测/标定） ----------
AXIS_Y = 638.0                 # 轴线条带 y 638-639 (2px), 柱压轴线上层
AX_X0, AX_W = 96.0, 880.0      # 轴线 x: 96 -> 976
BLUE_X0, PITCH_B = 137.76, 146.63   # 蓝柱左缘/间距
RED_X0,  PITCH_R = 164.14, 146.74   # 红柱左缘/间距 (红柱叠压蓝柱上层 ~11.6px)
BAR_W = 37.5
TOP_B = [410.5, 339.7, 323.4, 303.4, 278.4, 329.5]   # 蓝柱顶(渲染值, aa=False 就近取整)
TOP_R = [564.5, 534.5, 528.5, 519.5, 508.5, 533.5]   # 红柱顶(渲染值)
# 数值标签基线 = 参考图墨迹底 (末根红柱 1432 间距异常小, 按实测硬编码)
BASE_B = [389, 319, 301, 281, 256, 308]
BASE_R = [543, 514, 508, 499, 487, 520]

# 文字参数（渲染标定解出）
FS_TITLE, X_TITLE, BASE_TITLE = 38.4, 62.4, 96.6    # 标题(粗体, 半角括号紧排)
FS_SUB,   X_SUB,   BASE_SUB   = 23.0, 62.4, 144.1   # 副标题
FS_FOOT,  X_FOOT,  BASE_FOOT  = 15.3, 50.0, 719.2   # 脚注
FS_TICK,  BASE_TICK           = 15.4, 677.0         # 刻度
FS_VAL  = 16.2                                       # 数值
FS_LEG  = 15.2                                       # 图例文字
LEG_CX  = 922.5
LEG_BASE1, LEG_BASE2 = 410.0, 468.0

TITLE = "2021年至今季度销售额(万)和利润额(万)"
SUB   = "2022年第二季度销售额首次出现下降，降幅达到15%"
FOOT  = "*注：数据来源于公司销售系统，统计日期截至2022.06.30"

# ---------- 画布 ----------
fig = plt.figure(figsize=(W / 100, H / 100), dpi=100)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_xlim(0, W)
ax.set_ylim(H, 0)   # 数据坐标 = 画布像素坐标 (y 向下)
ax.axis("off")

def rect(x, y, w, h, color, z):
    ax.add_patch(Rectangle((x, y), w, h, facecolor=color, edgecolor="none",
                           antialiased=False, zorder=z))

rect(0, 0, W, H, BG, 0)                                    # 背景
rect(AX_X0, 638.0, AX_W, 2.0, C_AX, 1)                     # 轴线 (柱上层压盖)

# ---------- 柱体 (蓝 z2, 红 z3 叠压蓝柱) ----------
for i, v in enumerate(sales):
    x = BLUE_X0 + i * PITCH_B
    rect(x, TOP_B[i], BAR_W, AXIS_Y - TOP_B[i], C_BLUE, 2)
    rect(x, AXIS_Y, BAR_W, 1.0, CAP_BLUE, 2)               # 柱底亮边
for i, v in enumerate(profit):
    x = RED_X0 + i * PITCH_R
    rect(x, TOP_R[i], BAR_W, AXIS_Y - TOP_R[i], C_RED, 3)
    rect(x, AXIS_Y, BAR_W, 1.0, CAP_RED, 3)

# ---------- 图例: 两个不透明彩边框 (压在末根蓝柱上) ----------
def legend_box(x0, y0, x1, y1, color):
    rect(x0 + 2, y0 + 2, x1 - x0 - 4, y1 - y0 - 4, BG, 4)  # 内部填充(遮柱)
    rect(x0, y0, x1 - x0, 2, color, 5)                     # 上边
    rect(x0, y1 - 2, x1 - x0, 2, color, 5)                 # 下边
    rect(x0, y0, 2, y1 - y0, color, 5)                     # 左边
    rect(x1 - 2, y0, 2, y1 - y0, color, 5)                 # 右边

legend_box(868, 389, 977, 423, C_BLUE)   # 销售额框
legend_box(867, 446, 976, 480, C_RED)    # 利润额框

# ---------- 文字 ----------
def txt(x, y_base, s, fs, color=C_TXT, weight="normal", ha="center"):
    fig.text(x / W, (H - y_base) / H, s, fontsize=fs, color=color,
             ha=ha, va="baseline", fontweight=weight, fontfamily=FONT, zorder=6)

txt(X_TITLE, BASE_TITLE, TITLE, FS_TITLE, C_TITLE, "bold", "left")
txt(X_SUB,   BASE_SUB,   SUB,   FS_SUB,   C_TITLE, "normal", "left")
txt(X_FOOT,  BASE_FOOT,  FOOT,  FS_FOOT,  C_FOOT, "normal", "left")

for i, c in enumerate(cats):                       # 刻度: 居中于每槽
    txt(169.75 + i * 146.63, BASE_TICK, c, FS_TICK)
for i, v in enumerate(sales):                      # 数值: 居中于各自柱
    txt(BLUE_X0 + i * PITCH_B + BAR_W / 2, BASE_B[i], str(v), FS_VAL)
for i, v in enumerate(profit):
    txt(RED_X0 + i * PITCH_R + BAR_W / 2, BASE_R[i], str(v), FS_VAL)

txt(LEG_CX, LEG_BASE1, "销售额", FS_LEG)
txt(LEG_CX, LEG_BASE2, "利润额", FS_LEG)

fig.savefig(OUT, dpi=100)
print("saved:", OUT)
