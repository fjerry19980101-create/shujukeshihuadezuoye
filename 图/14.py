# -*- coding: utf-8 -*-

import math
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge
from matplotlib import font_manager as fm
import numpy as np
import pandas as pd

# ============================================================
# 0. 数据：从 xlsx 读取（同时给出兜底值，保证脚本单独可跑）
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\14 单值圆环图.xlsx"

PCT = 0.85           # 目标完成率
TITLE = "2022年上半年目标完成率"
SUBTITLE = "截至6月30日销售目标总体完成率达到85%"
LABEL = "目标完成率"
FOOTNOTE = "*注：数据来源于公司销售系统"

def load_data(path):
    """从工作簿里读出完成率与文案；失败则用兜底常量。"""
    val, texts = PCT, {}
    try:
        xl = pd.ExcelFile(path)
        for sheet in xl.sheet_names:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            for _, row in raw.iterrows():
                cells = [str(v) for v in row.values if pd.notna(v)]
                for c in cells:
                    s = c.strip()
                    if s.endswith("%") and texts.get("pct") is None:
                        try:
                            v = float(s.rstrip("%")) / 100.0
                            if 0 < v <= 1:
                                val = v
                                texts["pct"] = s
                        except ValueError:
                            pass
                    if "目标完成率" in s and len(s) > 6 and "title" not in texts:
                        texts["title"] = s
                    if s.startswith("截至") and "sub" not in texts:
                        texts["sub"] = s
                    if s.startswith("*注") and "note" not in texts:
                        texts["note"] = s
    except Exception as e:
        print("[warn] 读取 xlsx 失败，使用内置文案:", e)
    return val, texts

_pct, _t = load_data(XLSX)
if _pct != PCT:
    PCT = _pct
TITLE = _t.get("title", TITLE)
SUBTITLE = _t.get("sub", SUBTITLE)
FOOTNOTE = _t.get("note", FOOTNOTE)
print("PCT = %.4f  TITLE=%s" % (PCT, TITLE))

# ============================================================
# 1. 画布与颜色（全部为参考图实测值）
# ============================================================
W, H = 1088, 792                 # 画布像素
DPI = 100
BG = "#1A1E43"                   # 背景 #1A1E43 (26,30,67)
GRAY = "#303455"                 # 灰色余量段 (48,52,85)

# 渐变两端（色环实测）
C_START = (164, 61, 155)         # 300° 处（位于9-10点钟方向一带）→ 起点端
C_TOP   = (117, 49, 160)         # 12点方向 最深紫
C_END   = (230, 78, 108)         # 6点方向  玫红

# 圆心 / 半径
CX, CY = 553.1, 442.4
R_OUT = 223.5                    # 外半径
R_IN = 201.5                     # 内半径
BAND = R_OUT - R_IN              # 带厚 = 22.0
R_MID = (R_OUT + R_IN) / 2.0

# 角度（正上方=0°，顺时针为正）
TH_START = 300.0                 # 渐变弧起点
SWEEP = 360.0 * PCT              # 弧长（按占比）
TH_END = TH_START + SWEEP        # 渐变弧终点 = 606°
GRAY_A, GRAY_B = 246.2, 300.0    # 灰色段 246.2° → 300.0°

print("SWEEP = %.4f deg   arc %.2f -> %.2f" % (SWEEP, TH_START, TH_END))

# 字体
FP_REG = fm.FontProperties(fname="C:/Windows/Fonts/msyh.ttc")     # 微软雅黑 常规
FP_BOLD = fm.FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc", weight="bold")  # 粗体

# 字号（回归实测标定）
FS_TITLE = 34.6
FS_SUB = 23.0
FS_PCT = 58.0
FS_LABEL = 20.8
FS_NOTE = 15.3

# ============================================================
# 2. 建立「像素坐标系」画布：数据坐标 == 图片像素坐标，y 向下
# ============================================================
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                # y 向下，行号即 y
ax.axis("off")

# ============================================================
# 3. 渐变圆弧：颜色只与「离 12 点方向的角距」有关（关于竖直轴完全对称）
#    实测：300°与60°同色、330°与30°同色（diff=0）——即 PowerPoint 式
#    「沿竖直方向的线性渐变」被镜像到整圈上。
#    参数 x 采用「竖直归一化位置」: x = (1 - cos(角距))/2  ∈ [0(12点), 1(6点)]
#    颜色用参考图实测的关键色标做分段线性插值（逐 10° 采样，最贴近原图）。
# ============================================================
# 关键色标（角距 deg -> RGB），由参考图 r=212 处逐 5° 实测得到
GRAD_STOPS = [
    (0,   (117, 49, 160)),
    (10,  (119, 50, 159)),
    (20,  (123, 51, 159)),
    (30,  (130, 53, 159)),
    (40,  (140, 55, 158)),
    (50,  (151, 58, 157)),
    (60,  (164, 61, 155)),
    (70,  (175, 64, 152)),
    (80,  (187, 67, 149)),
    (90,  (199, 70, 145)),
    (100, (206, 72, 139)),
    (110, (214, 74, 134)),
    (120, (221, 76, 129)),
    (130, (224, 76, 123)),
    (140, (226, 77, 118)),
    (150, (228, 77, 113)),
    (160, (229, 78, 110)),
    (170, (230, 78, 108)),
    (180, (230, 78, 108)),
]

def ang_color(phi):
    """
    屏幕角度 phi（正上=0°，顺时针）-> 归一化颜色。
    取「与 12 点的最小角距」查色标，实现竖直镜像渐变。
    """
    d = phi % 360.0
    if d > 180.0:
        d = 360.0 - d           # 折到 0..180（角距）
    # 在色标里线性插值
    for i in range(len(GRAD_STOPS) - 1):
        d0, c0 = GRAD_STOPS[i]
        d1, c1 = GRAD_STOPS[i + 1]
        if d0 <= d <= d1:
            t = 0.0 if d1 == d0 else (d - d0) / (d1 - d0)
            rgb = [(1 - t) * c0[k] + t * c1[k] for k in range(3)]
            return tuple(c / 255.0 for c in rgb)
    c = GRAD_STOPS[-1][1]
    return tuple(v / 255.0 for v in c)

# 分段数：太多会在每条接缝处产生抗锯齿叠加、把整圈压暗
# （实测 1440 段时 r=212 处颜色从 (117,49,160) 掉到 (95,44,137)；
#   36~180 段则为精确值）。这里用 180 段 + 关闭抗锯齿，兼顾平滑与准确。
N_SEG = 180
for i in range(N_SEG):
    p0 = i / N_SEG
    p1 = (i + 1) / N_SEG
    phi0 = TH_START + SWEEP * p0
    phi1 = TH_START + SWEEP * p1
    phi_mid = TH_START + SWEEP * (p0 + p1) / 2.0
    col = ang_color(phi_mid)
    t0 = (270.0 + phi0) % 360.0
    t1 = (270.0 + phi1) % 360.0
    if t1 <= t0:
        t1 += 360.0
    ax.add_patch(Wedge((CX, CY), R_OUT, t0, t1,
                       width=BAND, facecolor=col, edgecolor="none",
                       antialiased=False))

# 灰色余量段（246.2° → 300.0°）
t0 = (270.0 + GRAY_A) % 360.0
t1 = (270.0 + GRAY_B) % 360.0
if t1 <= t0:
    t1 += 360.0
ax.add_patch(Wedge((CX, CY), R_OUT, t0, t1,
                   width=BAND, facecolor=GRAY, edgecolor="none"))

# ============================================================
# 4. 文字：fig.text() 的 y 从画布底部起算，需换算 y_fig = (H - y_px)/H
# ============================================================
def put(x_px, y_baseline_px, s, size, fp, color, ha="left"):
    ax.text(x_px, y_baseline_px, s, fontsize=size, color=color,
            ha=ha, va="baseline", fontproperties=fp, zorder=10)

# 标题：墨迹 y 50..96（baseline ≈ 底行 - 0.13*高 ≈ 90），x0 = 68
put(65, 90, TITLE, FS_TITLE, FP_BOLD, "white")
# 副标题：墨迹 y 119..149（CJK，baseline ≈ 底行 - 3 ≈ 146），x0 = 66
put(65, 145, SUBTITLE, FS_SUB, FP_REG, "white")
# 85%：墨迹 y 382..445，数字 baseline ≈ 底行 + 1 ≈ 446
ax.text(CX - 3, 446, "85%", fontsize=FS_PCT, color="white", ha="center",
        va="baseline", fontproperties=FP_REG, zorder=10)
# 目标完成率：墨迹 y 510..535（CJK，baseline ≈ 底行 - 3 ≈ 532）
ax.text(CX - 5.5, 533, LABEL, fontsize=FS_LABEL, color="#F2F2F2", ha="center",
        va="baseline", fontproperties=FP_REG, zorder=10)
# 脚注：墨迹 y 734..754（CJK，baseline ≈ 底行 - 3 ≈ 751），x0 = 61
put(59, 752, FOOTNOTE, FS_NOTE, FP_REG, "#D9D9D9")

# ============================================================
# 5. 输出
# ============================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "14 单值圆环图.png")
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.close(fig)
print("saved ->", OUT)
