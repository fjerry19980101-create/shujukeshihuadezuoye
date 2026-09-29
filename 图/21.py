# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:31:39 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""20 南丁格尔圆环图 复现（尺寸/角度/配色/字号均为参考图像素级实测值）"""
import os
import math
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge, Circle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\20 南丁格尔圆环图.xlsx"
RAW = pd.read_excel(XLSX, sheet_name="20 南丁格尔圆环图", header=None)
hdr = next(i for i, r in RAW.iterrows() if any(str(v).strip() == "年龄" for v in r.values))
df = RAW.iloc[hdr + 1:, 1:3].copy()
df.columns = ["年龄", "占比"]
df = df.dropna(subset=["年龄"]).reset_index(drop=True)
df["年龄"] = df["年龄"].astype(str).str.strip()
df["占比"] = pd.to_numeric(df["占比"], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_PURPLE = "#512DCB"    # [20,30)
C_BLUE = "#6073F3"      # [30,40)
C_TEAL = "#66CBDD"      # [40,50)
C_GOLD = "#F5C353"      # >=50
C_GRAY = "#A6A6A6"      # 引出线（灰）
C_TEXT = "#F2F2F2"      # 标题 / 副标题 / 脚注 / 类目文字

FS_TITLE, FS_SUB, FS_CALL, FS_NOTE = 32.5, 23.0, 14.3, 15.4

W, H, DPI = 1093, 796, 100            # 画布 1093 × 796 px

CX, CY = 544.5, 478.0                 # 圆心
R_HOLE = 138.5                        # 内孔半径
RING_H = 18.3                         # 每圈高度
N_RING = 5                            # 最多 5 圈
A0 = 0.0                              # 第 1 个扇区从正上方开始，顺时针

# 扇区配色 + 引出线（全部为实测像素坐标）
SERIES = [
    dict(name="[20,30)", color=C_PURPLE, txt_x=877, base1=244, base2=268,
         ul=(855, 282, 1006), bend_to=(809, 320), dot=(809, 320),
         gray=[(878, 247), (858, 247), (752, 388)]),
    dict(name="[30,40)", color=C_BLUE, txt_x=798, base1=668, base2=695,
         ul=(710, 720, 940), bend_to=(690, 673), dot=(690, 673),
         gray=[(612, 690), (790, 688)]),
    dict(name="[40,50)", color=C_TEAL, txt_x=172, base1=439, base2=463,
         ul=(103, 484, 290), bend_to=(316, 442), dot=(316, 442),
         gray=[(245, 441), (325, 449)]),
    dict(name=">=50", color=C_GOLD, txt_x=289, base1=195, base2=224,
         ul=(214, 242, 389), bend_to=(415, 287), dot=(415, 287),
         gray=[(277, 201), (368, 201), (460, 274)]),
]
DOT_R = 5.5                           # 引出圆点半径
LINE_LW = 2.9                         # 下划线/折线粗细（≈4px）
GRAY_LW = 1.4                         # 灰色引出线（≈2px）


def pol(r, adeg):
    """极坐标（正上方为 0°，顺时针为正）-> 画布像素坐标"""
    a = math.radians(adeg)
    return CX + r * math.sin(a), CY - r * math.cos(a)


def dark(hexcolor, f=0.85):
    """按比例压暗，用于圆环分隔线"""
    c = [int(hexcolor[i:i + 2], 16) for i in (1, 3, 5)]
    return "#%02X%02X%02X" % tuple(int(v * f) for v in c)


total = float(df["占比"].sum())
vmax = float(df["占比"].max())

# ---------- 3. 画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                     # y 向下，与图片行号一致
ax.axis("off")

# ---------- 4. 圆环扇区（径向高度 = 圈数，按占比排名） ----------
ang = A0
for i, row in df.iterrows():
    v = float(row["占比"])
    span = 360.0 * v / total
    # 圈数：与最大值成比例（实测 37.5%→5圈、29.2%→4圈、20.8%→3圈、12.5%→2圈）
    n = max(1, int(round(v / vmax * N_RING)))
    color = SERIES[i]["color"] if i < len(SERIES) else C_PURPLE
    r_out = R_HOLE + n * RING_H
    ax.add_patch(Wedge((CX, CY), r_out, ang - 90, ang + span - 90,
                       width=r_out - R_HOLE, facecolor=color,
                       edgecolor="none", zorder=3))
    # 圈间分隔弧线
    for k in range(1, n):
        rk = R_HOLE + k * RING_H
        ax.add_patch(Wedge((CX, CY), rk + 0.6, ang - 90, ang + span - 90,
                           width=1.2, facecolor=dark(color),
                           edgecolor="none", zorder=4))
    # 引出折线：下划线 + 斜折 + 圆点
    cf = SERIES[i] if i < len(SERIES) else None
    if cf:
        x0, y, x1 = cf["ul"]
        ax.plot([x0, x1], [y, y], color=color, lw=LINE_LW,
                solid_capstyle="butt", zorder=5)
        ax.plot([x1, cf["bend_to"][0]], [y, cf["bend_to"][1]], color=color,
                lw=LINE_LW, solid_capstyle="butt", zorder=5)
        ax.add_patch(Circle(cf["dot"], DOT_R, facecolor=color,
                            edgecolor="none", zorder=6))
        gx = [p[0] for p in cf["gray"]]
        gy = [p[1] for p in cf["gray"]]
        ax.plot(gx, gy, color=C_GRAY, lw=GRAY_LW, solid_capstyle="round",
                zorder=5)
        ax.annotate(cf["name"], (cf["txt_x"], cf["base1"]), ha="left",
                    va="baseline", fontsize=FS_CALL, color=C_TEXT, zorder=6)
        ax.annotate("%.1f%%" % (v * 100), (cf["txt_x"], cf["base2"]),
                    ha="left", va="baseline", fontsize=FS_CALL,
                    color=color, zorder=6)
    ang += span

# ---------- 5. 标题 / 副标题 / 脚注 ----------
fig.text(0.06185, 1 - 66.3 / 796, "2022年上半年各年龄段人数分布",
         fontsize=FS_TITLE, fontweight="bold", color=C_TEXT,
         ha="left", va="baseline")
fig.text(0.06240, 1 - 121.5 / 796, "公司平均年龄32.5，20-30员工比例最高占比37.5%",
         fontsize=FS_SUB, color=C_TEXT, ha="left", va="baseline")
fig.text(0.06160, 1 - 760 / 796, "*注：数据来源于公司人力资源系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_TEXT, ha="left", va="baseline")

# ---------- 6. 输出（保存到指定目录） ----------
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(os.path.join(OUT_DIR, "21.png"), dpi=DPI, facecolor=BG)
plt.show()
