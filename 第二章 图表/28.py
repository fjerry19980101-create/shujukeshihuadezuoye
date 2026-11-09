# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 08:49:18 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\28 复合柱形图.xlsx"
raw = pd.read_excel(XLSX, sheet_name="28 复合柱形图", header=None)
hdr = next(i for i, r in raw.iterrows() if any(str(v).strip() == "月份" for v in r.values))
df = raw.iloc[hdr + 1:, 1:4].copy()
df.columns = ["月份", "月度销量", "季度销量"]
df = df.dropna(subset=["月份"]).reset_index(drop=True)
df["月度销量"] = pd.to_numeric(df["月度销量"], errors="coerce")
df["季度销量"] = pd.to_numeric(df["季度销量"], errors="coerce")

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"                                         # 画布底色
Q_COLORS = ["#4A5BD1", "#FC6C0F", "#F6A50B", "#A2417B"]  # 一~四季度柱色
AXIS_C = "#454866"                                     # 底部轴线
C_TITLE = C_SUB = "#FFFFFF"
C_MVAL = C_QVAL = C_XLAB = "#F2F2F2"                   # 柱值/季度值/月标签
C_NOTE = "#D9D9D9"                                     # 脚注

FS_TITLE, FS_SUB = 34.5, 21.2
FS_QVAL, FS_MVAL, FS_XLAB, FS_NOTE = 14.4, 13.0, 15.5, 15.4

K = 24.2          # 比例尺：24.2 销量/像素
BASE = 627.0      # 数据基线像素行
AXIS_Y = 627.5    # 轴线中心
GAP_Q = 22.0      # 季度标签基线距色块顶
GAP_M = 19.8      # 柱值标签基线距柱顶
BX = [107, 341, 574, 807]   # 四个季度色块左边界
BW, BARW = 175.0, 23.0      # 色块宽 / 柱宽


def yp(v):        # 销量 -> 像素行
    return BASE - v / K


# ---------- 3. 像素坐标系画布 ----------
fig = plt.figure(figsize=(11.05, 7.56), dpi=100, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, 1105)
ax.set_ylim(756, 0)          # y 向下，与图片行号一致
ax.axis("off")

qsum = df["季度销量"].to_numpy(float).reshape(4, 3)[:, 0]
vals = df["月度销量"].to_numpy(float)

# ---------- 4. 季度背景块（同柱色 alpha 0.2） ----------
for q in range(4):
    ax.add_patch(Rectangle((BX[q], yp(qsum[q])), BW, BASE - yp(qsum[q]),
                           facecolor=Q_COLORS[q], alpha=0.20, edgecolor="none", zorder=1))

# ---------- 5. 柱子 ----------
centers = []
for m in range(len(vals)):
    q, s = m // 3, m % 3                       # 季度 / 组内序号
    cx = BX[q] + BW / 3 * (s + 0.5)
    centers.append(cx)
    ax.add_patch(Rectangle((cx - BARW / 2, yp(vals[m])), BARW, BASE - yp(vals[m]),
                           facecolor=Q_COLORS[q], edgecolor="none", zorder=2))

# ---------- 6. 底部轴线 ----------
ax.plot([107, 982], [AXIS_Y, AXIS_Y], color=AXIS_C, lw=1.44,
        solid_capstyle="butt", zorder=3)

# ---------- 7. 数值标签 ----------
for m in range(len(vals)):                     # 柱顶月销量
    ax.text(centers[m], yp(vals[m]) - GAP_M, "%d" % vals[m], ha="center",
            va="baseline", fontsize=FS_MVAL, color=C_MVAL, zorder=4)
for q in range(4):                             # 色块顶季度合计
    ax.text(BX[q] + BW / 2, yp(qsum[q]) - GAP_Q, "%d" % qsum[q], ha="center",
            va="baseline", fontsize=FS_QVAL, color=C_QVAL, zorder=4)

# ---------- 8. 月份标签 ----------
for m in range(len(vals)):
    ax.text(centers[m], 665, df["月份"].iloc[m], ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_XLAB, zorder=4)

# ---------- 9. 标题 / 副标题 / 脚注 ----------
fig.text(0.06733, 1 - 79 / 756, "2021年各月化妆品销量走势",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(0.06733, 1 - 124.5 / 756, "2021年第三季度销量最多9673，9月单月销量最大3621",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(0.05643, 1 - 724 / 756, "*注：数据来源于公司销售系统，统计日期截至2021.12.31",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

fig.savefig("28_复合柱形图_repro.png", dpi=100, facecolor=BG)
plt.show()
