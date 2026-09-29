# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 08:12:53 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""24 目标柱形图 复现（尺寸/坐标/配色/字号均为参考图像素级实测值）"""
import os
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\24 目标柱形图.xlsx"
RAW = pd.read_excel(XLSX, sheet_name="24 目标柱形图", header=None)
hdr = next(i for i, r in RAW.iterrows() if any(str(v).strip() == "商品" for v in r.values))
df = RAW.iloc[hdr + 1:, 1:4].copy()
df.columns = ["商品", "实际销量", "目标销量"]
df = df.dropna(subset=["商品"]).reset_index(drop=True)
df["实际销量"] = pd.to_numeric(df["实际销量"], errors="coerce")
df["目标销量"] = pd.to_numeric(df["目标销量"], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_ACT = "#016DBC"       # 实际销量（实心柱）
C_TGT = "#86889A"       # 目标销量（空心外框）
C_AXIS = "#152E5C"      # 底部轴线
C_TITLE = C_SUB = "#FFFFFF"
C_VAL = "#F2F2F2"       # 柱顶数值
C_XLAB = C_LEG = "#B2B2BC"   # 商品名 / 图例文字
C_NOTE = "#D9D9D9"      # 脚注

FS_TITLE, FS_SUB = 34.6, 21.1
FS_VAL, FS_XLAB, FS_LEG, FS_NOTE = 14.9, 16.3, 15.0, 15.4

W, H, DPI = 1105, 756, 100            # 画布 1105 × 756 px

X0, PITCH = 161.0, 156.0              # 第 1 组中心 / 组间距
BW_ACT, BW_TGT = 28.0, 38.0           # 实心柱宽 / 目标外框路径宽（外沿 41）
BASE, K = 610.0, 2.444                # 数值轴基线像素行 / 每像素单位数
GAP_BAR, GAP_TGT = 22.5, 17.0         # 数值标签距「实际柱顶 / 目标框顶」的净空
AXIS_X0, AXIS_X1, AXIS_Y = 83, 1019, 610.5
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
OUT = os.path.join(OUT_DIR, "24.png")


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

# ---------- 5. 逐商品绘制：先目标空心框，再实际实心柱 ----------
for i, r in df.iterrows():
    cx = centers[i]
    y_tgt, y_act = yp(r["目标销量"]), yp(r["实际销量"])
    # 5.1 目标销量：仅描边的空心柱
    ax.add_patch(Rectangle((cx - BW_TGT / 2, y_tgt), BW_TGT, BASE - y_tgt,
                           facecolor="none", edgecolor=C_TGT, lw=1.44, zorder=2))
    # 5.2 实际销量：实心柱
    ax.add_patch(Rectangle((cx - BW_ACT / 2, y_act), BW_ACT, BASE - y_act,
                           facecolor=C_ACT, edgecolor="none", zorder=3))
    # 5.3 柱顶数值：标签底边避让更高的那一个（实际柱 / 目标框）
    lab_bottom = min(y_act - GAP_BAR, y_tgt - GAP_TGT)
    ax.text(cx, lab_bottom + 1, "%d" % r["实际销量"], ha="center", va="baseline",
            fontsize=FS_VAL, color=C_VAL, zorder=4)

# ---------- 6. 商品名 ----------
for i, r in df.iterrows():
    ax.text(centers[i], 646.5, str(r["商品"]), ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_XLAB, zorder=4)

# ---------- 7. 图例（左上：空心方块 = 目标，实心方块 = 实际） ----------
ax.add_patch(Rectangle((115, 200), 14, 14, facecolor="none",
                       edgecolor=C_TGT, lw=1.44, zorder=4))
ax.text(138.5, 212, "目标销量", ha="left", va="baseline", fontsize=FS_LEG,
        color=C_LEG, zorder=4)
ax.add_patch(Rectangle((237, 199), 13, 14, facecolor=C_ACT,
                       edgecolor="none", zorder=4))
ax.text(257.5, 212, "实际销量", ha="left", va="baseline", fontsize=FS_LEG,
        color=C_LEG, zorder=4)

# ---------- 8. 标题 / 副标题 / 脚注 ----------
fig.text(0.07276, 1 - 76 / 756, "2022年上半年各商品销量完成情况",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(0.07367, 1 - 121.5 / 756, "防晒整体销量最好，达到856，面霜远超目标，超额完成30%",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(0.07090, 1 - 715 / 756, "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

# ---------- 9. 输出（保存到指定目录） ----------
os.makedirs(OUT_DIR, exist_ok=True)
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.show()
