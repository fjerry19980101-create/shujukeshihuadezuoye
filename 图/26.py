# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 08:59:57 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
"""26 柱形圆 复现（尺寸/坐标/配色/字号均为参考图像素级实测值）"""
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Circle

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\26 柱形圆.xlsx"
raw = pd.read_excel(XLSX, sheet_name="26 柱形圆", header=None)
hdr = next(i for i, r in raw.iterrows() if any(str(v).strip() == "区域" for v in r.values))
df = raw.iloc[hdr + 1:, 1:5].copy()
df.columns = ["区域", "销量", "占位", "同比去年"]
df = df.dropna(subset=["区域"]).reset_index(drop=True)
for c in ["销量", "占位", "同比去年"]:
    df[c] = pd.to_numeric(df[c], errors="coerce")
df = df.dropna().reset_index(drop=True)

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False

BG = "#1A1E43"          # 画布底色
C_BAR = "#E74E69"       # 柱（粉）
C_CIR = "#0070C0"       # 圆形气泡（蓝）
C_AXIS = "#454866"      # 底部轴线
C_TITLE = C_SUB = "#FFFFFF"
C_VAL = C_PCT = C_XLAB = "#F2F2F2"
C_NOTE = "#D9D9D9"

FS_TITLE, FS_SUB = 34.6, 21.1
FS_VAL, FS_PCT, FS_XLAB, FS_NOTE = 14.3, 14.3, 16.3, 15.4

W, H, DPI = 1088, 759, 100          # 画布 1088 × 759 px

X0, PITCH = 163.5, 149.5            # 第 1 根柱中心 / 柱间距
BARW = 46.5                         # 柱宽
BASE, K = 642.0, 11.05              # 数值轴基线像素行 / 每像素单位数
GAP_VAL = 19.8                      # 数值标签基线偏移
D_A, D_B = 2.686, 1.2               # 圆形直径 = D_A * 同比(%) - D_B  (px)
AXIS_Y, AXIS_X0, AXIS_X1 = 642.5, 89, 988


def yp(v):
    """销量 -> 像素行"""
    return BASE - v / K


# ---------- 3. 像素坐标系画布 ----------
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)                   # y 向下，与图片行号一致
ax.axis("off")

n = len(df)
centers = [X0 + PITCH * i for i in range(n)]

# ---------- 4. 底部轴线 ----------
ax.plot([AXIS_X0, AXIS_X1], [AXIS_Y, AXIS_Y], color=C_AXIS, lw=1.44,
        solid_capstyle="butt", zorder=1)

# ---------- 5. 柱形 ----------
for i, r in df.iterrows():
    ax.add_patch(Rectangle((centers[i] - BARW / 2, yp(r["销量"])), BARW,
                           BASE - yp(r["销量"]), facecolor=C_BAR,
                           edgecolor="none", zorder=2))

# ---------- 6. 顶部气泡：直径与同比正相关，圆心落在「占位」值的高度上 ----------
y_cir = yp(df["占位"].iloc[0])      # 占位列(4500) -> 圆心所在像素行
for i, r in df.iterrows():
    pct = r["同比去年"] * 100
    dia = D_A * pct - D_B
    ax.add_patch(Circle((centers[i], y_cir), dia / 2, facecolor=C_CIR,
                        edgecolor="none", zorder=3))
    ax.text(centers[i], y_cir + 2.2, "%.0f%%" % pct, ha="center", va="center",
            fontsize=FS_PCT, color=C_PCT, zorder=4)

# ---------- 7. 柱顶数值标签 ----------
for i, r in df.iterrows():
    ax.text(centers[i], yp(r["销量"]) - GAP_VAL, "%d" % r["销量"], ha="center",
            va="baseline", fontsize=FS_VAL, color=C_VAL, zorder=4)

# ---------- 8. 区域标签 ----------
for i, r in df.iterrows():
    ax.text(centers[i], 679, str(r["区域"]), ha="center", va="baseline",
            fontsize=FS_XLAB, color=C_XLAB, zorder=4)

# ---------- 9. 标题 / 副标题 / 脚注 ----------
fig.text(0.06011, 1 - 75 / 759, "各区域上半年销量以同比",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(0.06195, 1 - 118.5 / 759, "东北区域销量持续保持第一，华东和华南同比去年增长最多",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(0.04260, 1 - 718 / 759, "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

# ---------- 10. 输出 ----------
fig.savefig("26.png", dpi=DPI, facecolor=BG)
plt.show()
