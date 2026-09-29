# -*- coding: utf-8 -*-
"""
Created on Tue Sep 22 08:30:03 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt

# ---------- 1. 读数据 ----------
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(后15)\29 滑珠图.xlsx"
raw = pd.read_excel(XLSX, sheet_name="29 滑珠图", header=None)
hdr = next(i for i, r in raw.iterrows() if any(str(v).strip() == "区域" for v in r.values))
df = raw.iloc[hdr + 1:, 1:4].copy()
df.columns = ["区域", "完成率", "未完成占比"]
df = df.dropna(subset=["区域"])
df["区域"] = df["区域"].astype(str).str.strip()
df["完成率"] = pd.to_numeric(df["完成率"], errors="coerce")
df = (df.dropna(subset=["完成率"]).sort_values("完成率", ascending=False)
        .reset_index(drop=True))                     # 华南在最上

# ---------- 2. 样式常量（图片取样值） ----------
mpl.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Arial Unicode MS"]
mpl.rcParams["axes.unicode_minus"] = False
BG, BLUE, TRACK = "#1A1E43", "#0070C0", "#717389"
C_TITLE = C_SUB = C_PCT = "#FFFFFF"                  # 单值连续赋值，没问题
C_LABEL, C_NOTE = "#B3B4C0", "#D9D9D9"
FS_TITLE, FS_SUB = 34.6, 21.2
FS_YLAB, FS_PCT, FS_NOTE = 16.3, 16.0, 15.4

X_TITLE, X_SUB, X_NOTE = 0.07452, 0.07452, 0.07543   # ★ 订正：三个变量分别赋值
BL_TITLE, BL_SUB, BL_NOTE = 74, 118, 715             # 基线像素行

BAR_LW, DOT_MS, DOT_MEW = 26.6, 29.2, 1.8            # 条高 37px、圆点外径 43px
CAP_R = 0.0299                                       # 圆头端帽半径(18.5px)→数据单位

# ---------- 3. 画布 / 坐标系 ----------
fig = plt.figure(figsize=(11.05, 7.56), dpi=100, facecolor=BG)
ax = fig.add_axes([180/1105, (756-656.4)/756, 619/1105, 463.8/756])
ax.set_facecolor(BG)
n = len(df)
ax.set_xlim(0, 1)
ax.set_ylim(-0.5, n - 0.5)
ax.set_xticks([])
for s in ax.spines.values():
    s.set_visible(False)

# ---------- 4. 逐行绘制 ----------
for i, row in df.iterrows():
    y = n - 1 - i
    v = float(row["完成率"])
    ax.plot([CAP_R, 1 - CAP_R], [y, y], color=TRACK, lw=BAR_LW,
            solid_capstyle="round", zorder=2)                    # 滑道
    ax.plot([CAP_R, v], [y, y], color=BLUE, lw=BAR_LW,
            solid_capstyle="round", zorder=3)                     # 进度条
    ax.plot([v], [y], marker="o", ms=DOT_MS, mfc=BLUE, mec="white",
            mew=DOT_MEW, ls="none", zorder=4)                     # 圆点
    ax.text(v - 0.0569, y - 0.032, f"{v:.0%}", ha="right", va="center",
            fontsize=FS_PCT, color=C_PCT, zorder=5)               # 条内数值

ax.set_yticks(range(n))
ax.set_yticklabels(list(df["区域"])[::-1], fontsize=FS_YLAB, color=C_LABEL)
ax.tick_params(axis="y", length=0, pad=14.2)
for t in ax.get_yticklabels():
    t.set_verticalalignment("center")

# ---------- 5. 标题 / 副标题 / 脚注 ----------
fig.text(X_TITLE, 1 - BL_TITLE/756, "2022年上半年产品销量目标达成率情况",
         fontsize=FS_TITLE, fontweight="bold", color=C_TITLE, ha="left", va="baseline")
fig.text(X_SUB, 1 - BL_SUB/756, "华南完成率最高达到86%，华东最低35%",
         fontsize=FS_SUB, color=C_SUB, ha="left", va="baseline")
fig.text(X_NOTE, 1 - BL_NOTE/756, "*注：数据来源于公司销售系统，统计日期截至2022.06.30",
         fontsize=FS_NOTE, color=C_NOTE, ha="left", va="baseline")

fig.savefig("29_滑珠图_repro.png", dpi=100, facecolor=BG)
plt.show()
