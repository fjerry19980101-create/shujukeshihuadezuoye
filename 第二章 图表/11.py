# -*- coding: utf-8 -*-
"""
Created on Thu Sep 24 10:43:27 2026

@author: fjwsm
"""

# -*- coding: utf-8 -*-
import os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
from matplotlib.patches import FancyBboxPatch
import numpy as np
import pandas as pd
from scipy.interpolate import CubicSpline

# ============================================================
# 0. 数据读取
# ============================================================
XLSX = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\第二章 图表(前15)\11 平滑折线图.xlsx"

# 兜底数据（与 xlsx 一致）
MONTHS = ["5月", "6月", "7月", "8月", "9月", "10月", "11月", "12月", "1月", "2月", "3月"]
SALES = [146, 198, 296, 412, 506, 615, 789, 1021, 3782, 3215, 2936]
TITLE = "化妆品品类月度销量走势"
SUBTITLE = "2022年销量迅速增加，1月最高，销量达到3782"
FOOTNOTE = "*注：数据来源于公司销售系统，统计日期截至2022.03.31"

def load_data(path):
    """从 xlsx 读取月份与销量两列（表头行含 '月份' '销量'）。"""
    try:
        raw = pd.read_excel(path, sheet_name=0, header=None)
        hdr, mcol, vcol = None, None, None
        for i, row in raw.iterrows():
            cells = [(j, str(v).strip()) for j, v in enumerate(row.values) if pd.notna(v)]
            tags = [c for _, c in cells]
            if any("月份" in c for c in tags) and any("销量" in c for c in tags):
                hdr = int(i)
                mcol = next(j for j, c in cells if "月份" in c)
                vcol = next(j for j, c in cells if "销量" in c)
                break
        if hdr is None:
            raise ValueError("未找到表头（月份/销量）")
        df = raw.iloc[hdr + 1:, [mcol, vcol]].dropna()
        ms = [str(m).strip() for m in df[mcol].tolist()]
        vs = [int(float(v)) for v in df[vcol].tolist()]
        return ms, vs
    except Exception as e:
        print("[warn] 读取 xlsx 失败，使用内置数据:", e)
    return MONTHS, SALES

MONTHS, SALES = load_data(XLSX)

# ============================================================
# 1. 画布与颜色
# ============================================================
W, H = 1088, 792
DPI = 100
BG = "#1A1E43"          # 背景
ORANGE = "#E66B4C"      # 曲线 / 虚线 / Y轴刻度
CYAN = "#66CBDD"        # 2021 年带
YELLOW = "#F5C353"      # 2022 年带
WHITE = "#FFFFFF"       # 标题 / 副标题 / 年带文字 / 轴线
SOFT = "#F2F2F2"        # 月份标签 / 峰值标签
GREY = "#D9D9D9"        # 脚注

# 字体
FP_REG = fm.FontProperties(fname="C:/Windows/Fonts/msyh.ttc")
FP_BOLD = fm.FontProperties(fname="C:/Windows/Fonts/msyhbd.ttc")

# 字号（pt，宽度对标实测标定）
FS_TITLE = 38.4         # 标题（粗体）
FS_SUB = 28.0           # 副标题
FS_YLABEL = 13.2        # Y 轴刻度（橙）
FS_MONTH = 15.5         # X 轴月份
FS_PEAK = 15.3          # 峰值标注 3782
FS_YEAR = 17.3          # 年带文字 2021 / 2022
FS_NOTE = 15.3          # 脚注

# ============================================================
# 2. 坐标系：像素坐标，y 向下
# ============================================================
fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
ax = fig.add_axes([0, 0, 1, 1])
ax.set_facecolor(BG)
ax.set_xlim(0, W)
ax.set_ylim(H, 0)
ax.axis("off")

# ============================================================
# 3. 映射与 X 位置（全部实测）
# ============================================================
# 数值 -> y 像素：由 0/1,000/2,000/3,000/4,000 刻度标签最小二乘标定
A0, B0 = 604.32, -0.09296

def yp(v):
    return A0 + B0 * v

# 11 个月份数据点的 x 中心（= X 轴刻度 + 40）
X_CENTER = np.array([167, 250, 334, 416.5, 500.5, 584, 667.5, 751, 834, 918, 1001])

# ============================================================
# 4. 平滑曲线
# ============================================================
# 锚点 = 数据节点（映射） + 峰值区实测形状锚点（见文件头说明 2）
anchors = [(168.0, 589.5), (176.0, 589.5)]                      # 起点平坦段
anchors += [(x, yp(v)) for x, v in zip(X_CENTER[1:7], SALES[1:7])]   # 6月~11月
anchors += [                                                     # 12月起（实测形状）
    (746.0, 514.0), (751.0, 511.0),
    (756.0, 502.5), (762.0, 491.0), (768.0, 477.0), (774.0, 461.0),
    (780.0, 441.5), (786.0, 417.5), (792.0, 388.0), (798.0, 357.0),
    (804.0, 327.5), (810.0, 304.5), (816.0, 286.5), (822.0, 273.0),
    (828.0, 262.5), (834.0, 254.0), (841.0, 249.0),              # 841 为过冲顶点
    (854.0, 254.5), (866.0, 264.0), (878.0, 274.0), (890.0, 284.0),
    (902.0, 293.5), (914.0, 301.5), (926.0, 307.0), (938.0, 311.5),
    (950.0, 315.5), (962.0, 318.5), (974.0, 322.0), (986.0, 325.5),
    (998.0, 328.5), (1003.0, 330.5),                             # 曲线终点
]
ax_arr = np.array([a[0] for a in anchors])
ay_arr = np.array([a[1] for a in anchors])
curve = CubicSpline(ax_arr, ay_arr, bc_type="natural")

xd = np.linspace(168.0, 1003.0, 900)
ax.plot(xd, curve(xd), color=ORANGE, lw=1.44,                 # 2px
        solid_capstyle="round", solid_joinstyle="round", zorder=3)

# ============================================================
# 5. 峰值虚线与 3782 标注
# ============================================================
# 竖直虚线：x=848，y 246.5..599.5；24px 通 / 9px 断
# （matplotlib 虚线长度单位 pt 且按线宽缩放：px*0.72/lw_pt）
ax.plot([848, 848], [246.5, 599.5], color=ORANGE, lw=1.8,      # 2.5px
        linestyle=(0, (9.4, 3.8)), solid_capstyle="butt", zorder=2)

# 峰值标注（白字，无千分位）：位于数据点 1月 上方
ax.text(836.0, 224.0, str(max(SALES)), fontsize=FS_PEAK, color=SOFT,
        ha="center", va="baseline", fontproperties=FP_REG, zorder=5)

# ============================================================
# 6. X 轴：轴线 + 刻度 + 月份标签
# ============================================================
ax.plot([127, 1046], [603.5, 603.5], color=WHITE, lw=1.44,     # 2px
        solid_capstyle="butt", zorder=2)
TICKS = [127, 210, 294, 377, 460, 544, 627, 711, 794, 878, 961, 1045]
for tx in TICKS:
    ax.plot([tx, tx], [605, 613], color=WHITE, lw=1.44,
            solid_capstyle="butt", zorder=2)

for x, m in zip(X_CENTER, MONTHS):
    ax.text(x, 639.0, m, fontsize=FS_MONTH, color=SOFT,
            ha="center", va="baseline", fontproperties=FP_REG)

# ============================================================
# 7. Y 轴刻度标签（橙色，右对齐，含千分位）
# ============================================================
for val, by in [(4000, 240.0), (3000, 333.0), (2000, 426.0), (1000, 519.0), (0, 612.0)]:
    label = "{:,}".format(val) if val else "0"
    ax.text(111.5, by, label, fontsize=FS_YLABEL, color=ORANGE,
            ha="right", va="baseline", fontproperties=FP_REG)

# ============================================================
# 8. 年份色带（圆角矩形）与年带文字
# ============================================================
# 青色 2021：x 101..797，y 662..701；黄色 2022：x 808..1062，y 661..703；圆角 5px
ax.add_patch(FancyBboxPatch((101, 662), 697, 40,
                            boxstyle="round,pad=0,rounding_size=5",
                            facecolor=CYAN, edgecolor="none", zorder=2))
ax.add_patch(FancyBboxPatch((808, 661), 255, 43,
                            boxstyle="round,pad=0,rounding_size=5",
                            facecolor=YELLOW, edgecolor="none", zorder=2))

ax.text(449.5, 690.0, "2021", fontsize=FS_YEAR, color=WHITE,
        ha="center", va="baseline", fontproperties=FP_REG, zorder=3)
ax.text(935.0, 690.0, "2022", fontsize=FS_YEAR, color=WHITE,
        ha="center", va="baseline", fontproperties=FP_REG, zorder=3)

# ============================================================
# 9. 标题 / 副标题 / 脚注
# ============================================================
# 标题 bbox x59..643 y36..86，baseline ≈ 墨迹底行 - 0.13*高
ax.text(59.0, 79.5, TITLE, fontsize=FS_TITLE, color=WHITE,
        ha="left", va="baseline", fontproperties=FP_BOLD)
# 副标题 bbox x60..849 y106..141，CJK 墨迹底行 141 -> baseline ≈ 138
ax.text(60.0, 138.0, SUBTITLE, fontsize=FS_SUB, color=WHITE,
        ha="left", va="baseline", fontproperties=FP_REG)
# 脚注 bbox x51..591 y733..753，baseline ≈ 750
ax.text(51.0, 750.0, FOOTNOTE, fontsize=FS_NOTE, color=GREY,
        ha="left", va="baseline", fontproperties=FP_REG)

# ============================================================
# 10. 输出
# ============================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图"
os.makedirs(OUT_DIR, exist_ok=True)
OUT = os.path.join(OUT_DIR, "11 平滑折线图.png")
fig.savefig(OUT, dpi=DPI, facecolor=BG)
plt.close(fig)
print("saved ->", OUT)
