# -*- coding: utf-8 -*-
"""
第四章 人力资源可视化看板 —— Python 复现（Matplotlib）

数据源：'第四章 人力资源可视化看板.xlsx' -> 工作表 202203人员基础信息（A:H 为员工明细）
输出  ：同目录下 第四章 人力资源可视化看板_复现.png / .pdf

版式说明
--------
原看板由一块 11 列 × 25 行的深蓝单元格区域（#2C3B79）承载，
其上浮着 6 个深藏青图表卡片（#1D244A）与 1 个标题文本框。
本脚本按原看板的比例（画布 1606 × 816 设计像素）逐像素还原，
所有坐标均取自原文件的图形锚点与导出 PDF 的实测位置。
字体统一 微软雅黑；标题 14pt、图内标题 10pt、标签 7pt、总人数 11pt。

与原件的有意差异（两处，均为原文件自身的缺陷）
----------------------------------------------
1. 「总人数：1470」在原文件中是挂在图表内的浮动文本框，文字被文本框边界裁掉，
   这里保留同一位置与字号，但让其完整显示。
2. 原文件学历汇总表的公式引用错行（专科=170、专科以下=282 写反了），
   本脚本按员工明细重新统计，得到 专科=282、专科以下=170。
"""

import os
import numpy as np
import pandas as pd
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, Wedge
from matplotlib import font_manager

# --------------------------------------------------------------------------
# 0. 全局参数
# --------------------------------------------------------------------------
def pick_font():
    """优先用原看板字体 微软雅黑，缺失时回退，避免中文变方块"""
    installed = {f.name for f in font_manager.fontManager.ttflist}
    for name in ('Microsoft YaHei', 'Microsoft YaHei UI', 'SimHei', 'DengXian'):
        if name in installed:
            return name
    return 'sans-serif'


mpl.rcParams['font.family'] = pick_font()
mpl.rcParams['axes.unicode_minus'] = False

SRC = r'C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\作业' \
      r'\《Excel数据可视化——从图表到数据大屏》\第四章\第四章 人力资源可视化看板.xlsx'
OUT_DIR = os.path.dirname(SRC)
OUT_PNG = os.path.join(OUT_DIR, '第四章 人力资源可视化看板_复现.png')
OUT_PDF = os.path.join(OUT_DIR, '第四章 人力资源可视化看板_复现.pdf')

# 配色（取自原图表 XML）
C_BG      = '#2C3B79'   # 看板底色
C_CARD    = '#1D244A'   # 图表卡片底色
C_TEXT    = '#F2F2F2'   # 文字（bg1 lumMod 95%）
C_BAR     = '#7AD5FE'   # 条形/柱形主色
C_BLUE_L  = '#00B0F0'
C_ORANGE  = '#ED7D31'
C_SKY     = '#7AD5FE'
C_AMBER   = '#EDA530'
C_BLUE_D  = '#0070C0'

# 画布：设计像素 -> 英寸（比例取自原看板导出的实测尺寸）
W_PX, H_PX = 1606, 816
PPI = 191.0
FIG_W, FIG_H = W_PX / PPI, H_PX / PPI

FS_TITLE = 14    # 看板主标题
FS_KPI   = 11    # 总人数
FS_HEAD  = 10    # 卡片标题
FS_LABEL = 7     # 图例 / 数据标签 / 分类标签

PT = PPI / 72.0  # 1pt 对应的像素数，用于把像素偏移换算成 point 偏移
DY = 2.0         # 文字基线微调：matplotlib 的 va='center' 比 Excel 高约 2px


# --------------------------------------------------------------------------
# 1. 坐标工具：以看板左上角为原点、y 向下的设计坐标系
# --------------------------------------------------------------------------
def fbox(x0, y0, w, h):
    """设计坐标下的矩形 -> matplotlib figure 坐标 [left, bottom, width, height]"""
    return [x0 / W_PX, 1 - (y0 + h) / H_PX, w / W_PX, h / H_PX]


def fx(x):
    return x / W_PX


def fy(y):
    return 1 - y / H_PX


def pts(px):
    """像素偏移 -> 磅值偏移（供 annotate 的 offset points 使用）"""
    return px / PT


# --------------------------------------------------------------------------
# 2. 读取并统计数据
# --------------------------------------------------------------------------
def load_data(path):
    raw = pd.read_excel(path, sheet_name='202203人员基础信息', usecols='A:H')
    raw = raw.dropna(subset=['员工编号'])
    return {
        'total': len(raw),
        'age':   raw['年龄段'].value_counts(),
        'gender': raw['性别'].value_counts(),
        'edu':   raw['学历'].value_counts(),
        'marry': raw['婚姻状况'].value_counts(),
        'dept':  raw['部门'].value_counts(),
        'move':  raw['本月入转调'].value_counts(),
    }


def pct_labels(values):
    """最大余数法取整百分比：保证标签合计为 100%，与原看板显示口径一致"""
    total = float(sum(values))
    exact = [v / total * 100 for v in values]
    base = [int(np.floor(e)) for e in exact]
    rest = 100 - sum(base)
    order = sorted(range(len(values)), key=lambda i: exact[i] - base[i], reverse=True)
    for i in order[:rest]:
        base[i] += 1
    return base


# --------------------------------------------------------------------------
# 3. 组件绘制
# --------------------------------------------------------------------------
def draw_card(fig, x, y, w, h):
    """深藏青卡片底：无边框、无圆角"""
    fig.add_artist(Rectangle((fx(x), fy(y + h)), w / W_PX, h / H_PX,
                             transform=fig.transFigure, facecolor=C_CARD,
                             edgecolor='none', zorder=1))


def draw_legend(fig, x, y0, step, items):
    """右侧竖排图例：11px 方块 + 7pt 文字"""
    sq = 11
    for i, (label, color) in enumerate(items):
        cy = y0 + i * step
        fig.add_artist(Rectangle((fx(x), fy(cy + sq / 2)), sq / W_PX, sq / H_PX,
                                 transform=fig.transFigure, facecolor=color,
                                 edgecolor='none', zorder=3))
        fig.text(fx(x + sq + 7), fy(cy + DY), label, ha='left', va='center',
                 fontsize=FS_LABEL, color=C_TEXT, zorder=3)


def draw_doughnut(fig, box, cx, cy, r, values, labels, colors, legend_x, legend_y0,
                  legend_step, title):
    """甜甜圈卡片：卡片 + 标题 + 环形图 + 外部百分比标签 + 右侧图例"""
    draw_card(fig, *box)
    fig.text(fx(box[0] + box[2] / 2), fy(box[1] + 45.5 + DY), title,
             ha='center', va='center',
             fontsize=FS_HEAD, fontweight='bold', color=C_TEXT, zorder=3)

    # 环形图：单独建一个物理上为正方形的坐标系，保证圆不会变成椭圆
    ax = fig.add_axes(fbox(cx - r, cy - r, 2 * r, 2 * r), zorder=2)
    ax.set_xlim(-1, 1)
    ax.set_ylim(-1, 1)
    ax.axis('off')

    total = float(sum(values))
    pcts = pct_labels(values)
    start = 0.0                                   # 12 点方向起，顺时针排列
    for v, c in zip(values, colors):
        sweep = v / total * 360.0
        # matplotlib 的 Wedge 逆时针绘制，故起止角度做镜像
        ax.add_patch(Wedge((0, 0), 1.0, 90 - (start + sweep), 90 - start,
                           width=0.25, facecolor=c, edgecolor='none'))
        start += sweep

    # 百分比标签：放在扇形角平分线上、环外 1.52R 处（与原图 outsideEnd 效果一致）
    start = 0.0
    k = 1.52
    for v, p in zip(values, pcts):
        sweep = v / total * 360.0
        theta = np.deg2rad(start + sweep / 2.0)
        ax.text(k * np.sin(theta), k * np.cos(theta), f'{p}%',
                ha='center', va='center', fontsize=FS_LABEL, color=C_TEXT)
        start += sweep

    draw_legend(fig, legend_x, legend_y0, legend_step,
                list(zip(labels, colors)))


def draw_hbar_chart(fig, box, title, cats, values, bar_x, first_cy, slot,
                    bar_h, px_per_unit, ax_w, title_cx=None):
    """横向条形图（学历 / 部门）"""
    draw_card(fig, *box)
    x, y, w, h = box
    cx_title = title_cx if title_cx is not None else x + w / 2
    fig.text(fx(cx_title), fy(y + 45.5 + DY), title, ha='center', va='center',
             fontsize=FS_HEAD, fontweight='bold', color=C_TEXT, zorder=3)

    n = len(cats)
    top = first_cy - slot / 2
    ax = fig.add_axes(fbox(bar_x, top, ax_w, slot * n), zorder=2)
    ax.set_xlim(0, ax_w / px_per_unit)
    ax.set_ylim(n - 0.5, -0.5)
    ax.axis('off')

    ax.barh(range(n), values, height=bar_h / slot, color=C_BAR,
            edgecolor='none', zorder=2)

    for i, (c, v) in enumerate(zip(cats, values)):
        ax.annotate(c, xy=(0, i), xytext=(-pts(22), 0), textcoords='offset points',
                    ha='right', va='center', fontsize=FS_LABEL, color=C_TEXT)
        ax.annotate(f'{v}', xy=(v, i), xytext=(pts(20), 0), textcoords='offset points',
                    ha='left', va='center', fontsize=FS_LABEL, color=C_TEXT)


def draw_column_chart(fig, box, title, cats, values, centers, baseline, plot_top,
                      bar_w, ymax, title_cx=None):
    """纵向柱形图（入转调）"""
    draw_card(fig, *box)
    x, y, w, h = box
    cx_title = title_cx if title_cx is not None else centers.mean()
    fig.text(fx(cx_title), fy(y + 116 + DY), title, ha='center', va='center',
             fontsize=FS_HEAD, fontweight='bold', color=C_TEXT, zorder=3)

    slot = centers[1] - centers[0]
    ax = fig.add_axes(fbox(centers[0] - slot / 2, plot_top,
                           slot * len(centers), baseline - plot_top), zorder=2)
    ax.set_xlim(-0.5, len(centers) - 0.5)
    ax.set_ylim(0, ymax)
    ax.axis('off')

    ax.bar(range(len(values)), values, width=bar_w / slot, color=C_BAR,
           edgecolor='none', zorder=2)

    for i, (c, v) in enumerate(zip(cats, values)):
        ax.annotate(f'{v}', xy=(i, v), xytext=(0, pts(23)), textcoords='offset points',
                    ha='center', va='bottom', fontsize=FS_LABEL, color=C_TEXT)
        ax.annotate(c, xy=(i, 0), xytext=(0, -pts(23)), textcoords='offset points',
                    ha='center', va='top', fontsize=FS_LABEL, color=C_TEXT)


# --------------------------------------------------------------------------
# 4. 主流程
# --------------------------------------------------------------------------
def build_dashboard(d):
    fig = plt.figure(figsize=(FIG_W, FIG_H), dpi=200, facecolor=C_BG)

    # ---- 看板主标题 ----
    fig.text(fx(806), fy(44 + DY), '公司人员结构看板', ha='center', va='center',
             fontsize=FS_TITLE, fontweight='bold', color=C_TEXT, zorder=5)

    # ---- 卡片几何（设计像素：x, y, w, h；原点为看板左上角，y 向下）----
    # 上排 3 张甜甜圈卡片等高 276px，下排 3 张柱/条形卡片等高 391px，
    # 卡片之间留出 37px 的背景缝，呈 3 × 2 网格。
    C_AGE    = (33,   86,  486, 276)
    C_GENDER = (556,  86,  488, 276)
    C_MARRY  = (1083, 86,  486, 276)
    C_EDU    = (33,   391, 486, 391)
    C_MOVE   = (556,  391, 488, 391)
    C_DEPT   = (1083, 391, 486, 391)

    age_order    = ['18-24', '25-29', '30-34', '35-39', '40=<']
    age_colors   = [C_BLUE_L, C_ORANGE, C_SKY, C_AMBER, C_BLUE_D]
    age_values   = [int(d['age'].get(k, 0)) for k in age_order]

    draw_doughnut(fig, C_AGE, cx=220, cy=260, r=58,
                  values=age_values, labels=age_order, colors=age_colors,
                  legend_x=410, legend_y0=172.5, legend_step=39, title='年龄')

    gender_order  = ['男', '女']
    gender_colors = [C_AMBER, C_BLUE_D]
    gender_values = [int(d['gender'].get(k, 0)) for k in gender_order]
    draw_doughnut(fig, C_GENDER, cx=763, cy=260, r=58,
                  values=gender_values, labels=gender_order, colors=gender_colors,
                  legend_x=982, legend_y0=239, legend_step=42, title='性别')

    marry_order  = ['单身', '已婚', '离异']
    marry_colors = [C_AMBER, C_BLUE_D, C_SKY]
    marry_values = [int(d['marry'].get(k, 0)) for k in marry_order]
    draw_doughnut(fig, C_MARRY, cx=1279, cy=260, r=58,
                  values=marry_values, labels=marry_order, colors=marry_colors,
                  legend_x=1487, legend_y0=218, legend_step=42, title='婚姻状况')

    # 学历：条形自上而下，从「博士研究生」到「专科以下」
    # 注：原表 K24/K25 的 COUNTIFS 互相引用了对方的标签，导致「专科」「专科以下」数值互换；
    #     这里按员工明细重新统计，所以是 专科 282 / 专科以下 170。
    edu_order  = ['博士研究生', '硕士研究生', '本科', '专科', '专科以下']
    edu_values = [int(d['edu'].get(k, 0)) for k in edu_order]
    draw_hbar_chart(fig, C_EDU, '学历', edu_order, edu_values,
                    bar_x=171, first_cy=519.5, slot=51.75, bar_h=18,
                    px_per_unit=0.4003, ax_w=330, title_cx=279)

    # 部门
    dept_order  = ['财务部', '人力资源部', '行政部', '信息技术部', '研发部', '销售部']
    dept_values = [int(d['dept'].get(k, 0)) for k in dept_order]
    draw_hbar_chart(fig, C_DEPT, '部门', dept_order, dept_values,
                    bar_x=1218, first_cy=515, slot=43.1, bar_h=14.5,
                    px_per_unit=0.3204, ax_w=320, title_cx=1325)

    # 入转调
    move_order  = ['入职', '转入', '转出']
    move_values = [int(d['move'].get(k, 0)) for k in move_order]
    draw_column_chart(fig, C_MOVE, '入转调', move_order, move_values,
                      centers=np.array([657.5, 799.5, 940.5]),
                      baseline=722, plot_top=545, bar_w=44, ymax=50,
                      title_cx=808.5)

    # 总人数（原文件为图表内的浮动文本框，此处保留同一位置并完整显示）
    fig.text(fx(718), fy(443.5 + DY), f'总人数：{d["total"]}', ha='left', va='center',
             fontsize=FS_KPI, fontweight='bold', color=C_TEXT, zorder=4)

    return fig


def main():
    d = load_data(SRC)
    fig = build_dashboard(d)
    fig.savefig(OUT_PNG, dpi=300, facecolor=C_BG)
    fig.savefig(OUT_PDF, facecolor=C_BG)
    print(f'已保存：{OUT_PNG}')
    print(f'已保存：{OUT_PDF}')
    print(f'总人数 {d["total"]}｜年龄 {d["age"].to_dict()}｜性别 {d["gender"].to_dict()}')
    plt.show()


if __name__ == '__main__':
    main()
