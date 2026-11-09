# -*- coding: utf-8 -*-
"""
第三章 动态图表 —— 3 动态南丁格尔圆环图（plotly 可交互版）
====================================================================

与 matplotlib 版的区别
    · 版面参数（圆心 / 内径 / 层宽 / 配色 / 字号 / 引线坐标）完全沿用，
      同样量自 944×716 的两张截图；
    · 环带用 go.Barpolar 拼（base=内半径，r=层厚），分隔弧 / 引线 / 圆点 /
      标签用直角坐标系叠加，两条坐标系统一用「像素」作单位；
    · 交互交给 plotly：右上角两个按钮切换 7 天 / 30 天，扇区 hover 出提示。
      补间没有用 plotly 自带的 frames + Plotly.animate——实测它在无头浏览器里
      推不动，而且 newPlot 时还会把第一帧的 layout 盖回图上。改成 Python 端
      算好每一帧（frame_payloads），页面里用 Plotly.restyle / relayout 一帧
      一帧推，节奏和终点都完全可控。

为什么按钮不是 plotly 的 updatemenus
    updatemenus 的按钮是浏览器原生 <button>，宽高由文字和 padding 撑出来，
    做不出截图里「127×55 灰块 + 三环单选圆点」的样子。所以按钮的视觉用
    shapes + annotation 画在图上（导出 PNG 也在），点击热区由一小段
    post_script 里的透明 <div> 承担。

输出（保存在 图/第三章/）
    3 动态南丁格尔圆环图_plotly.html       ← 双击用浏览器打开即可交互
    3 动态南丁格尔圆环图_近7天_plotly.png
    3 动态南丁格尔圆环图_近30天_plotly.png

说明
    · 数据直接读同目录 xlsx 的 B3:G5（两组占比），读不到才用脚本内置值。
    · plotly 由浏览器渲染，字体度量与 Excel / matplotlib 都不同，因此做不到
      逐像素一致；版面元素的位置是对齐的，文字宽度会有 1~3px 差异。
"""

import json
import math
import os
import shutil
import subprocess
import sys
import tempfile
from urllib.parse import quote

import numpy as np
import plotly.graph_objects as go

# ==========================================================================
# 1. 版面参数（单位：像素，坐标系原点在左上角，y 向下）
# ==========================================================================
W, H = 944, 716
CX, CY = 477.0, 393.0            # 圆心
R_IN, R_OUT = 109.0, 210.0       # 内半径 / 外半径
RINGS = 6                        # 同心环带层数
RING_W = (R_OUT - R_IN) / RINGS  # 单层厚度 ≈ 16.83

BG = "#1A1E43"                   # 背景深藏蓝
C_TITLE = "#FFFFFF"
C_LABEL = "#F2F2F2"
C_NOTE = "#D9D9D9"
C_BTN = "#8E8F9A"                # 按钮灰底
C_RADIO = "#646464"              # 单选圆点外环
C_DARK = "#000000"

# plotly 的字号单位就是 px，不用像 matplotlib 那样 pt→px 换算
FS_TITLE, FS_LABEL, FS_NOTE, FS_BTN = 47.8, 24.3, 22.0, 24.0
FONT_TITLE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_TEXT = '"DengXian", "等线", "Microsoft YaHei", sans-serif'

TITLE_TEXT = "2022年6月30日流量来源分布"
NOTE_TEXT = "*注: 数据来源公司网站"
TITLE_POS = (62.0, 85.5)         # 文字左端 + 基线
NOTE_POS = (53.0, 672.0)

# 基线 → annotation 的 bottom 锚点。0.33em 是实测值：这个系数下标题、标签、
# 脚注三者的墨迹位置与截图都落在 ±1px 内。
def _desc(size):
    return 0.33 * size


# annotation 左锚点比 matplotlib 的 text 多出约 2.5px 的字形左侧留白，统一补偿
X_SHIFT = -2.5


# ==========================================================================
# 2. 数据与配色
# ==========================================================================
# 五个渠道：颜色 + 环带层数；个人主页内层黄、外两层绿，引线也是黄色
SEGMENTS = [
    dict(name="首页推荐", color="#4A5BD1", rings=6),
    dict(name="关注页面", color="#6073F3", rings=5),
    dict(name="搜索", color="#11A7AD", rings=4),
    dict(name="个人主页", color="#75BD42", rings=3, line_color="#FFC000",
         ring_colors=["#FFC000", "#75BD42", "#75BD42"]),
    dict(name="其他来源", color="#E66B4C", rings=2),
]

DATA = {
    "近7天": [0.44, 0.17, 0.15, 0.08, 0.16],
    "近30天": [0.38, 0.21, 0.15, 0.09, 0.17],
}

# 两个状态的标签位置与引线折点（量自两张截图）
#   text：文字左端 x、基线 y
#   line：水平段外端 → 拐点 → 端点圆点
LAYOUT = {
    "近7天": {
        "首页推荐": dict(text=(734.5, 256.0), line=[(860, 240.5), (720, 240.5), (681, 287)]),
        "关注页面": dict(text=(524.5, 631.5), line=[(711, 662.5), (576, 662.5), (551, 619)]),
        "搜索": dict(text=(168.5, 523.5), line=[(144, 577.5), (285, 577.5), (336, 535)]),
        "个人主页": dict(text=(158.5, 296.0), line=[(135, 330.5), (277, 330.5), (304, 376)]),
        "其他来源": dict(text=(208.5, 209.0), line=[(201, 221.5), (348, 221.5), (378, 267)]),
    },
    "近30天": {
        "首页推荐": dict(text=(720.5, 220.0), line=[(861, 240.5), (721, 240.5), (682, 287)]),
        "关注页面": dict(text=(576.5, 633.5), line=[(712, 662.5), (577, 662.5), (552, 619)]),
        "搜索": dict(text=(180.5, 546.5), line=[(145, 577.5), (286, 577.5), (337, 535)]),
        "个人主页": dict(text=(154.5, 315.0), line=[(136, 331.5), (278, 331.5), (305, 376)]),
        "其他来源": dict(text=(204.5, 213.0), line=[(202, 222.5), (349, 222.5), (379, 268)]),
    },
}

# 按钮：(x, y, w, h)
BTN_RECTS = {"近7天": (639, 121, 127, 55), "近30天": (783, 120, 127, 55)}

# 环带分隔虚线的节距。量自截图：沿半径 125.83 的圆弧取亮度剖面，
# 参考图的周期约 3.5px、明暗比约 2.5:1，所以取 2.5 + 1.0。
# 这个值配合 width=1.0、darken 系数 0.85，实测弧上最暗 314、对比 60，
# 参考图是 313 / 61，基本重合。
DASH = "2.5px,1px"

# 相邻环带叠画的宽度，用来消掉抗锯齿接缝（见 build_traces 里的说明）
SEAM_OVL = 1.0


def load_data_from_xlsx(path):
    """读取 xlsx 里 B3:G5 的两行占比，按 SEGMENTS 的顺序返回。"""
    try:
        import openpyxl
        ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
        head = [str(ws.cell(3, c).value).strip() for c in range(3, 8)]
        out = {}
        for r in (4, 5):
            key = str(ws.cell(r, 2).value).strip()
            row = {head[i]: float(ws.cell(r, 3 + i).value) for i in range(5)}
            out[key] = [row[s["name"]] for s in SEGMENTS]
        return out if len(out) == 2 else None
    except Exception as exc:
        print(f"[提示] 未能从 xlsx 读取数据（{exc}），使用脚本内置数值。")
        return None


_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "3 动态南丁格尔圆环图.xlsx")
if os.path.exists(_XLSX):
    _loaded = load_data_from_xlsx(_XLSX)
    if _loaded:
        DATA = _loaded


# ==========================================================================
# 3. 几何工具
# ==========================================================================
def darken(hex_color, factor=0.86):
    """压暗颜色，用于环带之间的分隔弧。"""
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return "rgb({},{},{})".format(int(r * factor), int(g * factor), int(b * factor))


def ease(t):
    """缓入缓出，和 matplotlib 版保持一致的节奏。"""
    return t * t * (3.0 - 2.0 * t)


def interp(p):
    """p=0 → 近7天，p=1 → 近30天。"""
    return [a + (b - a) * p for a, b in zip(DATA["近7天"], DATA["近30天"])]


def sector_angles(vals):
    """返回每个渠道的 (起始角, 圆心角)，单位度，从 12 点方向顺时针累计。"""
    out, phi = [], 0.0
    for v in vals:
        sweep = v * 360.0
        out.append((phi, sweep))
        phi += sweep
    return out


def ring_colors(seg):
    """补足到 RINGS 个颜色。"""
    cols = list(seg.get("ring_colors") or [seg["color"]] * seg["rings"])
    return cols + [seg["color"]] * (RINGS - len(cols))


def arc_polar(r, phi0, sweep, step_deg=2.5):
    """沿半径 r 的圆弧取点（极坐标）。2.5° 一段，弧长 ~5.5px，够平滑。

    角度用弧度，半径直接用像素值——极坐标轴的 r 就是 R_OUT=210 的像素刻度。
    """
    n = max(3, int(abs(sweep) / step_deg) + 1)
    th = [round(math.radians(a), 5) for a in np.linspace(phi0, phi0 + sweep, n)]
    return th, [round(r, 2)] * n


# ==========================================================================
# 4. trace 索引约定（逐帧 restyle 按索引更新，务必与 build_traces 的顺序一致）
# ==========================================================================
N_SEG = len(SEGMENTS)
IDX_RING = list(range(0, N_SEG))                 # 5 条环带
IDX_SEP = list(range(N_SEG, 2 * N_SEG))          # 5 组分隔弧
IDX_LINE = list(range(2 * N_SEG, 3 * N_SEG))     # 5 条引线
IDX_DOT = list(range(3 * N_SEG, 4 * N_SEG))      # 5 个端点圆点
IDX_BTN_RING = 4 * N_SEG                         # 按钮外环（静态）
IDX_BTN_WHITE = 4 * N_SEG + 1                    # 按钮白环（静态）
IDX_DOT7 = 4 * N_SEG + 2                         # 选中黑心：近7天
IDX_DOT30 = 4 * N_SEG + 3                        # 选中黑心：近30天
N_TRACE = 4 * N_SEG + 4


def build_traces():
    """建好所有 trace：环带、分隔弧、引线、圆点、按钮部件。"""
    traces = []

    # --- 环带：一个渠道一条，内部用 6 个等角度 bar 叠出 6 层 ---
    # 相邻两层要「叠」1px（SEAM_OVL）：两个 bar 在共同边界上都只盖一半像素时，
    # 背景会从缝里透出来，形成一道比填色更暗的接缝，看起来像条实线，会把
    # 上面那层虚线糊掉。让内层多画 1px，外层盖上去就把缝补实了。
    # 最外层不加，否则环带外缘会鼓出去，破坏 R_OUT=210 的对齐。
    for seg in SEGMENTS:
        cols = ring_colors(seg)
        base = [R_IN + k * RING_W for k in range(RINGS)]
        r = []
        for k in range(RINGS):
            if k >= seg["rings"]:
                r.append(0.0)
            elif k == seg["rings"] - 1:
                r.append(RING_W)
            else:
                r.append(RING_W + SEAM_OVL)
        traces.append(go.Barpolar(
            r=r, base=base,
            theta=[0.0] * RINGS, width=[0.0] * RINGS,
            thetaunit="radians",
            marker=dict(color=cols, line=dict(width=0)),
            name=seg["name"],
            showlegend=False,
            hovertemplate="<b>%{customdata}</b><extra></extra>",
            customdata=[seg["name"]] * RINGS,
        ))

    # --- 分隔弧：层与层之间的虚线弧，用 None 断线串在一个 trace 里 ---
    # 必须用 Scatterpolar 而不是 Scatter。plotly 的 SVG 里 cartesianlayer 排在
    # polarlayer 之前，直角坐标系的 trace 会被极坐标图层整个盖住——曾经用
    # go.Scatter 画这些虚线，改 dash 参数图像一个字节都不变，就是因为它们
    # 根本没显示，看到的「虚线」其实是相邻 bar 抗锯齿接缝的伪影。
    for seg in SEGMENTS:
        cols = ring_colors(seg)
        traces.append(go.Scatterpolar(
            theta=[], r=[], mode="lines", thetaunit="radians",
            line=dict(width=1.0, color=darken(cols[1] if seg["rings"] > 1 else cols[0], 0.85),
                      dash=DASH),
            hoverinfo="skip", showlegend=False,
        ))

    # --- 引线 ---
    for seg in SEGMENTS:
        traces.append(go.Scatter(
            x=[], y=[], mode="lines",
            line=dict(width=3, color=seg.get("line_color", seg["color"])),
            hoverinfo="skip", showlegend=False,
        ))

    # --- 端点圆点 ---
    for seg in SEGMENTS:
        traces.append(go.Scatter(
            x=[], y=[], mode="markers",
            marker=dict(size=10, color=seg.get("line_color", seg["color"])),
            hoverinfo="skip", showlegend=False,
        ))

    # --- 按钮：外环 / 白环 / 两个可选中的黑心 ---
    # size 是直径；实测 plotly 渲染出来比标称小约 2px，所以比截图量到的
    # 22 / 14 / 10 各放大 1~2，才能得到和参考图一样的 21 / 13 / 9。
    rx = [BTN_RECTS[k][0] + 17.5 for k in BTN_RECTS]
    ry = [BTN_RECTS[k][1] + BTN_RECTS[k][3] / 2.0 for k in BTN_RECTS]
    traces.append(go.Scatter(x=rx, y=ry, mode="markers", hoverinfo="skip",
                             showlegend=False, marker=dict(size=22, color=C_RADIO)))
    traces.append(go.Scatter(x=rx, y=ry, mode="markers", hoverinfo="skip",
                             showlegend=False, marker=dict(size=14, color="#FFFFFF")))
    traces.append(go.Scatter(x=[rx[0]], y=[ry[0]], mode="markers", hoverinfo="skip",
                             showlegend=False, marker=dict(size=11, color=C_DARK)))
    traces.append(go.Scatter(x=[rx[1]], y=[ry[1]], mode="markers", hoverinfo="skip",
                             showlegend=False, marker=dict(size=11, color=C_DARK),
                             visible=False))

    return traces


# ==========================================================================
# 5. 按进度 p 计算各元素的坐标 / 文字
# ==========================================================================
def state_at(p):
    """给定进度，算出所有需要随状态变化的量。"""
    vals = interp(p)
    angs = sector_angles(vals)

    rings = []      # (theta 列表, width 列表, customdata)
    seps = []       # (theta 列表, r 列表)  ← 极坐标，给 Scatterpolar 用
    lines = []      # (x 列表, y 列表)
    dots = []       # (x, y)
    labels = []     # (x, y, text)

    for i, seg in enumerate(SEGMENTS):
        phi0, sweep = angs[i]
        center = math.radians(phi0 + sweep / 2.0)
        width = math.radians(sweep)
        rings.append(([center] * RINGS, [width] * RINGS,
                      [f'{seg["name"]}　{vals[i] * 100:.0f}%'] * RINGS))

        cols = ring_colors(seg)
        ts, rs = [], []
        for k in range(1, seg["rings"]):
            at, ar = arc_polar(R_IN + k * RING_W, phi0, sweep)
            ts.extend(at + [None])
            rs.extend(ar + [None])
        seps.append((ts, rs))

        a = LAYOUT["近7天"][seg["name"]]
        b = LAYOUT["近30天"][seg["name"]]
        lx = [p0[0] + (p1[0] - p0[0]) * p for p0, p1 in zip(a["line"], b["line"])]
        ly = [p0[1] + (p1[1] - p0[1]) * p for p0, p1 in zip(a["line"], b["line"])]
        lines.append((lx, ly))
        dots.append((lx[-1], ly[-1]))

        tx = a["text"][0] + (b["text"][0] - a["text"][0]) * p
        ty = a["text"][1] + (b["text"][1] - a["text"][1]) * p
        labels.append((tx, ty + _desc(FS_LABEL),
                       f'{seg["name"]}, {vals[i] * 100:.0f}%'))

    return rings, seps, lines, dots, labels


def annotations_at(p):
    """标题 + 脚注 + 五条数据标签 + 两个按钮文字。"""
    _, _, _, _, labels = state_at(p)
    ann = [
        dict(x=TITLE_POS[0] + X_SHIFT, y=TITLE_POS[1] + _desc(FS_TITLE),
             text=f"<b>{TITLE_TEXT}</b>",
             showarrow=False, xanchor="left", yanchor="bottom", xref="x", yref="y",
             font=dict(family=FONT_TITLE, size=FS_TITLE, color=C_TITLE)),
        dict(x=NOTE_POS[0] + X_SHIFT, y=NOTE_POS[1] + _desc(FS_NOTE), text=NOTE_TEXT,
             showarrow=False, xanchor="left", yanchor="bottom", xref="x", yref="y",
             font=dict(family=FONT_TEXT, size=FS_NOTE, color=C_NOTE)),
    ]
    for (tx, ty, text) in labels:
        ann.append(dict(x=tx + X_SHIFT, y=ty, text=text, showarrow=False,
                        xanchor="left", yanchor="bottom", xref="x", yref="y",
                        font=dict(family=FONT_TEXT, size=FS_LABEL, color=C_LABEL)))
    for name, (bx, by, bw, bh) in BTN_RECTS.items():
        ann.append(dict(x=bx + 39, y=by + bh / 2.0 + 1, text=name, showarrow=False,
                        xanchor="left", yanchor="middle", xref="x", yref="y",
                        font=dict(family=FONT_TEXT, size=FS_BTN, color=C_DARK)))
    return ann


def shapes_static():
    """两个按钮的灰底矩形（不随状态变化）。

    layer 必须是 below：按钮的单选圆点是用 scatter 画的，若灰块画在
    trace 之上会把圆点整个盖掉。
    """
    out = []
    for name, (bx, by, bw, bh) in BTN_RECTS.items():
        out.append(dict(type="rect", xref="x", yref="y",
                        x0=bx, y0=by, x1=bx + bw, y1=by + bh,
                        fillcolor=C_BTN, line=dict(width=0), layer="below"))
    return out


# ==========================================================================
# 6. 组装 figure
# ==========================================================================
NFRAME = 12


def build_figure(p0=0.0):
    """p0 为初始进度，0=近7天，1=近30天。"""
    fig = go.Figure()
    for tr in build_traces():
        fig.add_trace(tr)

    # 初始状态
    rings, seps, lines, dots, _ = state_at(p0)
    for i in range(N_SEG):
        th, wd, cd = rings[i]
        fig.data[IDX_RING[i]].theta = th
        fig.data[IDX_RING[i]].width = wd
        fig.data[IDX_RING[i]].customdata = cd
        fig.data[IDX_SEP[i]].theta = seps[i][0]
        fig.data[IDX_SEP[i]].r = seps[i][1]
        fig.data[IDX_LINE[i]].x, fig.data[IDX_LINE[i]].y = lines[i]
        fig.data[IDX_DOT[i]].x = [dots[i][0]]
        fig.data[IDX_DOT[i]].y = [dots[i][1]]

    # 选中态：黑心只看当前状态
    fig.data[IDX_DOT7].visible = p0 < 0.5
    fig.data[IDX_DOT30].visible = p0 >= 0.5

    # --- 极坐标：domain 精确对齐到像素圆心与半径 ---
    half_w, half_h = R_OUT / W, R_OUT / H
    fig.update_layout(
        width=W, height=H,
        margin=dict(l=0, r=0, t=0, b=0),
        paper_bgcolor=BG, plot_bgcolor=BG,
        showlegend=False,
        hovermode="closest",
        dragmode=False,
        polar=dict(
            domain=dict(x=[CX / W - half_w, CX / W + half_w],
                        y=[1 - CY / H - half_h, 1 - CY / H + half_h]),
            bgcolor=BG,
            radialaxis=dict(range=[0, R_OUT], visible=False, showgrid=False,
                            showline=False, ticks=""),
            angularaxis=dict(rotation=90, direction="clockwise", visible=False,
                             showgrid=False, showline=False, ticks=""),
        ),
        xaxis=dict(range=[0, W], domain=[0, 1], visible=False, fixedrange=True,
                   showgrid=False, zeroline=False, showline=False, ticks=""),
        yaxis=dict(range=[H, 0], domain=[0, 1], visible=False, fixedrange=True,
                   showgrid=False, zeroline=False, showline=False, ticks=""),
        annotations=annotations_at(p0),
        shapes=shapes_static(),
    )
    return fig


def frame_payloads(p0=0.0):
    """逐帧数据，交给页面里的 JS 做补间。

    不用 plotly 自带的 frames / Plotly.animate：实测 animate 在无头浏览器里
    推进不动，而且 newPlot 时还会把第一帧的 layout 盖到图上。改成自己算好
    每一帧、再用 Plotly.restyle / relayout 一帧一帧推，行为完全可预期。

    帧序列从 p0 出发走到另一端（p0=0 时 0→1，p0=1 时 1→0）。
    """
    out = []
    for k in range(NFRAME):
        t = ease(k / (NFRAME - 1.0))
        p = p0 + (1.0 - 2.0 * p0) * t
        rings, seps, lines, dots, _ = state_at(p)
        out.append(dict(
            theta=[rings[i][0] for i in range(N_SEG)],
            width=[rings[i][1] for i in range(N_SEG)],
            st=[seps[i][0] for i in range(N_SEG)],      # 分隔弧：极坐标 theta
            sr=[seps[i][1] for i in range(N_SEG)],      # 分隔弧：半径
            sx=[lines[i][0] for i in range(N_SEG)]      # 引线 + 圆点：像素坐标
               + [[dots[i][0]] for i in range(N_SEG)],
            sy=[lines[i][1] for i in range(N_SEG)]
               + [[dots[i][1]] for i in range(N_SEG)],
            ann=annotations_at(p),
        ))
    return out


# ==========================================================================
# 7. 交互脚本（按钮热区 + 点击切状态）
# ==========================================================================
_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLOT_ID__');
  if (!gd || !gd.on) return;
  // 用绘图 div 作定位基准：它左上角就在页面 (0,0)，margin 已被 CSS 清零
  gd.style.position = 'relative';

  var FR = __FRAMES__;          // 逐帧数据，首帧 == 当前状态
  var IR = __IDX_RING__;        // 环带 trace 索引（极坐标）
  var ISP = __IDX_SEP__;        // 分隔弧 trace 索引（极坐标）
  var IS = __IDX_XY__;          // 引线 / 圆点 trace 索引（直角坐标）
  var N = FR.length;
  var busy = false;

  function apply(k) {
    var f = FR[k];
    Plotly.restyle(gd, {theta: f.theta, width: f.width}, IR);
    Plotly.restyle(gd, {theta: f.st, r: f.sr}, ISP);
    Plotly.restyle(gd, {x: f.sx, y: f.sy}, IS);
    Plotly.relayout(gd, {annotations: f.ann});
  }

  function play(dir) {
    if (busy) return;
    busy = true;
    var k = dir > 0 ? 1 : N - 2;
    var end = dir > 0 ? N - 1 : 0;
    function tick() {
      apply(k);
      if (k === end) { busy = false; return; }
      k += dir;
      setTimeout(tick, 26);
    }
    setTimeout(tick, 26);
  }

  var BTNS = [
    {id: '__PLOT_ID__-btn7',  rect: [639, 121, 127, 55], dir: __DIR7__,
     dots: [true, false], tip: '切换到近7天'},
    {id: '__PLOT_ID__-btn30', rect: [783, 120, 127, 55], dir: __DIR30__,
     dots: [false, true], tip: '切换到近30天'}
  ];

  BTNS.forEach(function (b) {
    var d = document.createElement('div');
    d.id = b.id;
    d.title = b.tip;
    d.style.cssText = 'position:absolute;cursor:pointer;' +
      'left:' + b.rect[0] + 'px;top:' + b.rect[1] + 'px;' +
      'width:' + b.rect[2] + 'px;height:' + b.rect[3] + 'px;';
    d.addEventListener('click', function () {
      play(b.dir);
      Plotly.restyle(gd, {visible: b.dots}, [__IDX_DOT7__, __IDX_DOT30__]);
    });
    gd.appendChild(d);
  });
})();
"""


def _post_script(payloads, p0=0.0):
    """把逐帧数据和播放方向塞进页面脚本。

    dir=+1 沿帧序走向末帧，dir=-1 回到首帧；p0 决定哪一端对应哪个按钮。
    """
    dir7 = -1 if p0 == 0 else 1
    dir30 = 1 if p0 == 0 else -1
    return (_POST_SCRIPT
            .replace("__PLOT_ID__", DIV_ID)
            .replace("__FRAMES__", json.dumps(payloads, ensure_ascii=False))
            .replace("__IDX_RING__", str(IDX_RING))
            .replace("__IDX_SEP__", str(IDX_SEP))
            .replace("__IDX_XY__", str(IDX_LINE + IDX_DOT))
            .replace("__DIR7__", str(dir7))
            .replace("__DIR30__", str(dir30))
            .replace("__IDX_DOT7__", str(IDX_DOT7))
            .replace("__IDX_DOT30__", str(IDX_DOT30)))


CSS_FIX = (
    "<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
    ".plotly-graph-div{overflow:visible !important;}</style>"
)


# ==========================================================================
# 8. 输出
# ==========================================================================
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图\第三章"
STEM = "3 动态南丁格尔圆环图"
DIV_ID = "ndrose"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


def save_html(path, p0=0.0):
    fig = build_figure(p0)
    html = fig.to_html(include_plotlyjs="inline", full_html=True, config=CONFIG,
                       div_id=DIV_ID, post_script=_post_script(frame_payloads(p0), p0))
    # 干掉 body 默认的 8px 边距，否则整张图会偏 (8, 8)
    html = html.replace("</head>", CSS_FIX + "</head>")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return path


def save_png(html_path, png_path, timeout=90):
    """用 Chrome 无头模式把 HTML 截成 PNG（与浏览器所见一致）。

    每次都要用一个全新的 --user-data-dir：复用同一个配置目录时 Chrome 会
    命中上一次的渲染缓存，导致连续截出来的是同一张图。
    """
    if not os.path.exists(CHROME):
        print("[提示] 未找到 Chrome，跳过 PNG 导出。")
        return None
    url = "file:///" + quote(os.path.abspath(html_path).replace("\\", "/"))
    if os.path.exists(png_path):
        os.remove(png_path)
    profile = tempfile.mkdtemp(prefix="nd-rose-")
    try:
        # 不能加 --virtual-time-budget：它让 Chrome 用虚拟时钟快进，会把
        # figure 里的 frames 提前应用掉，截出来的是动画第 2 帧而不是初态。
        subprocess.run([
            CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
            "--no-first-run", "--no-default-browser-check",
            "--force-device-scale-factor=1",
            f"--user-data-dir={profile}",
            f"--window-size={W},{H}",
            "--run-all-compositor-stages-before-draw",
            f"--screenshot={png_path}", url,
        ], capture_output=True, timeout=timeout)
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    return png_path if os.path.exists(png_path) else None


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)

    html_path = os.path.join(OUT_DIR, f"{STEM}_plotly.html")
    print("[OK] 已保存", save_html(html_path, 0.0))

    for period, p in (("近7天", 0.0), ("近30天", 1.0)):
        tmp_html = os.path.join(OUT_DIR, f"_{period}.html")
        save_html(tmp_html, p)
        png = os.path.join(OUT_DIR, f"{STEM}_{period}_plotly.png")
        if save_png(tmp_html, png):
            print("[OK] 已保存", png)
        os.remove(tmp_html)

    print("\n用浏览器打开：", html_path)
