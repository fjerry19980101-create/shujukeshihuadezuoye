# -*- coding: utf-8 -*-
"""
第三章 动态图表 —— 1 动态柱形图（plotly 可交互版）
====================================================================

数据与「动态」的来源
    工作簿里是 6 个月 × 6 个区域 的源数据（B3:H9），另有一个「控件单元格」
    J3，值为 5。图表区（B12:H13）用 INDEX/VLOOKUP 从源数据里取出第 J3 行，
    所以改 J3 就等于换月份——这就是 Excel 里那张柱形图「动」起来的机制。
    这里照搬同一套逻辑：读 B3:H9 和 J3，J3 决定初始月份。

版面参数全部量自参考截图（950×689）
    · 背景 #1A1E43；柱 #0097E0；网格 #454866（2px，虚实 16/6）；
    · 绘图区 x∈[131.3,889.7]，数值 0→y=567、800→y=193（800 单位 = 374px）；
    · 柱宽 50.5px ≈ 0.403 个类目带宽；
    · 标题 #FFFFFF 微软雅黑 Bold；刻度/类目 #F2F2F2；脚注 #D9D9D9；后两者等线。

纵轴为什么是「自适应」的
    参考图对应 5 月（最大值 705），Excel 自动定轴到 800，截图就是 800。
    但 6 月的数据有 965 和 897，固定 800 会把柱子截断——Excel 遇到这种情况
    同样会把轴上调。所以这里按 Excel 的行为来：主刻度固定 200，轴顶取
    max(800, 上取整到 200 的倍数)。1~5 月都是 800（和截图一致），6 月是 1000。

关于按钮
    参考截图里没有控件——Excel 的微调按钮放在图表对象之外，截图只截了图表。
    所以：
      · 导出的 PNG 不带任何控件，逐像素对齐参考图；
      · HTML 版在标题右侧的空白带（y≈130~170，参考图这里是空的）补一排
        月份选择器，沿用本章另一张图（3 动态南丁格尔圆环图）的视觉语言：
        灰块 + 三环单选圆点，选中项才有黑心。

输出（保存在 图/第三章/）
    1 动态柱形图_plotly.html          ← 双击打开，点月份切换，柱高做补间
    1 动态柱形图-1月.png … -6月.png   ← 六张静态图，与同目录「2 动态跑道图-1月」
                                        的命名风格一致
"""

import json
import math
import os
import shutil
import subprocess
import tempfile
from urllib.parse import quote

import plotly.graph_objects as go

# ==========================================================================
# 1. 版面参数（单位：像素，原点左上，y 向下）
# ==========================================================================
W, H = 950, 689

PLOT_L, PLOT_R = 131.3, 889.7     # 绘图区左右边界
Y_V0, Y_VMAX = 567.0, 193.0       # 数值 0 / 轴顶 对应的像素 y（轴顶=800 时）

BG = "#1A1E43"
C_TITLE = "#FFFFFF"
C_LABEL = "#F2F2F2"
C_NOTE = "#D9D9D9"
C_BAR = "#0097E0"
C_GRID = "#454866"
C_BTN = "#8E8F9A"
C_RADIO = "#646464"
C_DARK = "#000000"

GRID_W = 2.0
GRID_DASH = "16px,6px"            # 参考图实测：16px 实线 + 6px 空档
GRID_STEP = 200                   # 主刻度间隔
BAR_WIDTH = 0.403                 # 0.403 × 类目带宽(126.4px) ≈ 50.9px

FS_TITLE, FS_TICK, FS_CAT, FS_NOTE = 48.0, 21.0, 24.0, 21.3
FONT_TITLE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_TEXT = '"DengXian", "等线", "Microsoft YaHei", sans-serif'
# 刻度数字实测比等线更宽（37px vs 35px @ 同高），和微软雅黑更接近
FONT_TICK = '"Microsoft YaHei", "微软雅黑", sans-serif'

TITLE_TEXT = "2022年上半年各区域销售情况"
# 参考图用的是全角冒号「：」，不是 ASCII 冒号 + 空格
NOTE_TEXT = "*注：数据来源公司销售系统"

TITLE_INK = (73.0, 61.0, 108.0)   # 文字墨迹 左 / 上 / 下
NOTE_INK = (72.0, 634.0, 654.0)
CAT_INK_BOT = 606.0               # 类目标签墨迹下沿

# 三个文字锚点的微调量（渲染后量墨迹反推）
TITLE_X_SHIFT = -4.5
X_SHIFT = -2.5
TITLE_DY = -1.0
TICK_RIGHT = 115.0                # 刻度标签右边界
TICK_DX = 3.0
TICK_DY = -1.0
CAT_DY = 5.0

# 类目位置微调（正数会把柱子整体左移，这里量下来不需要）
X_NUDGE = 0.0


def _desc(size):
    """墨迹下沿 → annotation「bottom」锚点。0.20em 是实测的字体下伸量。"""
    return 0.20 * size


# ==========================================================================
# 2. 数据
# ==========================================================================
CATS = ["华北", "华南", "东北", "西北", "西南", "华东"]

FALLBACK = {
    "1月": [259, 247, 705, 513, 463, 384],
    "2月": [235, 266, 529, 324, 579, 282],
    "3月": [330, 342, 705, 540, 492, 308],
    "4月": [282, 342, 599, 270, 492, 384],
    "5月": [282, 323, 705, 486, 405, 308],
    "6月": [965, 380, 282, 567, 463, 897],
}
FALLBACK_MONTHS = list(FALLBACK)
FALLBACK_CTRL = 5

MONTHS, DATA, CTRL = FALLBACK_MONTHS, FALLBACK, FALLBACK_CTRL


def load_from_xlsx(path):
    """读源数据 B3:H9 与控件单元格 J3，返回 (月份列表, {月份: 六区域数值}, J3)。"""
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).worksheets[0]
    cats = [str(ws.cell(3, c).value).strip() for c in range(3, 9)]
    months, data = [], {}
    for r in range(4, 10):
        m = str(ws.cell(r, 2).value).strip()
        months.append(m)
        data[m] = [float(ws.cell(r, c).value) for c in range(3, 9)]
    ctrl = int(ws.cell(3, 10).value or 1)
    if cats != CATS:
        print(f"[提示] 表头与内置不一致：{cats}，仍按表内顺序绘制。")
    return months, data, ctrl


_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), "1 动态柱形图.xlsx")
if os.path.exists(_XLSX):
    try:
        MONTHS, DATA, CTRL = load_from_xlsx(_XLSX)
    except Exception as exc:
        print(f"[提示] 读取 xlsx 失败（{exc}），使用脚本内置数据。")

N_MONTH = len(MONTHS)
CTRL = max(1, min(CTRL, N_MONTH))          # J3 是 1 基的月份序号


def axis_max(vals):
    """按 Excel 的习惯定轴顶：主刻度 200，轴顶取到 200 的倍数，且不低于 800。"""
    return float(max(800, math.ceil(max(vals) / float(GRID_STEP)) * GRID_STEP))


def y_of(v, ymax):
    """数值 → 像素 y。"""
    return Y_V0 - (Y_V0 - Y_VMAX) * (v / ymax)


def px(x, y):
    """像素 → paper 坐标（0~1）。"""
    return x / W, 1.0 - y / H


# ==========================================================================
# 3. 形状与文字
# ==========================================================================
def rect_shape(x, y, w, h, fill):
    x0, y0 = px(x, y)
    x1, y1 = px(x + w, y + h)
    return dict(type="rect", xref="paper", yref="paper",
                x0=x0, x1=x1, y0=y1, y1=y0,
                fillcolor=fill, line=dict(width=0), layer="below")


def circle_shape(cx, cy, r, fill):
    x0, y0 = px(cx - r, cy - r)
    x1, y1 = px(cx + r, cy + r)
    return dict(type="circle", xref="paper", yref="paper",
                x0=x0, x1=x1, y0=y1, y1=y0,
                fillcolor=fill, line=dict(width=0), layer="below")


def gridline_shapes(ymax):
    """水平虚线网格。

    没用 yaxis 自带的 showgrid：plotly 不画轴两端的网格线（0 和轴顶都缺），
    线宽和虚线相控也不如自己画可控。改成 shape 后位置、粗细、虚实都精确。
    各线的像素 y 再取整——参考图上每条线都被吸附到整数像素（286.5→286、
    473.5→474），取整后 800 那一档正好复现量到的 193/286/380/474/567。
    """
    x0, _ = px(PLOT_L, 0.0)
    x1, _ = px(PLOT_R, 0.0)
    out = []
    v = 0.0
    while v <= ymax + 1e-9:
        yy = 1.0 - int(round(y_of(v, ymax))) / H
        out.append(dict(type="line", xref="paper", yref="paper",
                        x0=x0, x1=x1, y0=yy, y1=yy,
                        line=dict(color=C_GRID, width=GRID_W, dash=GRID_DASH),
                        layer="below"))
        v += GRID_STEP
    return out


def band_center(i):
    """第 i 个类目带的中心 x（像素）。"""
    w = (PLOT_R - PLOT_L) / N_MONTH
    return PLOT_L + (i + 0.5) * w


def axis_annotations(ymax):
    """刻度值与类目标签。

    plotly 5.24 的 ticklen 只画刻度线、不挪标签位置，靠它把标签推到量到的
    位置推不动。干脆关掉 showticklabels，全部用 annotation 摆，坐标就是
    截图上量到的像素。
    """
    out = []
    v = 0.0
    while v <= ymax + 1e-9:
        gy = Y_V0 - (Y_V0 - Y_VMAX) * (v / ymax)
        out.append(dict(x=(TICK_RIGHT + TICK_DX) / W,
                        y=1.0 - (gy + TICK_DY) / H,
                        text=str(int(v)), showarrow=False,
                        xanchor="right", yanchor="middle",
                        xref="paper", yref="paper",
                        font=dict(family=FONT_TICK, size=FS_TICK, color=C_LABEL)))
        v += GRID_STEP
    for i, c in enumerate(CATS):
        out.append(dict(x=band_center(i) / W,
                        y=1.0 - (CAT_INK_BOT + CAT_DY) / H,
                        text=c, showarrow=False,
                        xanchor="center", yanchor="bottom",
                        xref="paper", yref="paper",
                        font=dict(family=FONT_TEXT, size=FS_CAT, color=C_LABEL)))
    return out


def head_annotations():
    """标题 + 脚注（与月份无关）。"""
    tx, _, ty1 = TITLE_INK
    nx, _, ny1 = NOTE_INK
    return [
        dict(x=(tx + TITLE_X_SHIFT) / W,
             y=1.0 - (ty1 + _desc(FS_TITLE) + TITLE_DY) / H,
             text=f"<b>{TITLE_TEXT}</b>",
             showarrow=False, xanchor="left", yanchor="bottom",
             xref="paper", yref="paper",
             font=dict(family=FONT_TITLE, size=FS_TITLE, color=C_TITLE)),
        dict(x=(nx + X_SHIFT) / W,
             y=1.0 - (ny1 + _desc(FS_NOTE)) / H,
             text=NOTE_TEXT,
             showarrow=False, xanchor="left", yanchor="bottom",
             xref="paper", yref="paper",
             font=dict(family=FONT_TEXT, size=FS_NOTE, color=C_NOTE)),
    ]


# ==========================================================================
# 4. 月份选择器（只在 HTML 版出现）
# ==========================================================================
CHIP_W, CHIP_H, CHIP_GAP = 70.0, 40.0, 6.0
CHIP_Y = 130.0
CHIP_X0 = PLOT_R - (N_MONTH * CHIP_W + (N_MONTH - 1) * CHIP_GAP)


def chip_rect(i):
    return (CHIP_X0 + i * (CHIP_W + CHIP_GAP), CHIP_Y, CHIP_W, CHIP_H)


def chip_shapes(active):
    """灰块 + 三环单选圆点；只有当前月份那枚有黑心。"""
    out = []
    for i in range(N_MONTH):
        x, y, w, h = chip_rect(i)
        cy = y + h / 2.0
        out.append(rect_shape(x, y, w, h, C_BTN))
        out.append(circle_shape(x + 16.0, cy, 7.0, C_RADIO))
        out.append(circle_shape(x + 16.0, cy, 4.5, "#FFFFFF"))
        if i == active:
            out.append(circle_shape(x + 16.0, cy, 2.8, C_DARK))
    return out


def chip_annotations(active):
    out = []
    for i, m in enumerate(MONTHS):
        x, y, w, h = chip_rect(i)
        ax, ay = px(x + 30.0, y + h / 2.0 + 0.5)
        out.append(dict(x=ax, y=ay, text=m, showarrow=False,
                        xanchor="left", yanchor="middle",
                        xref="paper", yref="paper",
                        font=dict(family=FONT_TEXT, size=17.5, color=C_DARK)))
    return out


# ==========================================================================
# 5. 组装 figure
# ==========================================================================
def state_parts(month_idx, chips=False):
    """某个月份对应的 (shapes, annotations, y 轴 range)。"""
    ymax = axis_max(DATA[MONTHS[month_idx]])
    shapes = (gridline_shapes(ymax) + chip_shapes(month_idx)) if chips \
        else gridline_shapes(ymax)
    ann = head_annotations() + axis_annotations(ymax)
    if chips:
        ann += chip_annotations(month_idx)
    return shapes, ann, [0, ymax]


def build_figure(month_idx, chips=False):
    """month_idx 是 0 基的月份下标。chips=True 时画出月份选择器。"""
    vals = DATA[MONTHS[month_idx]]
    shapes, ann, yrange = state_parts(month_idx, chips)

    fig = go.Figure(go.Bar(
        x=CATS, y=vals,
        width=BAR_WIDTH,
        marker=dict(color=C_BAR, line=dict(width=0)),
        hovertemplate="<b>%{x}</b>　%{y:.0f}<extra></extra>",
    ))

    fig.update_layout(
        width=W, height=H,
        # autoexpand 必须关掉：默认开着时 plotly 会为了容纳溢出画布的标签
        # 自行放大 margin，绘图区被压缩，domain 就不再对应真实像素了。
        margin=dict(l=0, r=0, t=0, b=0, autoexpand=False),
        paper_bgcolor=BG, plot_bgcolor=BG,
        showlegend=False,
        hovermode="closest",
        dragmode=False,
        xaxis=dict(
            type="category",
            domain=[PLOT_L / W, PLOT_R / W],
            range=[-0.5 + X_NUDGE, N_MONTH - 0.5 + X_NUDGE],
            showgrid=False, zeroline=False, showline=False,
            ticks="", showticklabels=False,
            fixedrange=True,
        ),
        yaxis=dict(
            domain=[1.0 - Y_V0 / H, 1.0 - Y_VMAX / H],
            range=yrange,
            showgrid=False, zeroline=False, showline=False,
            ticks="", showticklabels=False,
            fixedrange=True,
        ),
        annotations=ann,
        shapes=shapes,
    )
    return fig


# ==========================================================================
# 6. 交互脚本
# ==========================================================================
NFRAME = 12   # 每段补间的帧数（页面里再按 ease 插值）


_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLOT_ID__');
  if (!gd || !gd.on) return;
  gd.style.position = 'relative';

  var V = __DATA__;           // V[月] = 6 个数值
  var SH = __SHAPES__;        // SH[月] = 该月的 shapes（网格 + 灰块 + 单选圆点）
  var AN = __ANN__;           // AN[月] = 该月的 annotations
  var YR = __YRANGE__;        // YR[月] = [0, 轴顶]
  var HOT = __HOTS__;         // [月份] -> [x, y, w, h, 名称]
  var N = __NFRAME__;
  var STEP = 28;              // 帧间隔 ms
  var cur = __CUR__;          // 当前月份（0 基）
  var busy = false;

  function ease(t) { return t * t * (3.0 - 2.0 * t); }

  function play(dst) {
    if (busy || dst === cur) return;
    busy = true;
    // 轴顶变了就先换坐标轴/网格/标签，再补间柱高
    Plotly.relayout(gd, {shapes: SH[dst], annotations: AN[dst],
                         'yaxis.range': YR[dst]});
    var a = V[cur], b = V[dst];
    var k = 1;
    function tick() {
      var t = ease(k / (N - 1.0));
      var y = [];
      for (var j = 0; j < a.length; j++) y.push(a[j] + (b[j] - a[j]) * t);
      Plotly.restyle(gd, {y: [y]}, [0]);
      if (k === N - 1) { cur = dst; busy = false; return; }
      k += 1;
      setTimeout(tick, STEP);
    }
    setTimeout(tick, STEP);
  }

  HOT.forEach(function (r, i) {
    var d = document.createElement('div');
    d.id = '__PLOT_ID__-m' + (i + 1);
    d.title = '切换到 ' + r[4];
    d.style.cssText = 'position:absolute;cursor:pointer;' +
      'left:' + r[0] + 'px;top:' + r[1] + 'px;' +
      'width:' + r[2] + 'px;height:' + r[3] + 'px;';
    d.addEventListener('click', function () { play(i); });
    gd.appendChild(d);
  });

  // 键盘 1~6 也能切
  document.addEventListener('keydown', function (e) {
    var n = parseInt(e.key, 10);
    if (n >= 1 && n <= HOT.length) play(n - 1);
  });
})();
"""

CSS_FIX = (
    "<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
    ".plotly-graph-div{overflow:visible !important;}</style>"
)

DIV_ID = "barfig"
CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


def _post_script():
    hots = []
    for i, m in enumerate(MONTHS):
        x, y, w, h = chip_rect(i)
        hots.append([x, y, w, h, m])
    parts = [state_parts(i, chips=True) for i in range(N_MONTH)]
    return (_POST_SCRIPT
            .replace("__PLOT_ID__", DIV_ID)
            .replace("__NFRAME__", str(NFRAME))
            .replace("__DATA__", json.dumps([DATA[m] for m in MONTHS]))
            .replace("__SHAPES__", json.dumps([p[0] for p in parts],
                                              ensure_ascii=False))
            .replace("__ANN__", json.dumps([p[1] for p in parts],
                                           ensure_ascii=False))
            .replace("__YRANGE__", json.dumps([p[2] for p in parts]))
            .replace("__HOTS__", json.dumps(hots, ensure_ascii=False))
            .replace("__CUR__", str(CTRL - 1)))


def save_html(path, month_idx=None):
    idx = CTRL - 1 if month_idx is None else month_idx
    fig = build_figure(idx, chips=True)
    html = fig.to_html(include_plotlyjs="inline", full_html=True, config=CONFIG,
                       div_id=DIV_ID, post_script=_post_script())
    html = html.replace("</head>", CSS_FIX + "</head>")
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    return path


# ==========================================================================
# 7. 输出
# ==========================================================================
OUT_DIR = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化"
           r"\新建文件夹\图\第三章")
STEM = "1 动态柱形图"
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"


def save_png(html_path, png_path, timeout=90):
    if not os.path.exists(CHROME):
        print("[提示] 未找到 Chrome，跳过 PNG 导出。")
        return None
    url = "file:///" + quote(os.path.abspath(html_path).replace("\\", "/"))
    if os.path.exists(png_path):
        os.remove(png_path)
    profile = tempfile.mkdtemp(prefix="nd-bar-")
    try:
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
    print("[OK] 已保存", save_html(html_path))

    # 静态图不带控件，逐像素对齐参考截图
    for i, m in enumerate(MONTHS):
        tmp_html = os.path.join(OUT_DIR, f"_{m}.html")
        fig = build_figure(i, chips=False)
        html = fig.to_html(include_plotlyjs="inline", full_html=True,
                           config=CONFIG, div_id=DIV_ID)
        html = html.replace("</head>", CSS_FIX + "</head>")
        with open(tmp_html, "w", encoding="utf-8") as fh:
            fh.write(html)
        png = os.path.join(OUT_DIR, f"{STEM}-{m}.png")
        if save_png(tmp_html, png):
            print("[OK] 已保存", png)
        os.remove(tmp_html)

    print("\n用浏览器打开：", html_path)
