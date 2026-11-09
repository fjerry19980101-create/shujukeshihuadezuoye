# -*- coding: utf-8 -*-
"""
「5 人力资源明细 / 透视表切片器」plotly 版：
按学历汇总平均月收入，柱体系竖向渐变，右上角是一块可点击的「部门」切片器。

参考基准：948 x 680（本机 Excel 导出图表 → 等比缩放到 948x680）
数据源  ：第三章 动态图表\\5 人力资源明细.xlsx（A1:G1471 明细，1470 条）
          └ 图表本身在「5 透视表切片器.xlsx」里，是「数据透视表 + 切片器」驱动的：
            行=学历、值=月收入平均值、切片器=部门。
输出    ：...\\作业\\图\\第三章\\5 人力资源明细_plotly.html
          ...\\作业\\图\\第三章\\5 人力资源明细_plotly.png

一、这张图的结构（拆自 5 透视表切片器.xlsx 的 chart1.xml / drawing2.xml）
    · 单系列「汇总」的簇状柱形，gapWidth=219、overlap=-28（单系列，overlap 无意义）
    · 柱体是**竖向线性渐变** #0070C0 → #00B0F0，而且是**每根柱各自铺满**（不是整幅图共用）
    · 数值标签：outEnd、9pt、颜色 bg1×95% = #F2F2F2、格式 0.00
    · 坐标轴：纵轴 delete=1（完全不显示）；横轴只留一条线 + 外向刻度线，标签 10pt
    · 网格线：白色 95% @20% 透明度、**长虚线**，且画在柱子**下面**
    · 图表区背景 #1A1E43
    · 标题与脚注都不在图表里，而是图表的「用户形状」（drawing2.xml）：
      标题「各学历平均工资情况」微软雅黑 18pt 加粗、纯白；脚注微软雅黑 8pt、bg1×85% = #D9D9D9

二、像素几何（单位 px，画布 948x680）
    · 绘图区 x 61..881，y 178.5..559；纵轴量程 0..10000（Excel 自动取整），主刻度 2500
    · 5 个类目槽各 164 px，柱心 143 / 307 / 471 / 635 / 799，柱宽 52（= 219% gapWidth 的换算值）
    · 网格线 5 条（0/2500/5000/7500/10000）在 y=559/463.5/368.5/273.5/178.5，
      实测 2px 高、实 15px 空 5px；横轴刻度线 6 根在 x=61+164k、y 560..566

三、三处「照抄」的细节
    · **每根柱自己的渐变**：用 SVG linearGradient 的 objectBoundingBox 坐标系，
      柱高变化时渐变自动跟着柱身重算，不用逐帧改。
    · **网格线画在柱子下面**：实测柱子内部在网格线所在行仍是纯净渐变，
      所以网格线放进一个插在 plotly SVG 之前的背板层，靠 DOM 顺序压在下面。
    · 纵轴完全不显示（连轴线都没有），只有横轴一条线 + 6 根外向刻度线，全用 DOM 画。

四、交互
    右上角是「部门」切片器，对应源文件的 slicerCache1.xml。7 个部门可多选：
    点一下选中/取消，**一个都不选 = 全部部门**（这正是 Excel 切片器的行为）。
    切换时柱高按 10 帧 / 26ms 做缓入缓出补间，数值标签跟着柱顶走、最后一帧再换数字。
    默认状态与源文件一致：只选「市场拓展」，所以首屏数值就是 6138.32/7187.00/... 那一组。
"""
import json
import os
import shutil
import subprocess
import tempfile
from urllib.parse import quote

import plotly.graph_objects as go

# ==========================================================================
# 0. 路径
# ==========================================================================
BASE = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\作业")
XLSX = os.path.join(BASE, "第三章 动态图表", "5 人力资源明细.xlsx")
OUT_DIR = os.path.join(BASE, "图", "第三章")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# ==========================================================================
# 1. 数据
# ==========================================================================
# 透视表的行标签顺序（按拼音排，Excel 就是这么排的，别改成自定义顺序）
LEVELS = ["本科", "博士", "高中", "硕士", "专科"]
# 切片器里的部门顺序，取自 slicerCache1.xml 的 items 序列（也是拼音序）
DEPTS = ["财务", "工程", "后勤", "人力资源", "市场拓展", "销售", "行政"]
DEPT_DEFAULT = ["市场拓展"]          # 源文件打开时的切片器状态

TITLE = "各学历平均工资情况"
NOTE = "*注：数据来源于人力资源管理系统，统计日期截至2022.03.31"


def load_data(path):
    """读明细，返回 {(部门, 学历): (求和, 条数)}。失败则抛错（不内置兜底，避免假数据）。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    it = ws.iter_rows(values_only=True)
    hdr = list(next(it))
    idep, ilv, isal = hdr.index("部门"), hdr.index("学历"), hdr.index("月收入")
    bucket = {}
    for row in it:
        if row[0] is None:
            continue
        key = (row[idep], row[ilv])
        s, n = bucket.get(key, (0.0, 0))
        bucket[key] = (s + float(row[isal]), n + 1)
    wb.close()
    return bucket


BUCKET = load_data(XLSX)


def agg(depts):
    """按部门列表汇总，返回 5 个学历的 (求和, 条数)。

    **必须按行汇总后再除**——透视表的平均值是对所有匹配行求平均，
    不是「各部门均值再平均」（各部门人数不同，两者不等）。
    """
    sel = set(depts)
    sums, cnts = [0.0] * len(LEVELS), [0] * len(LEVELS)
    for (dep, lv), (s, n) in BUCKET.items():
        if lv in LEVELS and (not sel or dep in sel):
            i = LEVELS.index(lv)
            sums[i] += s
            cnts[i] += n
    return sums, cnts


def averages(depts):
    sums, cnts = agg(depts)
    return [s / n if n else 0.0 for s, n in zip(sums, cnts)]


# ==========================================================================
# 2. 画布与几何
# ==========================================================================
W, H = 948, 680
BG = "#1A1E43"

PX0, PX1 = 61.0, 881.0            # 绘图区左右
PY0, PY1 = 178.5, 559.0           # 绘图区上下（PY0 = 10000，PY1 = 0）
Y_MAX = 10000.0
SLOT = (PX1 - PX0) / len(LEVELS)  # 164
BAR_W = 52.0
WIDTH_U = BAR_W / SLOT

GRID_A = 0.225                    # #E6E6E6 @ 22.5% 叠在深蓝上的实测透明度
AXIS_A = 0.26
DASH_ON, DASH_OFF, DASH_H = 15, 5, 2
GRID_STEP = 2500.0                # 主网格线间距（Excel 里 valAx 的 majorUnit）

# 柱体渐变（每个柱各自从顶到底铺满）
# Excel 里写的是 #0070C0 → #00B0F0，但对参考图逐行线性拟合出来是 #006BBB → #00B9F5，
# 跨度略大（G 107→185 而不是 112→176）。按实测值走，柱体内部才没有残余色差。
GRAD_TOP, GRAD_BOTTOM = "#006BBB", "#00B9F5"

# 文字（按参考图的墨迹位置摆）
FONT_TITLE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_CAT = '"DengXian Light", "等线 Light", "DengXian", "等线", sans-serif'
FONT_NUM = '"Calibri", "等线", sans-serif'
FONT_NOTE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_UI = '"Microsoft YaHei", "微软雅黑", sans-serif'

FS_TITLE = 42.1
FS_CAT = 23.2
FS_NUM = 20.6
FS_NOTE = 18.6

C_TEXT = "#FFFFFF"
C_NOTE = "#D9D9D9"
C_NUM = "#F2F2F2"

# 切片器配色（逐像素取自 Excel 原版切片器）
C_PANEL_BD = "#4874CB"            # 面板边框 + 标题下方那条分隔线
C_BTN_BD = "#CCCCCC"              # 未选中按钮的边框
C_BTN_BG = "#FFFFFF"
C_BTN_SEL_BG = "#B5C6EA"          # 选中项：浅蓝紫底
C_BTN_SEL_BD = "#999999"          # 选中项的边框是更深的灰，不是蓝
C_BTN_TX = "#000000"

# 墨迹锚点（参考图实测；上下位置用「与参考图逐像素色差最小」扫描定出来的）
TITLE_INK = (55.0, 65.5)          # 渲染时 top = 65.5 - 2 = 63.5
NOTE_INK = (45.0, 632.0)
CAT_INK_Y = 579.5
NUM_GAP = 33.94                   # 数值标签盒顶 = 柱顶 - 33.94

# 切片器面板
# 比例取自 Excel 原版切片器（面板 1081x190）：标题栏 49、上留白 10、按钮高 42、行距 6、底部留白 41。
# 但按原比例缩到 312 宽只有 8px 字号、完全看不清，所以面板整体放大，只保留它的**相对节奏**。
PANEL_X, PANEL_Y, PANEL_W = 586.0, 54.0, 312.0
PANEL_HDR = 22.0                  # 标题栏高度
PANEL_PAD = 6.0                   # 标题栏下方 / 左右内边距
PANEL_PB = 15.0                   # 底部留白（Excel 里这块很宽，照抄）
ITEM_W, ITEM_H, ITEM_GAP = 72.0, 22.0, 4.0
PANEL_H = PANEL_HDR + PANEL_PAD + 2 * ITEM_H + ITEM_GAP + PANEL_PB

DIV_ID = "hrmfig"
CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


def bar_top(v):
    return PY1 - v * (PY1 - PY0) / Y_MAX


def cat_x(i):
    return PX0 + SLOT * (i + 0.5)


# ==========================================================================
# 3. 图
# ==========================================================================
def build_figure(values):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=list(range(len(LEVELS))), y=values, width=WIDTH_U,
        marker=dict(color="#0070C0", line=dict(width=0)),
        customdata=LEVELS,
        hovertemplate="<b>%{customdata}</b>　平均月收入 %{y:,.2f}<extra></extra>",
    ))
    ax_off = dict(showgrid=False, showline=False, zeroline=False,
                  showticklabels=False, ticks="", title=None)
    fig.update_layout(
        width=W, height=H,
        # 底色必须透明：网格线画在 SVG 下面的背板层里，SVG 一旦不透明就会把它盖掉
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=0, r=0, t=0, b=0, autoexpand=False),
        xaxis=dict(type="linear", range=[-0.5, len(LEVELS) - 0.5],
                   domain=[PX0 / W, PX1 / W], **ax_off),
        yaxis=dict(type="linear", range=[0, Y_MAX],
                   domain=[1 - PY1 / H, 1 - PY0 / H], **ax_off),
        showlegend=False, bargap=0.0, bargroupgap=0.0,
        hoverlabel=dict(bgcolor="#2A2F5A", bordercolor="#4A5080",
                        font=dict(color="#FFFFFF", size=13, family=FONT_UI)),
        dragmode=False,
    )
    return fig


# ==========================================================================
# 4. HTML / 交互
# ==========================================================================
CSS_FIX = ("<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
           ".plotly-graph-div{overflow:visible !important;background:" + BG + ";}"
           ".main-svg .barlayer path{shape-rendering:crispEdges;}</style>")

_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLOT_ID__');
  if (!gd) return;
  gd.style.position = 'relative';

  var DATA0 = __VALUES__;          // 首屏（= 源文件切片器状态）
  var SUMS  = __SUMS__;            // {部门: [5 个学历的收入合计]}
  var CNTS  = __CNTS__;            // {部门: [5 个学历的人数]}
  var DEPTS = __DEPTS__;
  var CATS  = __CATS__;
  var N     = CATS.length;
  var sel   = __SEL__;             // 默认选中的部门
  var STEP = 26, NFRAME = 10;

  function px(v) { return __PY1__ - v * (__PY1__ - __PY0__) / ymax; }
  function cx(i) { return __PX0__ + __SLOT__ * (i + 0.5); }

  /* ---------- 背板层：网格线 + 横轴 + 刻度线（必须压在柱子下面） ---------- */
  var sc = gd.querySelector('.svg-container') || gd;
  var back = document.createElement('div');
  back.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;' +
                       'pointer-events:none;z-index:0;';
  sc.insertBefore(back, sc.firstChild);

  function line(parent, css) {
    var d = document.createElement('div');
    d.style.cssText = 'position:absolute;' + css;
    parent.appendChild(d);
    return d;
  }

  /* 横轴与刻度线是固定的，先画 */
  line(back, 'left:__PX0__px;top:' + (__PY1__ - 1) + 'px;width:' + (__PX1__ - __PX0__ + 1) + 'px;' +
             'height:__DASH_H__px;background:rgba(230,230,230,__AXIS_A__);');
  for (var t = 0; t <= N; t++) {
    line(back, 'left:' + (__PX0__ + __SLOT__ * t - 1) + 'px;top:' + (__PY1__ + 1) + 'px;' +
               'width:2px;height:7px;background:rgba(230,230,230,__AXIS_A__);');
  }

  /* 网格线跟着纵轴量程走：Excel 的纵轴是自动的，数据超了就往上跳一档（2500 一档） */
  var ymax = __YMAX__, gridBox = document.createElement('div');
  gridBox.style.cssText = 'position:absolute;left:0;top:0;width:100%;height:100%;';
  back.insertBefore(gridBox, back.firstChild);

  function drawGrid(mx) {
    gridBox.innerHTML = '';
    for (var v = 0; v <= mx + 0.5; v += __YSTEP__) {
      var y = __PY1__ - v * (__PY1__ - __PY0__) / mx;
      line(gridBox, 'left:__PX0__px;top:' + (y - 1) + 'px;width:' + (__PX1__ - __PX0__) + 'px;' +
                    'height:__DASH_H__px;background:repeating-linear-gradient(to right,' +
                    'rgba(230,230,230,__GRID_A__) 0 __DASH_ON__px,' +
                    'rgba(230,230,230,0) __DASH_ON__px ' + (__DASH_ON__ + __DASH_OFF__) + 'px);');
    }
  }
  drawGrid(ymax);

  function axisFor(v) {
    var mx = Math.max.apply(null, v);
    return Math.max(__YMAX__, Math.ceil(mx / __YSTEP__) * __YSTEP__);
  }

  /* ---------- 覆盖层：标题 / 类目 / 数值 / 脚注 / 切片器 ---------- */
  var ov = document.createElement('div');
  ov.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;' +
                     'pointer-events:none;user-select:none;z-index:2;' +
                     'font-family:' + __FONT_UI__ + ';';
  gd.appendChild(ov);

  function mk(parent, css, txt) {
    var d = document.createElement('div');
    d.style.cssText = 'position:absolute;' + css;
    if (txt !== undefined) d.textContent = txt;
    parent.appendChild(d);
    return d;
  }

  mk(ov, 'left:__TIX__px;top:__TIY__px;font-family:' + __FONT_TITLE__ + ';' +
         'font-size:__FST__px;font-weight:700;line-height:1;color:__CT__;white-space:pre;', __TITLE__);
  /* 脚注：整句比参考宽约 3px，用 -0.06px 的字距把累积误差摊掉（CSS 值必须原样注入） */
  mk(ov, 'left:__NIX__px;top:__NIY__px;font-family:' + __FONT_NOTE__ + ';' +
         'font-size:__FSN__px;line-height:1;color:__CN__;white-space:pre;' +
         'letter-spacing:-0.06px;', __NOTE__);

  for (var i = 0; i < N; i++) {
    mk(ov, 'left:' + cx(i) + 'px;top:__CATY__px;transform:translateX(-50%);' +
           'font-family:' + __FONT_CAT__ + ';font-size:__FSC__px;line-height:1;' +
           'color:__CT__;white-space:pre;', CATS[i]);
  }

  var numLbl = [];
  for (var j = 0; j < N; j++) {
    numLbl[j] = mk(ov, 'left:' + cx(j) + 'px;top:' + (px(DATA0[j]) - __NUM_GAP__) + 'px;' +
                        'transform:translateX(-50%);font-family:' + __FONT_NUM__ + ';' +
                        'font-size:__FSNUM__px;line-height:1;color:__CNUM__;white-space:pre;',
                   DATA0[j].toFixed(2));
  }

  /* ---------- 部门切片器（照抄 Excel 原版切片器的配色与线框） ---------- */
  var panel = mk(ov, 'left:__PAX__px;top:__PAY__px;width:__PAW__px;height:__PAH__px;' +
                     'background:#FFFFFF;border:1px solid ' + __PANELBD__ + ';box-sizing:border-box;' +
                     'pointer-events:auto;');
  mk(panel, 'left:10px;top:3px;font-family:' + __FONT_UI__ + ';font-size:13.5px;' +
            'line-height:16px;color:' + __BTNTX__ + ';white-space:pre;', '部门');
  mk(panel, 'left:0;top:' + __PAHDR__ + 'px;width:100%;height:1px;background:' + __PANELBD__ + ';');

  /* 右上角两个小图标：左边「多选」（三个勾 + 三条横线），右边「清除筛选」（漏斗 + 红叉） */
  var ic = mk(panel, 'left:' + (__PAW__ - 46) + 'px;top:4px;width:40px;height:15px;');
  ic.innerHTML =
    '<svg width="40" height="15" viewBox="0 0 40 15" fill="none">' +
      '<g stroke="#404040" stroke-width="1.3" stroke-linecap="round">' +
        '<path d="M1 3.4l1.6 1.7 2.6-3.1"/><path d="M1 7.6l1.6 1.7 2.6-3.1"/>' +
        '<path d="M1 11.8l1.6 1.7 2.6-3.1"/>' +
        '<path d="M9.5 3.2h6"/><path d="M9.5 7.5h6"/><path d="M9.5 11.8h6"/>' +
      '</g>' +
      '<g stroke="#404040" stroke-width="1.3" stroke-linejoin="round">' +
        '<path d="M22 2.2h9.4l-3.9 4.8v6.1"/>' +
        '<path d="M22 2.2l3.9 4.8"/>' +
      '</g>' +
      '<g stroke="#E04A4A" stroke-width="1.4" stroke-linecap="round">' +
        '<path d="M29.6 8.6l3.6 3.6"/><path d="M33.2 8.6l-3.6 3.6"/>' +
      '</g>' +
    '</svg>';

  var items = [], on = {};
  for (var d = 0; d < DEPTS.length; d++) on[DEPTS[d]] = (sel.indexOf(DEPTS[d]) >= 0);

  for (var d2 = 0; d2 < DEPTS.length; d2++) {
    (function (idx, name) {
      var col = idx % 4, row = Math.floor(idx / 4);
      var b = mk(panel, 'left:' + (__PAD__ + col * (__IW__ + __IGAP__)) + 'px;' +
                        'top:' + (__PAHDR__ + __PAD__ + row * (__IH__ + __IGAP__)) + 'px;' +
                        'width:__IW__px;height:__IH__px;box-sizing:border-box;' +
                        'display:flex;align-items:center;justify-content:center;' +
                        'font-family:' + __FONT_UI__ + ';font-size:12.8px;line-height:1;' +
                        'cursor:pointer;user-select:none;border:1px solid ' + __BTNBD__ + ';' +
                        'background:' + __BTNBG__ + ';color:' + __BTNTX__ + ';', name);
      b.id = 'sl' + idx;
      items.push(b);
      b.addEventListener('click', function () { toggle(name); });
    })(d2, DEPTS[d2]);
  }

  function paint() {
    for (var k = 0; k < items.length; k++) {
      var nm = DEPTS[k], onIt = !!on[nm];
      items[k].style.background = onIt ? __SELBG__ : __BTNBG__;
      items[k].style.borderColor = onIt ? __SELBD__ : __BTNBD__;
    }
  }
  paint();

  /* ---------- 柱子换成竖向渐变（objectBoundingBox → 每根柱各自铺满） ---------- */
  function paintBars() {
    var svg = gd.querySelector('svg.main-svg');
    if (!svg) return;
    if (!svg.querySelector('#barGrad')) {
      var defs = document.createElementNS('http://www.w3.org/2000/svg', 'defs');
      defs.innerHTML = '<linearGradient id="barGrad" x1="0" y1="0" x2="0" y2="1">' +
                       '<stop offset="0%" stop-color="__G0__"/>' +
                       '<stop offset="100%" stop-color="__G1__"/></linearGradient>';
      svg.insertBefore(defs, svg.firstChild);
    }
    var pts = gd.querySelectorAll('.barlayer g.point > path');
    for (var i = 0; i < pts.length; i++) {
      // plotly 把填充写在 path 的内联 style 上，属性 fill 压不住，必须改 style
      pts[i].style.fill = 'url(#barGrad)';
      pts[i].style.stroke = 'none';
    }
  }
  paintBars();

  /* ---------- 切片器逻辑 ---------- */
  var busy = false, cur = DATA0;

  function pooled(list) {
    var out = [];
    for (var i = 0; i < N; i++) {
      var s = 0, n = 0;
      for (var k = 0; k < list.length; k++) { s += SUMS[list[k]][i]; n += CNTS[list[k]][i]; }
      out.push(n ? s / n : 0);
    }
    return out;
  }

  function apply(v) {
    Plotly.restyle(gd, {y: [v]}, [0]).then(paintBars);
    for (var i = 0; i < N; i++) {
      numLbl[i].style.top = (px(v[i]) - __NUM_GAP__) + 'px';
    }
  }

  function setText(v) {
    for (var i = 0; i < N; i++) numLbl[i].textContent = v[i].toFixed(2);
  }

  function goto(target) {
    if (busy) return;
    busy = true;
    var amax = axisFor(target);
    if (amax !== ymax) {                             // 数据超量程 → 纵轴自动跳一档
      ymax = amax;
      drawGrid(ymax);
      Plotly.relayout(gd, {'yaxis.range': [0, ymax]});
    }
    var from = cur, k = 1;
    (function tick() {
      var t = k / NFRAME;
      var e = t * t * (3 - 2 * t);                 // smoothstep
      var v = [];
      for (var i = 0; i < N; i++) v.push(from[i] + (target[i] - from[i]) * e);
      apply(v);
      if (k >= NFRAME) {
        setText(target); cur = target; busy = false; return;
      }
      k += 1;
      setTimeout(tick, STEP);
    })();
  }

  function toggle(name) {
    if (busy) return;
    on[name] = !on[name];
    paint();
    var picked = DEPTS.filter(function (d) { return on[d]; });
    goto(pooled(picked.length ? picked : DEPTS));   // 一个都不选 = 全部部门
  }

  /* 键盘 1..7 切换部门 */
  document.addEventListener('keydown', function (e) {
    var n = parseInt(e.key, 10);
    if (n >= 1 && n <= DEPTS.length) toggle(DEPTS[n - 1]);
  });
})();
"""


def build_html():
    fig = build_figure(averages(DEPT_DEFAULT))
    sums, cnts = {}, {}
    for d in DEPTS:
        s, n = agg([d])
        sums[d] = [round(v, 2) for v in s]
        cnts[d] = n

    script = (_POST_SCRIPT
              .replace("__PLOT_ID__", DIV_ID)
              .replace("__VALUES__", json.dumps([round(v, 4) for v in averages(DEPT_DEFAULT)]))
              .replace("__SUMS__", json.dumps(sums, ensure_ascii=False))
              .replace("__CNTS__", json.dumps(cnts, ensure_ascii=False))
              .replace("__DEPTS__", json.dumps(DEPTS, ensure_ascii=False))
              .replace("__CATS__", json.dumps(LEVELS, ensure_ascii=False))
              .replace("__SEL__", json.dumps(DEPT_DEFAULT, ensure_ascii=False))
              .replace("__W__", str(W)).replace("__H__", str(H))
              .replace("__PX0__", str(PX0)).replace("__PX1__", str(PX1))
              .replace("__PY0__", str(PY0)).replace("__PY1__", str(PY1))
              .replace("__YMAX__", str(Y_MAX)).replace("__SLOT__", str(SLOT))
              .replace("__YSTEP__", str(GRID_STEP))
              .replace("__DASH_ON__", str(DASH_ON)).replace("__DASH_OFF__", str(DASH_OFF))
              .replace("__DASH_H__", str(DASH_H))
              .replace("__GRID_A__", str(GRID_A)).replace("__AXIS_A__", str(AXIS_A))
              .replace("__FONT_UI__", json.dumps(FONT_UI))
              .replace("__FONT_TITLE__", json.dumps(FONT_TITLE))
              .replace("__FONT_CAT__", json.dumps(FONT_CAT))
              .replace("__FONT_NUM__", json.dumps(FONT_NUM))
              .replace("__FONT_NOTE__", json.dumps(FONT_NOTE))
              .replace("__TITLE__", json.dumps(TITLE, ensure_ascii=False))
              .replace("__NOTE__", json.dumps(NOTE, ensure_ascii=False))
              .replace("__TIX__", str(TITLE_INK[0])).replace("__TIY__", str(TITLE_INK[1] - 2))
              .replace("__NIX__", str(NOTE_INK[0])).replace("__NIY__", str(NOTE_INK[1] + 1))
              .replace("__CATY__", str(CAT_INK_Y))
              .replace("__NUM_GAP__", str(NUM_GAP))
              .replace("__FST__", str(FS_TITLE)).replace("__FSC__", str(FS_CAT))
              .replace("__FSNUM__", str(FS_NUM)).replace("__FSN__", str(FS_NOTE))
              .replace("__CT__", C_TEXT).replace("__CN__", C_NOTE).replace("__CNUM__", C_NUM)
              .replace("__PAX__", str(PANEL_X)).replace("__PAY__", str(PANEL_Y))
              .replace("__PAW__", str(PANEL_W)).replace("__PAH__", str(PANEL_H))
              .replace("__PAHDR__", str(PANEL_HDR)).replace("__PAD__", str(PANEL_PAD))
              .replace("__IW__", str(ITEM_W)).replace("__IH__", str(ITEM_H))
              .replace("__IGAP__", str(ITEM_GAP))
              .replace("__PANELBD__", json.dumps(C_PANEL_BD))
              .replace("__BTNBD__", json.dumps(C_BTN_BD))
              .replace("__BTNBG__", json.dumps(C_BTN_BG))
              .replace("__SELBG__", json.dumps(C_BTN_SEL_BG))
              .replace("__SELBD__", json.dumps(C_BTN_SEL_BD))
              .replace("__BTNTX__", json.dumps(C_BTN_TX))
              .replace("__G0__", GRAD_TOP).replace("__G1__", GRAD_BOTTOM))
    html = fig.to_html(include_plotlyjs=True, div_id=DIV_ID, config=CONFIG,
                       full_html=True, post_script=script)
    return html.replace("</head>", CSS_FIX + "</head>")


# ==========================================================================
# 5. 导出
# ==========================================================================
def save_html(path):
    html = build_html()
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    print("[ok] %s  (%.1f MB)" % (path, os.path.getsize(path) / 1048576))


def save_png(html_path, out_path):
    prof = tempfile.mkdtemp(prefix="chrome_")
    try:
        subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                        "--disable-lcd-text",   # 灰度抗锯齿，与参考图（原尺寸导出后缩放）更接近
                        "--force-device-scale-factor=1", "--user-data-dir=" + prof,
                        "--window-size=%d,%d" % (W, H),
                        "--run-all-compositor-stages-before-draw",
                        "--virtual-time-budget=4000",
                        "--screenshot=" + out_path,
                        "file:///" + quote(html_path.replace("\\", "/"))],
                       capture_output=True, timeout=180)
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    print("[ok] %s" % out_path)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    html_path = os.path.join(OUT_DIR, "5 人力资源明细_plotly.html")
    save_html(html_path)
    save_png(html_path, os.path.join(OUT_DIR, "5 人力资源明细_plotly.png"))


if __name__ == "__main__":
    main()
