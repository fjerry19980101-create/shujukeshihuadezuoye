# -*- coding: utf-8 -*-
"""
「7 动态滑珠图」plotly 版：横向 100% 堆叠条 + 一个可滑动的珠子 + 右侧百分比，
右上角是「月份 / 区域」单选项 + 一个下拉框，切换时整排珠子一起滑过去。

参考基准：948 x 685（本机 Excel 导出图表 → 等比缩放，原尺寸 924x667）
数据源  ：第三章 动态图表\\7 动态滑珠图.xlsx
          · 源数据   B2:H8   = 6 个区域 × 6 个月的达成率（0~1）
          · 图表数据 B15:F20 = 按 K2 / K3 用 VLOOKUP / HLOOKUP 取出来的当期数据
          · K2 = 两个单选项（1=月份 2=区域），K3 = 下拉框（1~6）
          · F15:F20 = 辅助 1/3/5/7/9/11，喂给散点系列定位六行
输出    ：...\\作业\\图\\第三章\\7 动态滑珠图_plotly.html / .png

一、这张图的构造（拆自 chart1.xml）
    · 两个图组叠在一起：
      ① barChart   barDir=bar（横向）、grouping=stacked、gapWidth=200、overlap=100
         系列1 完成率   #0070C0
         系列2 未完成率 bg1×95%@30%  → 压在深蓝底上就是 #5B5E78
      ② scatterChart  只有标记：circle、size=15pt、填充 #0070C0、描边 bg1×95% 1pt
         x = 完成率（和值轴同一根轴），y = 辅助 1/3/5/7/9/11
    · 散点的 y 轴是一根隐藏的次值轴，量程 0~12：六行落在 1/3/5/7/9/11 上，
      换算成占满比例的 (i-0.5)/6 —— 与柱子的类目中心**完全重合**，
      这就是「辅助 = 1,3,5,7,9,11」这串怪数据的用途（本脚本直接用类目轴落点，不需要它）。
    · 值轴 0~1、max=1 写死；两根轴都没有轴线、没有刻度、没有网格 → 画面上只剩色块
    · 数值标签：dLblPos=r（珠子右侧）、9pt、#F2F2F2，内容是单元格文本（CELLRANGE 字段），
      所以显示的是「32%」而不是「0.32」
    · 图表区 #1A1E43；标题与脚注是 userShapes（图表里的两个矩形，不是工作表文本框）

二、像素几何（画布 948x685）
    · 绘图区 x 112.3..887.9（宽 775.6），y 181.5..616.0（高 434.5）
    · 六行中心 217.5 / 289.9 / 362.3 / 434.7 / 507.1 / 579.5（行距 72.4）
    · 柱高 24.1（= 行距 ÷ 3，gapWidth=200 的换算），圆角 0
    · 珠子直径 40（含 2.66px 描边），圆心正好落在「完成率」的分界处

三、文字（全部用 DOM 画，位置按参考图墨迹实测）
    · 标题  微软雅黑 Bold 18pt #FFFFFF，墨迹顶 46、左 56
    · 类目  宋体 9pt #F2F2F2，右对齐到 x=88.5
    · 数值  Calibri 9pt #F2F2F2，墨迹左沿 = 珠心 + 38.5
    · 脚注  微软雅黑 8pt #D9D9D9，墨迹从 (49, 640) 起

四、交互（照抄原文件的控件语义）
    · 单选项「月份」（默认，K2=1）：类目 = 6 个区域，下拉框选 1~6 月 → 六根条都是那个月
    · 单选项「区域」（K2=2）：类目 = 1~6 月，下拉框选区域 → 六根条都是那个区域
    · 下拉框的候选项本身就是 $B$15:$B$20，随 K2 变：所以切到「区域」后它列的是区域名
    · 切换时柱宽 / 珠子 / 百分比同步补间（12 帧），最后一帧才换文字
"""
import json
import os
import shutil
import subprocess
import tempfile
from decimal import Decimal, ROUND_HALF_UP
from urllib.parse import quote

import openpyxl
import plotly.graph_objects as go

# ==========================================================================
# 0. 路径
# ==========================================================================
BASE = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\作业")
XLSX = os.path.join(BASE, "第三章 动态图表", "7 动态滑珠图.xlsx")
OUT_DIR = os.path.join(BASE, "图", "第三章")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

W, H = 948, 685
DIV_ID = "slipfig"

# ==========================================================================
# 1. 数据
# ==========================================================================
BG = "#1A1E43"
C_BAR = "#0070C0"          # 完成率
C_REST = "#5B5E78"         # 未完成率（bg1×95% @30% 压在 #1A1E43 上的合成色）
C_RING = "#F2F2F2"         # 珠子描边
C_TITLE = "#FFFFFF"
C_TEXT = "#F2F2F2"
C_NOTE = "#D9D9D9"

MODE_MONTH, MODE_AREA = 1, 2
DEFAULT_MODE = MODE_MONTH
DEFAULT_IDX = 1            # K3 = 1


def load_table(path):
    """读 B2:H8 → (区域列表, 月份列表, {区域: [6 个月的达成率]})。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    months = [ws.cell(row=2, column=c).value for c in range(3, 9)]
    regions, table = [], {}
    for r in range(3, 9):
        name = ws.cell(row=r, column=2).value
        regions.append(name)
        table[name] = [float(ws.cell(row=r, column=c).value) for c in range(3, 9)]
    wb.close()
    return regions, [str(m) for m in months], table


REGIONS, MONTHS, TABLE = load_table(XLSX)


def pct(v):
    """0.536 → '54%'。Excel 的 0% 格式是四舍五入（half-up），不是银行的 banker's rounding。"""
    q = Decimal(str(v * 100)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
    return "%d%%" % int(q)


def state(mode, idx):
    """复刻 B15:F20 那两个 IF：返回 (类目列表, 值列表)，类目自下而上。"""
    if mode == MODE_MONTH:
        cats = list(REGIONS)
        vals = [TABLE[r][idx - 1] for r in REGIONS]
    else:
        cats = list(MONTHS)
        vals = [TABLE[REGIONS[idx - 1]][j] for j in range(len(MONTHS))]
    return cats, vals


# ==========================================================================
# 2. 几何与字体
# ==========================================================================
PLOT_X0, PLOT_X1 = 112.8, 888.4
PLOT_Y0, PLOT_Y1 = 181.5, 616.0
PLOT_W = PLOT_X1 - PLOT_X0                 # 775.6
PLOT_H = PLOT_Y1 - PLOT_Y0                 # 434.5

N_ROW = 6
ROW_C0 = PLOT_Y0 + PLOT_H / N_ROW / 2      # 217.7 第一行（最上面那行）的中心
ROW_PITCH = PLOT_H / N_ROW                 # 72.4
BAR_W = 1.0 / 3.0                          # gapWidth=200 → 柱高 = 行距 ÷ 3

PX_PER_PT = W / 355.95                     # 图表宽 355.95pt → 948px
M_SIZE = 41.0                              # 珠子外径（含描边）
M_LINE = 2.66                              # 1pt 描边

FONT_TITLE = "Microsoft YaHei"
FONT_CAT = "SimSun"
FONT_NUM = "Calibri"
FONT_NOTE = "Microsoft YaHei"
FONT_UI = "Microsoft YaHei UI"

FS_TITLE = 48.2
FS_CAT = 24.2
FS_NUM = 24.9
FS_NOTE = 21.3
FS_UI = 22.0

TITLE_X, TITLE_Y = 54.0, 43.0              # 墨迹左上
CAT_RX = 91.0                              # 类目右对齐位置
CAT_DY = -1.3                              # 类目墨迹中心相对行中心的偏移
NOTE_X, NOTE_Y = 48.5, 639.0
LBL_GAP = 37.4                             # 数值墨迹左沿 = 珠心 + LBL_GAP

TITLE = "2022年上半年区域销量目标达成率情况"
NOTE = "*注：数据来源公司销售系统,日期截至2022.06.30"

# 窗体控件（照抄原文件的窗体控件坐标，相对图表左上角；原值单位 pt，这里换成 px）
def _ctrl(x_pt, y_pt, w_pt, h_pt):
    k = W / 355.95
    return (round(x_pt * k, 1), round(y_pt * k, 1),
            round(w_pt * k, 1), round(h_pt * k, 1))


CTRL = {
    "area":  _ctrl(173.2, 48.6, 44.6, 17.6),   # 区域（未选中）
    "month": _ctrl(230.6, 49.4, 46.1, 16.9),   # 月份（默认选中）
    "drop":  _ctrl(288.7, 49.0, 43.9, 17.2),
}

# ==========================================================================
# 3. 图
# ==========================================================================
def build_figure(mode, idx):
    cats, vals = state(mode, idx)
    rest = [1.0 - v for v in vals]
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=cats, x=vals, orientation="h", width=BAR_W,
        marker=dict(color=C_BAR, line=dict(width=0)),
        hoverinfo="skip", showlegend=False,
    ))
    fig.add_trace(go.Bar(
        y=cats, x=rest, orientation="h", width=BAR_W,
        marker=dict(color=C_REST, line=dict(width=0)),
        hoverinfo="skip", showlegend=False,
    ))
    fig.add_trace(go.Scatter(
        y=cats, x=vals, mode="markers",
        marker=dict(symbol="circle", size=M_SIZE - M_LINE,
                    color=C_BAR,
                    line=dict(color=C_RING, width=M_LINE)),
        hoverinfo="skip", showlegend=False,
    ))
    fig.update_layout(
        width=W, height=H,
        paper_bgcolor=BG, plot_bgcolor="rgba(0,0,0,0)",
        margin=dict(l=PLOT_X0, r=W - PLOT_X1, t=PLOT_Y0, b=H - PLOT_Y1,
                    autoexpand=False),
        barmode="stack",
        bargap=0.0,
        showlegend=False,
        dragmode=False,
        hovermode=False,
        transition=dict(duration=0),
        font=dict(family=FONT_NUM),
    )
    fig.update_xaxes(range=[0, 1], visible=False, fixedrange=True)
    # 用 trace 顺序排类目（而不是写死 categoryarray）：
    # 切模式时 restyle(y=...) 会直接换成另一组类目，categoryarray 会拦住新类目导致整排消失
    fig.update_yaxes(categoryorder="trace",
                     visible=False, fixedrange=True,
                     range=[-0.5, N_ROW - 0.5])
    return fig


# ==========================================================================
# 4. HTML / 交互
# ==========================================================================
CSS_FIX = ("<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
           ".plotly-graph-div{overflow:visible !important;}"
           ".main-svg{background:transparent !important;}</style>")

_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLID__');
  if (!gd) return;
  gd.style.position = 'relative';

  var N = __N__;
  var ROW_C0 = __ROWC0__, PITCH = __PITCH__;
  var X0 = __X0__, PW = __PW__;
  var TABLE = __TABLE__, REGIONS = __REGIONS__, MONTHS = __MONTHS__;
  var LI = __LI__, LBL_GAP = __LBLGAP__;
  var mode = __MODE__, idx = __IDX__;

  /* ---------- 覆盖层 ---------- */
  var ov = document.createElement('div');
  ov.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;' +
                     'pointer-events:none;user-select:none;z-index:2;';
  gd.appendChild(ov);

  function mk(css, txt) {
    var d = document.createElement('div');
    d.style.cssText = 'position:absolute;' + css;
    if (txt !== undefined) d.textContent = txt;
    ov.appendChild(d);
    return d;
  }

  /* 标题 */
  mk('left:__TX__px;top:__TY__px;font-family:' + __FT__ + ';font-size:__FST__px;' +
     'font-weight:700;line-height:1;color:__CT__;white-space:pre;', __TITLE__);

  /* 脚注 */
  mk('left:__NX__px;top:__NY__px;font-family:' + __FN__ + ';font-size:__FSN__px;' +
     'line-height:1;color:__CN__;white-space:pre;', __NOTE__);

  /* 类目：右对齐到 x=CAT_RX，第 0 行在最上面 */
  var catL = [];
  for (var i = 0; i < N; i++) {
    catL[i] = mk('left:0;width:__CATRX__px;text-align:right;top:0;' +
      'font-family:' + __FC__ + ';font-size:__FSC__px;line-height:1;' +
      'color:__CC__;white-space:pre;');
  }

  /* 数值标签：珠心 + LBL_GAP，左对齐 */
  var numL = [];
  for (var j = 0; j < N; j++) {
    numL[j] = mk('left:0;top:0;font-family:' + __FNUM__ + ';font-size:__FSNUM__px;' +
      'line-height:1;color:__CNUM__;white-space:pre;');
  }

  function pctf(v) { var r = Math.floor(v * 100 + 0.5); return r + '%'; }

  function paintCats(cats) {
    for (var i = 0; i < N; i++) {
      catL[i].textContent = cats[N - 1 - i];
      catL[i].style.top = (ROW_C0 + PITCH * i - __FSC__ * 0.5 + __CATDY__) + 'px';
    }
  }

  function paintNums(vals) {
    for (var j = 0; j < N; j++) {
      var cy = ROW_C0 + PITCH * (N - 1 - j);
      numL[j].style.left = (X0 + vals[j] * PW + LBL_GAP) + 'px';
      numL[j].style.top = (cy - __FSNUM__ * 0.5 + __NUMDY__) + 'px';
    }
  }
  function setNumText(vals) {
    for (var j = 0; j < N; j++) numL[j].textContent = pctf(vals[j]);
  }

  function series(mode, idx) {
    var cats, vals = [];
    if (mode === 1) {
      cats = REGIONS.slice();
      for (var i = 0; i < N; i++) vals.push(TABLE[REGIONS[i]][idx - 1]);
    } else {
      cats = MONTHS.slice();
      for (var j = 0; j < N; j++) vals.push(TABLE[REGIONS[idx - 1]][j]);
    }
    return { cats: cats, vals: vals };
  }

  var cur = series(mode, idx);
  paintCats(cur.cats);
  paintNums(cur.vals);
  setNumText(cur.vals);

  function apply(vals) {
    var rest = vals.map(function (v) { return 1 - v; });
    Plotly.restyle(gd, { 'x': [vals, rest, vals] }, [0, 1, 2]);
    paintNums(vals);
  }

  /* ---------- 单选项 + 下拉框（照抄窗体控件位置） ---------- */
  var ui = document.createElement('div');
  ui.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;z-index:3;';
  gd.appendChild(ui);

  function ctrlBox(key, label, on) {
    var b = __CTRL__[key];
    var el = document.createElement('div');
    el.style.cssText = 'position:absolute;left:' + b[0] + 'px;top:' + b[1] + 'px;' +
      'width:' + b[2] + 'px;height:' + b[3] + 'px;box-sizing:border-box;' +
      'display:flex;align-items:center;padding-left:4px;' +
      'background:#F2F2F2;border:1px solid #A6A6A6;' +
      'font-family:' + __FUI__ + ';font-size:__FSUI__px;color:#000;' +
      'cursor:pointer;user-select:none;';
    var dot = document.createElement('span');
    dot.style.cssText = 'display:inline-block;width:' + (b[3] * 0.36) + 'px;height:' +
      (b[3] * 0.36) + 'px;border-radius:50%;border:1px solid #6E6E6E;' +
      'background:#FFF;margin-right:4px;flex:0 0 auto;position:relative;';
    var inner = document.createElement('span');
    inner.style.cssText = 'position:absolute;left:22%;top:22%;width:56%;height:56%;' +
      'border-radius:50%;background:#1A1A1A;display:' + (on ? 'block' : 'none') + ';';
    dot.appendChild(inner);
    el.appendChild(dot);
    var tx = document.createElement('span');
    tx.textContent = label;
    el.appendChild(tx);
    el.dataset.inner = '1';
    el._inner = inner;
    ui.appendChild(el);
    return el;
  }

  var rbArea = ctrlBox('area', '区域', mode === 2);
  var rbMonth = ctrlBox('month', '月份', mode === 1);

  /* 下拉框 */
  var db = __CTRL__['drop'];
  var combo = document.createElement('div');
  combo.style.cssText = 'position:absolute;left:' + db[0] + 'px;top:' + db[1] + 'px;' +
    'width:' + db[2] + 'px;height:' + db[3] + 'px;box-sizing:border-box;' +
    'display:flex;align-items:stretch;' +
    'background:#FFF;border:1px solid #A6A6A6;' +
    'font-family:' + __FUI__ + ';font-size:__FSUI__px;color:#000;' +
    'cursor:pointer;user-select:none;';
  var comboText = document.createElement('span');
  comboText.style.cssText = 'display:flex;align-items:center;padding-left:5px;flex:1 1 auto;';
  combo.appendChild(comboText);
  var arrowBox = document.createElement('span');
  arrowBox.style.cssText = 'flex:0 0 auto;width:' + Math.round(db[3] * 0.62) + 'px;' +
    'display:flex;align-items:center;justify-content:center;' +
    'border-left:1px solid #C8C8C8;background:#F0F0F0;';
  var arrow = document.createElement('span');
  arrow.textContent = '\u25BC';
  arrow.style.cssText = 'font-size:' + (__FSUI__ * 0.55) + 'px;color:#3A3A3A;line-height:1;';
  arrowBox.appendChild(arrow);
  combo.appendChild(arrowBox);
  ui.appendChild(combo);

  var list = document.createElement('div');
  list.style.cssText = 'position:absolute;left:' + db[0] + 'px;top:' + (db[1] + db[3]) +
    'px;width:' + db[2] + 'px;background:#FFF;border:1px solid #7F7F7F;display:none;' +
    'box-shadow:2px 2px 4px rgba(0,0,0,0.35);z-index:10;';
  ui.appendChild(list);

  function items() { return mode === 1 ? MONTHS : REGIONS; }

  function paintCombo() {
    comboText.textContent = items()[idx - 1];
    list.innerHTML = '';
    items().forEach(function (name, k) {
      var it = document.createElement('div');
      it.textContent = name;
      it.style.cssText = 'padding:2px 5px;font-family:' + __FUI__ + ';font-size:' +
        (__FSUI__ * 0.8) + 'px;color:#000;cursor:pointer;';
      it.style.background = (k === idx - 1) ? '#CCE4F7' : '#FFF';
      it.addEventListener('mouseenter', function () { it.style.background = '#E5F1FB'; });
      it.addEventListener('mouseleave', function () {
        it.style.background = (k === idx - 1) ? '#CCE4F7' : '#FFF'; });
      it.addEventListener('click', function (ev) {
        ev.stopPropagation();
        list.style.display = 'none';
        var n = k + 1;
        if (n === idx) return;
        idx = n; paintCombo(); transitionTo();
      });
      list.appendChild(it);
    });
  }
  combo.addEventListener('click', function (ev) {
    ev.stopPropagation();
    list.style.display = (list.style.display === 'none') ? 'block' : 'none';
  });
  document.addEventListener('click', function () { list.style.display = 'none'; });

  /* ---------- 状态切换 ---------- */
  var busy = false, NFR = 12, MS = 22, vNow = cur.vals.slice();

  function transitionTo() {
    if (busy) return;
    busy = true;
    var to = series(mode, idx);
    var from = vNow.slice(), goal = to.vals.slice();
    if (to.cats.join('|') !== cur.cats.join('|')) {
      cur.cats = to.cats;
      paintCats(cur.cats);
      Plotly.restyle(gd, { 'y': [to.cats, to.cats, to.cats] }, [0, 1, 2]);
    }
    var k = 1;
    (function tick() {
      var t = k / NFR, e = t * t * (3 - 2 * t), v = [];
      for (var i = 0; i < N; i++) v.push(from[i] + (goal[i] - from[i]) * e);
      apply(v);
      if (k >= NFR) { vNow = goal.slice(); setNumText(goal); busy = false; return; }
      k += 1;
      setTimeout(tick, MS);
    })();
  }

  function setMode(m) {
    if (m === mode) return;
    mode = m;
    rbArea._inner.style.display = (m === 2) ? 'block' : 'none';
    rbMonth._inner.style.display = (m === 1) ? 'block' : 'none';
    paintCombo();
    transitionTo();
  }

  rbArea.addEventListener('click', function () { setMode(2); });
  rbMonth.addEventListener('click', function () { setMode(1); });

  paintCombo();

  document.addEventListener('keydown', function (e) {
    if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
      var n = idx + (e.key === 'ArrowUp' ? -1 : 1);
      if (n < 1 || n > N) return;
      idx = n; paintCombo(); transitionTo();
    }
    if (e.key === '1') setMode(1);
    if (e.key === '2') setMode(2);
  });
})();
"""


def build_html():
    script = (_POST_SCRIPT
              .replace("__PLID__", DIV_ID)
              .replace("__W__", str(W)).replace("__H__", str(H))
              .replace("__N__", str(N_ROW))
              .replace("__ROWC0__", str(ROW_C0)).replace("__PITCH__", str(ROW_PITCH))
              .replace("__X0__", str(PLOT_X0)).replace("__PW__", str(PLOT_W))
              .replace("__LI__", str(PLOT_X0))
              .replace("__LBLGAP__", str(LBL_GAP))
              .replace("__CATDY__", str(CAT_DY)).replace("__NUMDY__", "0")
              .replace("__TABLE__", json.dumps(TABLE, ensure_ascii=False))
              .replace("__REGIONS__", json.dumps(REGIONS, ensure_ascii=False))
              .replace("__MONTHS__", json.dumps(MONTHS, ensure_ascii=False))
              .replace("__MODE__", str(DEFAULT_MODE)).replace("__IDX__", str(DEFAULT_IDX))
              .replace("__CTRL__", json.dumps(CTRL))
              .replace("__TX__", str(TITLE_X)).replace("__TY__", str(TITLE_Y))
              .replace("__CATRX__", str(CAT_RX))
              .replace("__NX__", str(NOTE_X)).replace("__NY__", str(NOTE_Y))
              .replace("__TITLE__", json.dumps(TITLE, ensure_ascii=False))
              .replace("__NOTE__", json.dumps(NOTE, ensure_ascii=False))
              .replace("__FT__", json.dumps(FONT_TITLE)).replace("__FC__", json.dumps(FONT_CAT))
              .replace("__FNUM__", json.dumps(FONT_NUM)).replace("__FN__", json.dumps(FONT_NOTE))
              .replace("__FUI__", json.dumps(FONT_UI))
              .replace("__FST__", str(FS_TITLE)).replace("__FSC__", str(FS_CAT))
              .replace("__FSNUM__", str(FS_NUM)).replace("__FSN__", str(FS_NOTE))
              .replace("__FSUI__", str(FS_UI))
              .replace("__CT__", C_TITLE).replace("__CC__", C_TEXT)
              .replace("__CN__", C_NOTE).replace("__CNUM__", C_TEXT))
    fig = build_figure(DEFAULT_MODE, DEFAULT_IDX)
    html = fig.to_html(include_plotlyjs=True, div_id=DIV_ID,
                       config={"displayModeBar": False, "staticPlot": False,
                               "scrollZoom": False, "doubleClick": False},
                       full_html=True, post_script=script)
    return html.replace("</head>", CSS_FIX + "</head>")


# ==========================================================================
# 5. 导出
# ==========================================================================
def save_html(path):
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(build_html())
    print("[ok] %s  (%.1f MB)" % (path, os.path.getsize(path) / 1048576))


def save_png(html_path, out_path):
    prof = tempfile.mkdtemp(prefix="chrome_")
    try:
        subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
                        "--disable-lcd-text",
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
    html_path = os.path.join(OUT_DIR, "7 动态滑珠图_plotly.html")
    save_html(html_path)
    save_png(html_path, os.path.join(OUT_DIR, "7 动态滑珠图_plotly.png"))
    cats, vals = state(DEFAULT_MODE, DEFAULT_IDX)
    for c, v in zip(cats, vals):
        print("   %-4s %s" % (c, pct(v)))


if __name__ == "__main__":
    main()
