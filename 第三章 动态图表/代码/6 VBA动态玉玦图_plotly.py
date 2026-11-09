# -*- coding: utf-8 -*-
"""
「6 VBA动态玉玦图」plotly 版：四个同心环，每环是一段弧（值 ÷ 50% 的圆心角），
右上角两个按钮切换「近7天 / 近30天」，标题前缀随周期变化。

参考基准：948 x 692（本机 Excel 导出图表 → 等比缩放）
数据源  ：第三章 动态图表\\6 VBA动态玉玦图.xlsx
          · 源数据   B3:F5   = 周期 × 4 个渠道的百分比
          · 图表数据 B10:F12 = VLOOKUP(H3) 取当期值，占位 = 50% - 值
          · H3       = 控件输入（窗体按钮写入），默认「近30天」
          · 按钮 1/2 的宏：把 "近7天"/"近30天" 写进单元格（原文件里宏录到了 L3）
输出    ：...\\作业\\图\\第三章\\6 VBA动态玉玦图_plotly.html / .png

一、这张图的结构（拆自 chart1.xml）
    · doughnutChart，4 个系列 = 4 个同心环，最先的系列在最内圈
        内→外：个人主页 #00B0F0 → 搜索 #11A7AD → 关注页面 #FFC000 → 首页推荐 #E74E69
    · 每环 2 个数据点：[当期值, 占位]，占位 = 50% - 值，所以**两段合计恒等于 50%**
      → 值那一段占整圈的比例 = 值 / 0.5 = 值的 2 倍（0.09 → 64.8°）
    · holeSize = 35、firstSliceAng = 0（12 点起，顺时针）
    · 环边框 #1A1E43 / 1.5pt：相邻环各画一半，合成中间那道约 4.5px 的深色分隔带
    · 占位段 noFill，只剩同样颜色的边框 → 在深蓝底上完全看不见，弧就是「戛然而止」
    · 数值标签 bg1×95% = #F2F2F2、9pt、0%；只显示值那一点（占位点被 dLbl delete=1）
    · 图表区 #1A1E43，边框同色（等于没有）

二、像素几何（画布 948x692，1pt = 2.675px）
    · 圆心 (455, 380.5)，外半径 225，内孔半径 79，环厚 36.5，分隔带宽 4.5
    · 四个数值标签落在环中线上（r = 97.1 / 134.3 / 171.6 / 208.1），角度 = 值的 360 倍
      （弧是 0..值×720°，中点就是值×360°）

三、文字（全部用 DOM 画，位置按参考图墨迹实测）
    · 标题  微软雅黑 Bold 18pt 白：**动态前缀「近30天/近7天」+ 固定「流量来源分布」**
            前缀右对齐到 x=231.5，正文从 x=239 起，两者之间有约 7.5px 的自然间隙
    · 类目  宋体 10pt #AAABB7，右对齐到 x=445，行距 40px（首行墨迹顶 158）
    · 数值  Calibri 9pt #F2F2F2
    · 脚注  微软雅黑 8pt #D9D9D9，墨迹从 (60, 641) 起

四、交互
    右上角两个按钮（位置照抄原文件的窗体按钮：相对图表 (238.1,49.9) / (291.4,50.2) pt）。
    切换时四个环按 smoothstep 补间扫过去，数值标签跟着弧中点走、最后一帧才换数字，
    标题前缀同步改。键盘 ← / → 也能切。
"""
import json
import os
import shutil
import subprocess
import tempfile
from urllib.parse import quote

import openpyxl
import plotly.graph_objects as go

# ==========================================================================
# 0. 路径
# ==========================================================================
BASE = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\作业")
XLSX = os.path.join(BASE, "第三章 动态图表", "6 VBA动态玉玦图.xlsx")
OUT_DIR = os.path.join(BASE, "图", "第三章")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# ==========================================================================
# 1. 数据
# ==========================================================================
# 系列顺序 = chart1.xml 里 c:ser 的 idx 顺序 = 内 → 外
SERIES = [
    ("个人主页", "#00B0F0"),
    ("搜索", "#11A7AD"),
    ("关注页面", "#FFC000"),
    ("首页推荐", "#E74E69"),
]
PERIODS = ["近7天", "近30天"]        # B4 / B5 的顺序
DEFAULT_PERIOD = "近30天"            # H3 的初始值

TITLE_BODY = "流量来源分布"           # 标题的固定部分（前缀是周期）
NOTE = "*注：数据来源公司网站"


def load_data(path):
    """读 B3:F5 → {周期: [4 个值]}，顺序与 SERIES 一致。"""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    hdr = [ws.cell(row=3, column=c).value for c in range(2, 7)]     # B3..F3
    if hdr[0] != "周期":
        raise ValueError("表头不对：B3 应为「周期」，实际 %r" % (hdr[0],))
    names = [str(v) for v in hdr[1:]]
    if names != [s[0] for s in SERIES]:
        raise ValueError("渠道顺序与图表不一致：%r" % (names,))

    out = {}
    for r in range(4, 6):
        per = str(ws.cell(row=r, column=2).value)
        out[per] = [float(ws.cell(row=r, column=c).value) for c in range(3, 7)]
    wb.close()
    return out


VALUES = load_data(XLSX)

# ==========================================================================
# 2. 画布与几何
# ==========================================================================
W, H = 948, 692
BG = "#1A1E43"

CX, CY = 456.0, 380.5
# 五条半径（内孔 → 外沿）直接按参考图的分隔带中心量出来，
# Excel 的环其实不是等厚的（36~37.5 交替），照抄比强行均分更准
RADII = [80.2, 116.0, 152.5, 189.5, 223.4]
R_IN, R_OUT = RADII[0], RADII[-1]
N_RING = len(SERIES)
R_STEP = (R_OUT - R_IN) / N_RING
# plotly 的饼图圆心比 domain 算出来的位置固定偏一点点，实测后补回来
PIE_DX, PIE_DY = 0.0, 0.0
RING_STROKE = 4.5                     # 相邻环边框叠出来的分隔带宽度
HALF_MAX = 0.5                        # 占位 = 50% - 值，所以一段弧最多占半圈

# 每环数值标签的落点半径（参考图实测，略大于环中线）
LBL_R = [97.1, 134.3, 171.6, 208.1]

C_NUM = "#F2F2F2"
C_CAT = "#AAABB7"
C_NOTE = "#D9D9D9"
C_TITLE = "#FFFFFF"

FONT_TITLE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_CAT = '"SimSun", "宋体", serif'
FONT_NUM = '"Calibri", sans-serif'
FONT_NOTE = '"Microsoft YaHei", "微软雅黑", sans-serif'
FONT_UI = '"Microsoft YaHei", "微软雅黑", sans-serif'

FS_TITLE = 47.95
FS_CAT = 28.0
FS_NUM = 24.07
FS_NOTE = 21.1

# 文字锚点（参考图墨迹实测）
TITLE_BODY_X = 238.0
TITLE_PREFIX_RX = 231.5               # 前缀右边缘
TITLE_INK_Y = 53.0
CAT_RX = 449.0                        # 类目右对齐
CAT_Y0, CAT_PITCH = 155.5, 40.0
NOTE_X, NOTE_Y = 59.0, 641.0

# 窗体按钮（相对图表 pt × 1pt=2.675px）
BTN = [(238.1, 49.9, "近7天"), (291.4, 50.2, "近30天")]
PX_PER_PT = W / 354.4

DIV_ID = "juefig"
CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


def apply_font(fig, family, size, color):
    fig.update_layout(font=dict(family=family, size=size, color=color))


# ==========================================================================
# 3. 图
# ==========================================================================
def ring_domain(r_out):
    """把半径 r_out 的圆映射成 pie 的 domain（画布 948x692，margin=0）。"""
    px, py = CX + PIE_DX, CY + PIE_DY
    return ([(px - r_out) / W, (px + r_out) / W],
            [1 - (py + r_out) / H, 1 - (py - r_out) / H])


def build_figure(period):
    vals = VALUES[period]
    fig = go.Figure()
    for k, (name, color) in enumerate(SERIES):
        r_in, r_out = RADII[k], RADII[k + 1]
        dx, dy = ring_domain(r_out)
        fig.add_trace(go.Pie(
            values=[vals[k], HALF_MAX - vals[k]],
            # 占位段彻底透明；label 只给值那一段
            labels=[name, "占位"],
            marker=dict(colors=[color, "rgba(0,0,0,0)"],
                        line=dict(color=BG, width=0)),
            domain=dict(x=dx, y=dy),
            hole=r_in / r_out,
            rotation=0,                    # 12 点起
            direction="clockwise",         # 与 Excel firstSliceAng=0 + 顺时针一致
            sort=False,
            textinfo="none",
            hoverinfo="skip",
            showlegend=False,
        ))
    fig.update_layout(
        width=W, height=H,
        paper_bgcolor=BG, plot_bgcolor=BG,
        margin=dict(l=0, r=0, t=0, b=0, autoexpand=False),
        showlegend=False,
        dragmode=False,
        hovermode=False,
        transition=dict(duration=0),
    )
    return fig


# ==========================================================================
# 4. HTML / 交互
# ==========================================================================
CSS_FIX = ("<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
           ".plotly-graph-div{overflow:visible !important;}</style>")

_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLID__');
  if (!gd) return;
  gd.style.position = 'relative';

  var CX = __CX__, CY = __CY__, R_IN = __RIN__, R_OUT = __ROUT__;
  var RADII = __RADII__;
  var N = __SERIES__.length;
  var VALS = __VALUES__;          // {周期: [4 个值]}
  var PERIODS = __PERIODS__;
  var cur = __CUR__;              // 当前周期名
  var LBL_R = __LBLR__;

  function pt(r, a) {             // a: 从 12 点起顺时针的角度（度）
    var t = a * Math.PI / 180;
    return [CX + r * Math.sin(t), CY - r * Math.cos(t)];
  }

  /* ---------- 环之间的深色分隔带（叠在 plotly 之上，免得被外环盖掉） ---------- */
  var svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
  svg.setAttribute('width', __W__); svg.setAttribute('height', __H__);
  svg.setAttribute('viewBox', '0 0 __W__ __H__');
  svg.style.cssText = 'position:absolute;left:0;top:0;pointer-events:none;z-index:1;';
  for (var k = 1; k < N; k++) {
    var r = RADII[k];
    var c = document.createElementNS('http://www.w3.org/2000/svg', 'circle');
    c.setAttribute('cx', CX); c.setAttribute('cy', CY); c.setAttribute('r', r);
    c.setAttribute('fill', 'none');
    c.setAttribute('stroke', '__BG__');
    c.setAttribute('stroke-width', __RSTK__);
    svg.appendChild(c);
  }
  gd.appendChild(svg);

  /* ---------- 文字层 ---------- */
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

  /* 标题：动态前缀右对齐到固定位置 + 固定正文从 x=239 起 */
  var pre = mk('left:0;width:__TPRX__px;text-align:right;top:__TIY__px;' +
               'font-family:' + __FT__ + ';font-size:__FST__px;font-weight:700;' +
               'line-height:1;color:__CT__;white-space:pre;', cur);
  mk('left:__TBX__px;top:__TIY__px;' +
     'font-family:' + __FT__ + ';font-size:__FST__px;font-weight:700;' +
     'line-height:1;color:__CT__;white-space:pre;', __TBODY__);

  /* 类目：右对齐到 x=445，行距 40 */
  for (var i = 0; i < N; i++) {
    mk('left:0;width:__CATRX__px;text-align:right;top:' + (__CATY0__ + __CATP__ * i) + 'px;' +
       'font-family:' + __FC__ + ';font-size:__FSC__px;line-height:1;' +
       'color:__CC__;white-space:pre;', __SERIES__[N - 1 - i][0]);
  }

  /* 脚注 */
  mk('left:__NX__px;top:__NY__px;font-family:' + __FN__ + ';font-size:__FSN__px;' +
     'line-height:1;color:__CN__;white-space:pre;', __NOTE__);

  /* 数值标签：落在环中线、弧的中点角度上 */
  var numL = [];
  for (var j = 0; j < N; j++) {
    var q = pt(LBL_R[j], VALS[cur][j] * 360);
    numL[j] = mk('left:' + q[0] + 'px;top:' + q[1] + 'px;transform:translate(-50%,-50%);' +
                 'font-family:' + __FNUM__ + ';font-size:__FSNUM__px;line-height:1;' +
                 'color:__CNUM__;white-space:pre;',
                 Math.round(VALS[cur][j] * 100) + '%');
  }

  function placeNums(v) {
    for (var j = 0; j < N; j++) {
      var q = pt(LBL_R[j], v[j] * 360);
      numL[j].style.left = q[0] + 'px';
      numL[j].style.top = q[1] + 'px';
    }
  }
  function setText(v) {
    for (var j = 0; j < N; j++) numL[j].textContent = Math.round(v[j] * 100) + '%';
  }

  /* ---------- 窗体按钮（照抄原文件的窗体按钮样式与位置） ---------- */
  var btnBox = document.createElement('div');
  btnBox.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;z-index:3;';
  gd.appendChild(btnBox);
  var btns = [];
  __BTNS__.forEach(function (b) {
    var el = document.createElement('div');
    el.textContent = b[2];
    el.style.cssText = 'position:absolute;left:' + (b[0] * __PPT__) + 'px;' +
      'top:' + (b[1] * __PPT__) + 'px;width:' + (45 * __PPT__) + 'px;' +
      'height:' + (16.5 * __PPT__) + 'px;box-sizing:border-box;' +
      'display:flex;align-items:center;justify-content:center;' +
      'background:#E1E1E1;border:1px solid #8B8B8B;border-radius:2px;' +
      'font-family:' + __FUI__ + ';font-size:13px;color:#000;cursor:pointer;user-select:none;';
    el.addEventListener('click', function () { go(b[2]); });
    btnBox.appendChild(el);
    btns.push(el);
  });

  function paintBtns() {
    for (var i = 0; i < btns.length; i++) {
      var on = btns[i].textContent === cur;
      btns[i].style.background = on ? '#C9C9C9' : '#E1E1E1';
    }
  }

  /* ---------- 切换动画 ---------- */
  var busy = false, vNow = VALS[cur].slice(), NFR = 12, MS = 22;

  function apply(v) {
    var y = [[v[0], __HALF__ - v[0]], [v[1], __HALF__ - v[1]],
             [v[2], __HALF__ - v[2]], [v[3], __HALF__ - v[3]]];
    Plotly.restyle(gd, {'values': y}, [0, 1, 2, 3]);
    placeNums(v);
  }

  function go(name) {
    if (busy || name === cur) return;
    busy = true;
    var from = vNow.slice(), to = VALS[name].slice(), k = 1;
    pre.textContent = name;
    cur = name;
    paintBtns();
    (function tick() {
      var t = k / NFR, e = t * t * (3 - 2 * t), v = [];
      for (var i = 0; i < N; i++) v.push(from[i] + (to[i] - from[i]) * e);
      apply(v);
      if (k >= NFR) { vNow = to.slice(); setText(to); busy = false; return; }
      k += 1;
      setTimeout(tick, MS);
    })();
  }

  paintBtns();
  document.addEventListener('keydown', function (e) {
    if (e.key === 'ArrowLeft') go('近7天');
    if (e.key === 'ArrowRight') go('近30天');
  });
})();
"""


def build_html():
    fig = build_figure(DEFAULT_PERIOD)
    script = (_POST_SCRIPT
              .replace("__PLID__", DIV_ID)
              .replace("__CX__", str(CX)).replace("__CY__", str(CY))
              .replace("__RIN__", str(R_IN)).replace("__ROUT__", str(R_OUT))
              .replace("__RADII__", json.dumps(RADII))
              .replace("__RSTEP__", str(R_STEP)).replace("__RSTK__", str(RING_STROKE))
              .replace("__HALF__", str(HALF_MAX))
              .replace("__SERIES__", json.dumps(SERIES, ensure_ascii=False))
              .replace("__VALUES__", json.dumps(VALUES, ensure_ascii=False))
              .replace("__PERIODS__", json.dumps(PERIODS, ensure_ascii=False))
              .replace("__CUR__", json.dumps(DEFAULT_PERIOD, ensure_ascii=False))
              .replace("__LBLR__", json.dumps(LBL_R))
              .replace("__BTNS__", json.dumps(BTN, ensure_ascii=False))
              .replace("__PPT__", str(PX_PER_PT))
              .replace("__W__", str(W)).replace("__H__", str(H))
              .replace("__BG__", BG)
              .replace("__TPRX__", str(TITLE_PREFIX_RX)).replace("__TBX__", str(TITLE_BODY_X))
              .replace("__TIY__", str(TITLE_INK_Y)).replace("__TBODY__",
                                                            json.dumps(TITLE_BODY, ensure_ascii=False))
              .replace("__CATRX__", str(CAT_RX)).replace("__CATY0__", str(CAT_Y0))
              .replace("__CATP__", str(CAT_PITCH))
              .replace("__NX__", str(NOTE_X)).replace("__NY__", str(NOTE_Y))
              .replace("__NOTE__", json.dumps(NOTE, ensure_ascii=False))
              .replace("__FT__", json.dumps(FONT_TITLE)).replace("__FC__", json.dumps(FONT_CAT))
              .replace("__FNUM__", json.dumps(FONT_NUM)).replace("__FN__", json.dumps(FONT_NOTE))
              .replace("__FUI__", json.dumps(FONT_UI))
              .replace("__FST__", str(FS_TITLE)).replace("__FSC__", str(FS_CAT))
              .replace("__FSNUM__", str(FS_NUM)).replace("__FSN__", str(FS_NOTE))
              .replace("__CT__", C_TITLE).replace("__CC__", C_CAT)
              .replace("__CN__", C_NOTE).replace("__CNUM__", C_NUM))
    html = fig.to_html(include_plotlyjs=True, div_id=DIV_ID, config=CONFIG,
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
    html_path = os.path.join(OUT_DIR, "6 VBA动态玉玦图_plotly.html")
    save_html(html_path)
    save_png(html_path, os.path.join(OUT_DIR, "6 VBA动态玉玦图_plotly.png"))


if __name__ == "__main__":
    main()
