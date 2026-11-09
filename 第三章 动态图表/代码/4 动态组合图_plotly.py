# -*- coding: utf-8 -*-
"""
「4 动态组合图」plotly 版：柱形（销售额 + 利润）+ 折线（利润率，次坐标轴），
深蓝底 + 可点击图例（对应源表里的三个复选框）。

参考基准：948 x 680（本机 Excel 导出图表 → 等比缩放到 948x680）
数据源  ：第三章 动态图表\\4 动态组合图.xlsx
           B17:F17  类目（口红/面膜/隔离/粉底液）
           B18:F18  销售额，B19:F19 利润，B20:F20 利润率
输出    ：...\\新建文件夹\\图\\第三章\\4 动态组合图_plotly.html
          ...\\新建文件夹\\图\\第三章\\4 动态组合图_plotly.png

一、这张图的结构（拆自 xl/charts/chart1.xml）
    · 柱：销售额 #0070C0、利润 #E74E69，簇状，gapWidth=219、overlap=-27
    · 折线：利润率 #FFC000，1.5pt 无标记，走**次坐标轴**（-1..1，格式 0%）
    · 两个数值轴都**不显示刻度标签**、没有网格线；没有图例
    · 图表区背景 #1A1E43

二、像素几何（单位 px，画布 948x680）
    · 绘图区 x 40.71..900.90，y 169.51..562.86（Excel 的 manualLayout 换算）
    · 主轴 0..4000 → 393.35 px；柱宽 48.2、同组两柱间隙 13.0
      （gapWidth/overlap 换算：catW=4.46*barW，4 个类目槽刚好铺满 860.19 px）
    · 类目中心 148.2 / 363.3 / 578.3 / 793.4
    · **次坐标轴不是 ±1 铺满绘图区**：实测 y = 364.22 - 193.89*v，
      也就是 v=-1 落在 y=558.1 而不是 562.9。所以 yaxis2 的 range 要按实测反解，
      直接给 [-1, 1] 会让折线两端各偏 2 px。

三、四处「照抄」的细节
    · 柱内数字是 Excel 的「数据标签·居中·在末端内侧」：水平居中、上沿压在柱顶往下 16.6 px，
      不是 plotly 默认的垂直居中，所以全部用 DOM 层画，不用 trace 的 text。
    · 利润率标签在折点正上方，标签**下沿**距折点 23 px。
    · 标题在 XML 里写的是「等线 Light + b=1」——等线 Light 没有粗体字形，Excel 做的是伪粗体，
      所以这里也用 0.3px 描边补那点加粗，而不是换成正体粗字。
    · 脚注在图表用户形状里（drawing2.xml），微软雅黑 8pt、bg1×85% = #D9D9D9。

四、两个实现选择
    · 柱的位置和宽度**不用 bargap/bargroupgap**：plotly 的换算式跟 Excel 不一样（实测柱子会窄 8 px），
      改成「1 类目槽 = 1 坐标单位」+ 显式 width，位置才能逐像素对上。
    · 导出的 PNG 加 --disable-lcd-text（灰度抗锯齿）：本轮的参考基准是 Excel 按 2 倍导出再缩到
      948x680 的，彩色次像素的成分已被平均掉，用灰度反而更接近（整图色差 9.74 → 9.58）。
      叠加前面的标题描边等修正后，最终整图平均色差 9.43/765、>60 的像素占 3.25%，
      其中柱区 3.67、折线区 10.11，残差集中在文字的灰度/彩色抗锯齿差异上。

五、交互
    标题下方那组彩色方块是可点的图例：点一下隐藏该系列（柱高归零 / 折线原地淡出），
    再点恢复；也可按键盘 1/2/3。它对应源表 C13:E13 的三个复选框。
    注意：参考图里本来没有图例（Excel 里图例被删掉了），这组方块是为了「可交互」补的，
    位置放在标题右侧的空白处，不压任何数据。
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
BASE = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹")
XLSX = os.path.join(BASE, "第三章 动态图表", "4 动态组合图.xlsx")
OUT_DIR = os.path.join(BASE, "图", "第三章")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# ==========================================================================
# 1. 数据
# ==========================================================================
CATS = ["口红", "面膜", "隔离", "粉底液"]
SALES = [2800, 2564, 2282, 2353]        # 销售额
PROFIT = [1896, 1563, 986, 1324]        # 利润
MARGIN = [0.677142857142857, 0.609594383775351,
          0.432077125328659, 0.562685932851679]   # 利润率
TITLE = "2022年化妆品销售情况"
NOTE = "*注：数据来源公司销售系统"
SERIES = [("销售额", "#0070C0"), ("利润额", "#E74E69"), ("利润率", "#FFC000")]


def load_data(path):
    """读 C17:F20（类目 + 三行数值）。读不到就用内置值。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        cats = [str(ws.cell(17, c).value).strip() for c in range(3, 7)]
        rows = []
        for r in (18, 19, 20):
            rows.append([float(ws.cell(r, c).value) for c in range(3, 7)])
        if not all(cats):
            raise ValueError("类目为空")
        return cats, rows[0], rows[1], rows[2]
    except Exception as exc:                      # noqa: BLE001
        print(f"[warn] 读取 {path} 失败（{exc}），改用内置数据")
        return CATS, SALES, PROFIT, MARGIN


CATS, SALES, PROFIT, MARGIN = load_data(XLSX)

# ==========================================================================
# 2. 画布与几何
# ==========================================================================
W, H = 948, 680
BG = "#1A1E43"

PX0, PX1 = 40.71, 900.90          # 绘图区左右
PY0, PY1 = 169.51, 562.86         # 绘图区上下（y 向下）
Y_MAX = 4000.0                    # 主轴量程

BAR_W = 48.2                      # 柱宽
GROUP_GAP = 13.0                  # 同组两柱间隙
# 柱不用 bargap/bargroupgap（plotly 的换算式与 Excel 不一致，实测柱会窄 8px），
# 改成显式坐标 + 显式宽度：1 个类目槽 = 1 个坐标单位。
SLOT = (PX1 - PX0) / float(len(CATS))
HALF_U = (BAR_W + GROUP_GAP) / 2.0 / SLOT      # 柱心相对组中心的偏移（单位）
WIDTH_U = BAR_W / SLOT

# 次坐标轴实测映射：y = Y2_INT - Y2_SLOPE * v
Y2_INT, Y2_SLOPE = 364.22, 193.89
Y2_LO = (Y2_INT - PY1) / Y2_SLOPE     # -1.0248（画在绘图区下沿时对应的 v）
Y2_HI = (Y2_INT - PY0) / Y2_SLOPE     # +1.0044

AXIS_COLOR = "#393252"            # 底轴线（lt1 50% 灰 @10% 叠在深蓝上）
LINE_W = 3.59                     # 折线 1.5pt
LINE_COLOR = "#FFC000"

# 文字（全部按参考图的墨迹位置摆）
FONT_TITLE = '"DengXian Light", "等线 Light", "DengXian", "等线", sans-serif'
FONT_TEXT = '"DengXian", "等线", "Microsoft YaHei", sans-serif'
FONT_NOTE = '"Microsoft YaHei", "微软雅黑", sans-serif'   # 脚注是图表用户形状，明确写死微软雅黑

FS_TITLE = 48.0
FS_CAT = 23.4
FS_BAR = 19.0
FS_PCT = 22.0
FS_NOTE = 18.8

C_TEXT = "#FFFFFF"
C_NOTE = "#D9D9D9"

TITLE_INK = (77.0, 48.0)
CAT_INK_Y = 583.0
BAR_LBL_DY = 16.6                 # 标签墨迹上沿 = 柱顶 + 17.6
PCT_LBL_DY = 36.5                 # 标签墨迹上沿 = 折点 y - 38（即下沿 -23）
NOTE_INK = (67.0, 638.0)

# div 盒子顶到墨迹顶的距离（line-height:1 时的字形上留白，用来把墨迹摆准）
INK_OFF_TITLE = 0.0
INK_OFF_CAT = 0.0
INK_OFF_BAR = 2.4
INK_OFF_PCT = 2.8
INK_OFF_NOTE = 0.4

# 图例（可点击，对应源表三个复选框）
LEG_X, LEG_Y = 79.0, 116.0
LEG_PITCH, LEG_SW, LEG_GAP = 121.0, 13.0, 9.0
FS_LEG = 20.0
C_LEG_ON, C_LEG_OFF = "#E8E9F0", "#7E839E"

DIV_ID = "combofig"
CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


def bar_top(v):
    return PY1 - (v / Y_MAX) * (PY1 - PY0)


def pct_y(v):
    return Y2_INT - Y2_SLOPE * v


def cat_x(i):
    slot = (PX1 - PX0) / len(CATS)
    return PX0 + slot * (i + 0.5)


# ==========================================================================
# 3. figure
# ==========================================================================
CEN_U = [i + 0.5 for i in range(len(CATS))]      # 组中心的坐标（单位）


def bar_x(side):
    """side=0 左柱（销售额），1 右柱（利润）。"""
    off = (-HALF_U) if side == 0 else HALF_U
    return [c + off for c in CEN_U]


def build_figure():
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=bar_x(0), y=SALES, width=WIDTH_U, name="销售额",
        marker=dict(color=SERIES[0][1], line=dict(width=0)),
        customdata=CATS,
        hovertemplate="<b>%{customdata}　销售额</b>　%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Bar(
        x=bar_x(1), y=PROFIT, width=WIDTH_U, name="利润额",
        marker=dict(color=SERIES[1][1], line=dict(width=0)),
        customdata=CATS,
        hovertemplate="<b>%{customdata}　利润额</b>　%{y:,.0f}<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=CEN_U, y=MARGIN, name="利润率", yaxis="y2", mode="lines",
        line=dict(color=LINE_COLOR, width=LINE_W, shape="linear"),
        customdata=CATS,
        hovertemplate="<b>%{customdata}　利润率</b>　%{y:.0%}<extra></extra>",
    ))

    fig.update_layout(
        width=W, height=H,
        margin=dict(l=0, r=0, t=0, b=0, autoexpand=False),
        paper_bgcolor=BG, plot_bgcolor=BG,
        showlegend=False,
        barmode="overlay",
        hovermode="closest",
        hoverlabel=dict(bgcolor="#2A2F5A", bordercolor="#4A5080",
                        font=dict(color="#F2F2F2", size=15, family=FONT_TEXT)),
        dragmode=False,
        xaxis=dict(
            type="linear", domain=[PX0 / W, PX1 / W],
            showticklabels=False, showgrid=False, zeroline=False,
            showline=False, ticks="", range=[0, len(CATS)],
            fixedrange=True,
        ),
        yaxis=dict(
            range=[0, Y_MAX], domain=[(H - PY1) / H, (H - PY0) / H],
            showticklabels=False, showgrid=False, zeroline=False,
            showline=False, ticks="", fixedrange=True,
        ),
        yaxis2=dict(
            range=[Y2_LO, Y2_HI], domain=[(H - PY1) / H, (H - PY0) / H],
            overlaying="y", side="right",
            showticklabels=False, showgrid=False, zeroline=False,
            showline=False, ticks="", fixedrange=True,
        ),
    )

    # 底轴线
    fig.add_shape(type="line", xref="paper", yref="paper",
                  x0=0, x1=1, y0=(H - PY1) / H, y1=(H - PY1) / H,
                  line=dict(color=AXIS_COLOR, width=1))
    return fig


# ==========================================================================
# 4. 页面里的 HTML 覆盖层 + 交互脚本
# ==========================================================================
_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLOT_ID__');
  if (!gd) return;
  gd.style.position = 'relative';

  var CATS = __CATS__;
  var DATA = __DATA__;          // [销售额, 利润, 利润率]
  var SER  = __SERIES__;        // [[名, 色], ...]
  var N    = CATS.length;
  var STEP = 26, NFRAME = 12;

  var ov = document.createElement('div');
  ov.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;' +
                     'pointer-events:none;user-select:none;font-family:' + __FONT__ + ';';
  gd.appendChild(ov);

  function mk(parent, css, txt) {
    var d = document.createElement('div');
    d.style.cssText = css;
    if (txt !== undefined) d.textContent = txt;
    parent.appendChild(d);
    return d;
  }

  /* ---- 标题 ---- */
  /* 标题：XML 里是「等线 Light + b=1」，Excel 于是做了伪粗体（原字体没有粗体字形），
     这里用 0.3px 描边模拟那点加粗，实测比纯 Light 更接近参考 */
  mk(ov, 'position:absolute;left:__TIX__px;top:__TIY__px;' +
         'font-family:' + __FONT_TITLE__ + ';' +
         'font-size:__FST__px;line-height:1;color:__CT__;white-space:pre;' +
         '-webkit-text-stroke:0.3px currentColor;', __TITLE__);

  /* ---- 脚注 ---- */
  mk(ov, 'position:absolute;left:__NIX__px;top:__NIY__px;' +
         'font-family:' + __FONT_NOTE__ + ';' +
         'font-size:__FSN__px;line-height:1;color:__CN__;white-space:pre;', __NOTE__);

  /* ---- 类目标签 ---- */
  for (var i = 0; i < N; i++) {
    mk(ov, 'position:absolute;left:' + __CATX__[i] + 'px;top:__CATY__px;' +
           'transform:translateX(-50%);font-size:__FSC__px;line-height:1;' +
           'color:__CT__;white-space:pre;', CATS[i]);
  }

  /* ---- 柱内数字：水平居中、上沿 = 柱顶 + 17.6 ---- */
  var barLbl = [];
  for (var s = 0; s < 2; s++) {
    barLbl[s] = [];
    for (var i = 0; i < N; i++) {
      barLbl[s][i] = mk(ov, 'position:absolute;left:' + __BARXL__[s][i] + 'px;' +
                            'top:' + (__BARTOP__[s][i] + __BAR_DY__ - __OFF_BAR__) + 'px;' +
                            'transform:translateX(-50%);font-size:__FSB__px;line-height:1;' +
                            'color:__CT__;white-space:pre;',
                        DATA[s][i].toFixed(0));
    }
  }

  /* ---- 利润率百分比：下沿距折点 23 ---- */
  var pctLbl = [];
  for (var i2 = 0; i2 < N; i2++) {
    pctLbl[i2] = mk(ov, 'position:absolute;left:' + __CATX__[i2] + 'px;' +
                          'top:' + (__PCTY__[i2] - __OFF_PCT__) + 'px;' +
                          'transform:translateX(-50%);font-size:__FSP__px;line-height:1;' +
                          'color:__CT__;white-space:pre;',
                      (DATA[2][i2] * 100).toFixed(0) + '%');
  }

  /* ---- 可点图例（对应源表 C13:E13 三个复选框） ---- */
  var on = [true, true, true];
  var busy = false;
  var items = [];
  for (var k = 0; k < 3; k++) {
    (function (idx) {
      var box = mk(ov, 'position:absolute;left:' + (__LEGX__ + idx * __LGP__) + 'px;' +
                        'top:__LEGY__px;display:flex;align-items:center;' +
                        'height:' + __FSL__ + 'px;pointer-events:auto;cursor:pointer;');
      box.id = 'lg' + idx;
      var sw = mk(box, 'display:inline-block;width:__SW__px;height:__SW__px;' +
                       'margin-right:__GAP__px;box-sizing:border-box;border-radius:1px;');
      var tx = mk(box, 'display:inline-block;font-size:__FSL__px;line-height:1;' +
                       'color:__CLON__;white-space:pre;', SER[idx][0]);
      sw.style.background = SER[idx][1];
      sw.style.border = '0 solid ' + SER[idx][1];
      box.onclick = function (e) { e.stopPropagation(); toggle(idx); };
      items.push({box: box, sw: sw, tx: tx});
    })(k);
  }

  function paint() {
    for (var k = 0; k < 3; k++) {
      var it = items[k];
      if (on[k]) {
        it.sw.style.background = SER[k][1];
        it.sw.style.border = '0 solid ' + SER[k][1];
        it.tx.style.color = __CLONJ__;
      } else {
        it.sw.style.background = 'transparent';
        it.sw.style.border = '1px solid ' + SER[k][1];
        it.tx.style.color = __CLOFFJ__;
      }
    }
  }

  /* ---- 显隐补间 ---- */
  function ease(t) { return t * t * (3.0 - 2.0 * t); }

  function apply(k, sc) {
    if (k < 2) {                       // 柱：高度归零
      Plotly.restyle(gd, {y: [DATA[k].map(function (v) { return v * sc; })]}, [k]);
    } else {                           // 折线：原地淡出（压到 0 会跑到图中线附近，难看）
      Plotly.restyle(gd, {opacity: sc}, [k]);
    }
  }

  function fade(k, show) {
    var els = k < 2 ? barLbl[k] : pctLbl;
    for (var i = 0; i < els.length; i++) {
      els[i].style.transition = 'opacity 200ms';
      els[i].style.opacity = show ? 1 : 0;
    }
  }

  function toggle(k) {
    if (busy) return;
    busy = true;
    var from = on[k] ? 1 : 0, to = on[k] ? 0 : 1;
    on[k] = !on[k];
    fade(k, to === 1);
    paint();
    var f = 0;
    (function tick() {
      var sc = from + (to - from) * ease(f / (NFRAME - 1.0));
      apply(k, sc);
      if (f >= NFRAME - 1) { busy = false; return; }
      f += 1;
      setTimeout(tick, STEP);
    })();
  }

  document.addEventListener('keydown', function (e) {
    var n = parseInt(e.key, 10);
    if (n >= 1 && n <= 3) toggle(n - 1);
  });

  paint();
})();
"""

CSS_FIX = ("<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
           ".plotly-graph-div{overflow:visible !important;}</style>")


def build_html():
    fig = build_figure()
    bar_centers = [[cat_x(i) - (BAR_W + GROUP_GAP) / 2 for i in range(len(CATS))],
                   [cat_x(i) + (BAR_W + GROUP_GAP) / 2 for i in range(len(CATS))]]

    script = (_POST_SCRIPT
              .replace("__PLOT_ID__", DIV_ID)
              .replace("__CATS__", json.dumps(CATS, ensure_ascii=False))
              .replace("__DATA__", json.dumps([SALES, PROFIT, MARGIN]))
              .replace("__SERIES__", json.dumps(SERIES, ensure_ascii=False))
              .replace("__W__", str(W)).replace("__H__", str(H))
              .replace("__FONT__", json.dumps(FONT_TEXT))
              .replace("__FONT_TITLE__", json.dumps(FONT_TITLE))
              .replace("__FONT_NOTE__", json.dumps(FONT_NOTE))
              .replace("__TITLE__", json.dumps(TITLE, ensure_ascii=False))
              .replace("__NOTE__", json.dumps(NOTE, ensure_ascii=False))
              .replace("__TIX__", str(TITLE_INK[0]))
              .replace("__TIY__", str(TITLE_INK[1] - INK_OFF_TITLE))
              .replace("__NIX__", str(NOTE_INK[0]))
              .replace("__NIY__", str(NOTE_INK[1] - INK_OFF_NOTE))
              .replace("__FST__", str(FS_TITLE)).replace("__FSN__", str(FS_NOTE))
              .replace("__FSC__", str(FS_CAT)).replace("__FSB__", str(FS_BAR))
              .replace("__FSP__", str(FS_PCT)).replace("__FSL__", str(FS_LEG))
              .replace("__CT__", C_TEXT).replace("__CN__", C_NOTE)
              .replace("__CLONJ__", json.dumps(C_LEG_ON))
              .replace("__CLOFFJ__", json.dumps(C_LEG_OFF))
              .replace("__CLON__", C_LEG_ON).replace("__CLOFF__", C_LEG_OFF)
              .replace("__CATX__", json.dumps([round(cat_x(i), 2) for i in range(len(CATS))]))
              .replace("__CATY__", str(CAT_INK_Y - INK_OFF_CAT))
              .replace("__BARTOP__", json.dumps([[round(bar_top(v), 2) for v in SALES],
                                                 [round(bar_top(v), 2) for v in PROFIT]]))
              .replace("__BARXL__", json.dumps(bar_centers))
              .replace("__PCTY__", json.dumps([round(pct_y(v), 2) for v in MARGIN]))
              .replace("__BAR_DY__", str(BAR_LBL_DY)).replace("__OFF_BAR__", str(INK_OFF_BAR))
              .replace("__OFF_PCT__", str(PCT_LBL_DY + INK_OFF_PCT))
              .replace("__LEGX__", str(LEG_X)).replace("__LEGY__", str(LEG_Y))
              .replace("__LGP__", str(LEG_PITCH)).replace("__SW__", str(LEG_SW))
              .replace("__GAP__", str(LEG_GAP)))
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
                        "--disable-lcd-text",   # 灰度抗锯齿，与参考图（2 倍导出后缩放）更接近
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
    html_path = os.path.join(OUT_DIR, "4 动态组合图_plotly.html")
    save_html(html_path)
    save_png(html_path, os.path.join(OUT_DIR, "4 动态组合图_plotly.png"))


if __name__ == "__main__":
    main()
