# -*- coding: utf-8 -*-
"""
「2 动态跑道图」plotly 版：深蓝底 + 6 圈同心圆弧 + 部门标签 + 右上月份下拉框。

参考图：948 x 680（1月 ~ 4月 四张截图，样式一致）
数据源：第三章 动态图表\\2 动态跑道图.xlsx
          B3:F9   源数据（部门 + 1月~4月）
          H3:I10  图表数据（部门 + 当月值），I 列由 VLOOKUP 取当月
          G17     控件单元格（下拉框的链接格，决定初始月份）
输出  ：...\\新建文件夹\\图\\第三章\\2 动态跑道图_plotly.html
        ...\\新建文件夹\\图\\第三章\\2 动态跑道图-<月份>_plotly.png

一、这张图为什么叫「跑道」
每圈 = 一个部门，圆心在 (423.63, 381.82)，6 圈由外向内按人数降序排列。
弧长比例的分母不是最大值，而是「当月总人数 / 2」——即 Excel 里
    占位 = SUM(I4:I9)*0.5 - I4
的两段式圆环：彩色段 + 占位段，两段加起来正好半圈多一点，所以每圈都不会闭环。
    扫角 = 人数 / (当月总人数/2) × 360°
实测四个月全部吻合（1 月 6 圈误差 ≤ 0.4°）。

二、几何（从参考图逐像素量测）
    · 圆心 (423.63, 381.82)，最外环外径 228.33 px，环间距 21.31 px
    · 环带宽：参考图实测 12.1~12.5 px（Excel 的 3pt 背景色描边吃掉了一部分）；
      代码取 BAND = 13.45，略放大以抵消抗锯齿吃掉的那一像素
    · 所有弧从 12 点方向起，顺时针；首尾各内缩 4.45 px 弧长（同上，描边所致）
    · 颜色由内向外：紫 #7030A0 / 蓝 #0070C0 / 浅蓝 #37A2DA
                    青 #11A7AD / 黄 #F5C353 / 红 #E74E69

三、一个必须照抄的「怪癖」
左侧标签列按源表行序自上而下（人力部在最上），而彩色圈是**由外向内**降序排列，
两者顺序并不对应——参考图就这样，Excel 的数据标签也不在图上，标签是独立文本框。

四、两处「照抄」的细节
  · 标签字形宽扁：参考图里那一列是 12px 等线被横向拉宽了约 1.42 倍（数字只有 9px 高），
    所以用 CSS 的 scaleX 还原，而不是换个字号——换字号会把字宽也一起改掉。
  · 下拉框右下角有一小段 1px 白线（Excel 组合框的尺寸手柄残影），也一并补上了。

五、交互
    点右上角下拉框 → 弹出月份列表；也可以按键盘 1~4 切换。
    切换时弧长按 14 帧补间（28ms/帧，缓入缓出），文字在最后一帧更新。
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
# 0. 路径
# ==========================================================================
BASE = (r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹")
XLSX = os.path.join(BASE, "第三章 动态图表", "2 动态跑道图.xlsx")
OUT_DIR = os.path.join(BASE, "图", "第三章")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

# ==========================================================================
# 1. 数据
# ==========================================================================
MONTHS = ["1月", "2月", "3月", "4月"]

# 兜底数据（与 xlsx 的 B4:F9 一致），读不到文件时用
DEPTS = ["人力部", "行政部", "财务部", "工程部", "采购部", "销售部"]
VALUES = [[130, 138, 130, 129],
          [226, 206, 217, 226],
          [238, 228, 255, 226],
          [293, 305, 314, 264],
          [326, 349, 316, 359],
          [451, 456, 406, 433]]
CTRL_MONTH = "1月"          # xlsx 里 G17 的值


def load_data(path):
    """读源数据块 B3:F9 与控件单元格 G17。"""
    try:
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.worksheets[0]
        depts, vals = [], []
        for r in range(4, 10):
            name = ws.cell(r, 2).value
            if name is None:
                raise ValueError("部门列为空")
            depts.append(str(name).strip())
            vals.append([float(ws.cell(r, c).value) for c in range(3, 7)])
        g17 = ws.cell(17, 7).value
        return depts, vals, (str(g17).strip() if g17 else "1月")
    except Exception as exc:                     # noqa: BLE001
        print(f"[warn] 读取 {path} 失败（{exc}），改用内置数据")
        return DEPTS, VALUES, CTRL_MONTH


DEPTS, VALUES, CTRL_MONTH = load_data(XLSX)
if CTRL_MONTH not in MONTHS:
    CTRL_MONTH = MONTHS[0]
CUR0 = MONTHS.index(CTRL_MONTH)

# ==========================================================================
# 2. 画布与几何（单位：像素）
# ==========================================================================
W, H = 948, 680
BG = "#1A1E43"

CX, CY = 423.63, 381.82
R_IN0 = 109.60          # 最内圈（人力部）内半径
PITCH = 21.31           # 相邻两圈的半径差
BAND = 13.45            # 可见填充带宽度
R_MAX = 240.0           # 极坐标轴量程，画布上一格 = 1px
INSET_PX = 4.45         # 弧首尾各内缩的弧长（Excel 3pt 描边留下的缺口）

# ==========================================================================
# 3. 配色与字体
# ==========================================================================
# 由内向外，与 DEPTS 顺序一致（人力部在最内圈）
COLORS = ["#7030A0", "#0070C0", "#37A2DA", "#11A7AD", "#F5C353", "#E74E69"]

C_TITLE = "#F2F2F2"
C_LABEL = "#F2F2F2"
C_NOTE = "#D9D9D9"
C_DD_BORDER = "#109968"
C_DD_BTN_BORDER = "#7F9DB9"

FONT_TITLE = '"DengXian Light", "等线 Light", "DengXian", "等线", sans-serif'
FONT_TEXT = '"DengXian", "等线", "Microsoft YaHei", sans-serif'

FS_TITLE = 48.2
FS_LABEL = 12.0
FS_NOTE = 21.45
FS_MONTH = 26.6

# 文字锚点（参考图实测的墨迹位置，不是盒子的位置）
TITLE_X, TITLE_Y = 65.0, 57.0            # 墨迹左上
NOTE_X, NOTE_Y = 75.0, 630.8
LABEL_ROWS = [154.0, 174.0, 195.0, 215.0, 235.0, 256.0]   # 每行墨迹上沿
LABEL_X0 = 336.0                          # 名字布局盒左边界（缩放前）
NAME_W, NAME_GAP, NUM_W = 36.0, 21.0, 18.0   # 布局盒宽度（缩放前的字号宽度）
LABEL_SX = 1.42                           # 横向拉伸比
LABEL_DY = 0.0                            # 整行位置微调
LABEL_DY_N = 1.0                          # 名字纵向微调
LABEL_DY_V = -1.0                         # 数字纵向微调

# 下拉框
DD_X, DD_Y, DD_W, DD_H, DD_BW = 678.0, 71.0, 138.0, 39.0, 3.0
DD_BTN_X, DD_BTN_Y, DD_BTN = 816.0, 78.0, 30.0
DD_TEXT_X, DD_TEXT_Y = 684.0, 84.0
DD_MONTH_X, DD_MONTH_Y = 678.0, 112.0    # 弹出列表左上角（贴着框的下沿）
DD_GRIP_X, DD_GRIP_Y = 812.0, 105.0      # 组合框右下角 1px 白色折角（照抄）

DIV_ID = "trackfig"
CONFIG = {"displayModeBar": False, "scrollZoom": False,
          "doubleClick": False, "responsive": False}


# ==========================================================================
# 4. 角度计算
# ==========================================================================
def ring_mid(k):
    """第 k 圈（0 = 最内）的中径。"""
    return R_IN0 + PITCH * k + BAND / 2.0


def sweeps(values):
    """返回 6 组 (起始角, 终止角)，单位度。分母是当月总人数的一半。"""
    total = float(sum(values))
    out = []
    for k, v in enumerate(values):
        full = (v / (total / 2.0)) * 360.0 if total > 0 else 0.0
        inset = math.degrees(INSET_PX / ring_mid(k))
        a0 = inset
        a1 = max(a0, full - inset)
        out.append((a0, a1))
    return out


def angles(values):
    """给 Barpolar 用的 (中心角, 张角)，单位弧度。"""
    th, wd = [], []
    for a0, a1 in sweeps(values):
        th.append(math.radians((a0 + a1) / 2.0))
        wd.append(math.radians(a1 - a0))
    return th, wd


# ==========================================================================
# 5. figure
# ==========================================================================
def build_figure(values):
    th, wd = angles(values)
    fig = go.Figure()
    fig.add_trace(go.Barpolar(
        r=[BAND] * len(DEPTS),
        base=[R_IN0 + PITCH * k for k in range(len(DEPTS))],
        theta=th, width=wd,
        thetaunit="radians",
        marker=dict(color=COLORS, line=dict(width=0)),
        name="",
        showlegend=False,
        hovertemplate="<b>%{customdata}</b><extra></extra>",
        customdata=[f"{d}　{v:.0f}" for d, v in zip(DEPTS, values)],
    ))

    half = R_MAX / W, R_MAX / H
    fig.update_layout(
        width=W, height=H,
        # autoexpand 默认开着时 plotly 会为了容纳溢出画布的标签放大 margin，
        # domain 就不再对应真实像素
        margin=dict(l=0, r=0, t=0, b=0, autoexpand=False),
        paper_bgcolor=BG, plot_bgcolor=BG,
        showlegend=False,
        hovermode="closest",
        dragmode=False,
        polar=dict(
            domain=dict(x=[CX / W - half[0], CX / W + half[0]],
                        y=[1 - CY / H - half[1], 1 - CY / H + half[1]]),
            bgcolor="rgba(0,0,0,0)",
            radialaxis=dict(range=[0, R_MAX], visible=False, showgrid=False,
                            showline=False, ticks=""),
            angularaxis=dict(rotation=90, direction="clockwise", visible=False,
                             showgrid=False, showline=False, ticks=""),
        ),
    )
    return fig


# ==========================================================================
# 6. 帧数据
# ==========================================================================
NFRAME = 14


def ease(t):
    return t * t * (3.0 - 2.0 * t)


def frame_payloads(cur):
    """FR[目标月][帧 k] = {t: 中心角, w: 张角, cd: 悬停文本}。"""
    out = []
    a = [VALUES[i][cur] for i in range(len(DEPTS))]
    for dst in range(len(MONTHS)):
        b = [VALUES[i][dst] for i in range(len(DEPTS))]
        seq = []
        for k in range(NFRAME):
            t = ease(k / (NFRAME - 1.0))
            v = [x + (y - x) * t for x, y in zip(a, b)]
            th, wd = angles(v)
            seq.append(dict(t=[round(x, 6) for x in th],
                            w=[round(x, 6) for x in wd],
                            cd=[f"{d}　{q:.0f}" for d, q in zip(DEPTS, v)]))
        out.append(seq)
    return out


def texts():
    out = []
    for m, name in enumerate(MONTHS):
        total = int(round(sum(VALUES[i][m] for i in range(len(DEPTS)))))
        out.append(dict(title=f"{name}公司总人数为{total}",
                        month=name,
                        nums=[f"{VALUES[i][m]:.0f}" for i in range(len(DEPTS))]))
    return out


# ==========================================================================
# 7. 页面里的 HTML 覆盖层 + 交互脚本
# ==========================================================================
_POST_SCRIPT = r"""
(function () {
  var gd = document.getElementById('__PLOT_ID__');
  if (!gd) return;
  gd.style.position = 'relative';

  var FR  = __FRAMES__;
  var TX  = __TEXTS__;
  var DN  = __DEPTS__;                       // 部门名（由内向外）
  var N   = FR[0].length;
  var STEP = 28;
  var cur = __CUR__;
  var busy = false;

  var ov = document.createElement('div');
  ov.style.cssText = 'position:absolute;left:0;top:0;width:__W__px;height:__H__px;' +
                     'pointer-events:none;user-select:none;font-family:' + __FONT__ + ';';
  gd.appendChild(ov);

  /* ---- 标题 ---- */
  var title = document.createElement('div');
  title.id = 'ttl';
  title.textContent = TX[cur].title;
  title.style.cssText = 'position:absolute;left:__TX__px;top:__TY__px;' +
    'font-size:__FST__px;line-height:1;color:__CT__;white-space:pre;';
  ov.appendChild(title);

  /* ---- 脚注 ---- */
  var note = document.createElement('div');
  note.textContent = __NOTE__;
  note.style.cssText = 'position:absolute;left:__NX__px;top:__NY__px;' +
    'font-size:__FSN__px;line-height:1;color:__CN__;white-space:pre;';
  ov.appendChild(note);

  /* ---- 6 行部门标签 ----
     参考图里的标签是 12px 等线被横向拉宽了约 1.4 倍（字形宽扁、数字只有 9px 高），
     所以这里不改字号，改用 scaleX 还原；两段都以右边界为基准拉伸，右对齐不动。 */
  var nums = [];
  for (var i = 0; i < DN.length; i++) {
    var row = document.createElement('div');
    row.style.cssText = 'position:absolute;top:' + (__ROWS__[i] + __LDY__) + 'px;left:' +
      __LX0__ + 'px;height:__FSL__px;white-space:nowrap;';

    var nm = document.createElement('span');
    nm.textContent = DN[i];
    nm.style.cssText = 'display:inline-block;width:__NAME_W__px;text-align:right;' +
      'font-size:__FSL__px;line-height:1;color:__CL__;position:relative;top:__LDY_N__px;' +
      'transform:scaleX(__LSX__);transform-origin:right top;';

    var vl = document.createElement('span');
    vl.textContent = TX[cur].nums[i];
    vl.style.cssText = 'display:inline-block;width:__NUM_W__px;text-align:right;' +
      'margin-left:__NAME_GAP__px;font-size:__FSL__px;line-height:1;color:__CL__;' +
      'position:relative;top:__LDY_V__px;' +
      'transform:scaleX(__LSX__);transform-origin:right top;';

    row.appendChild(nm);
    row.appendChild(vl);
    ov.appendChild(row);
    nums.push(vl);
  }

  /* ---- 下拉框：绿框 + 月份 + 右侧箭头按钮 + 弹出的列表 ---- */
  var dd = document.createElement('div');
  dd.id = 'ddbox';
  dd.style.cssText = 'position:absolute;left:__DDX__px;top:__DDY__px;' +
    'width:__DDW__px;height:__DDH__px;border:__DDBW__px solid __DDB__;' +
    'box-sizing:border-box;pointer-events:auto;cursor:pointer;';
  var mtxt = document.createElement('span');
  mtxt.textContent = TX[cur].month;
  mtxt.style.cssText = 'position:absolute;left:' + (__DTX__ - __DDX__ - __DDBW__) + 'px;top:' +
    (__DTY__ - __DDY__ - __DDBW__ - 3) + 'px;font-size:__FSM__px;line-height:1;color:__CT__;';
  dd.appendChild(mtxt);

  var grip = document.createElement('div');
  grip.style.cssText = 'position:absolute;left:__GRIPX__px;top:__GRIPY__px;width:4px;height:6px;' +
    'border-left:1px solid #FFFFFF;border-top:1px solid #FFFFFF;box-sizing:border-box;';
  ov.appendChild(grip);
  ov.appendChild(dd);

  var btn = document.createElement('div');
  btn.style.cssText = 'position:absolute;left:__DBX__px;top:__DBY__px;' +
    'width:__DBTN__px;height:__DBTN__px;border:1px solid __DBB2__;box-sizing:border-box;' +
    'background:linear-gradient(90deg,#FBFBFB,#F5F5F5);pointer-events:auto;cursor:pointer;';
  var tri = document.createElement('div');
  tri.style.cssText = 'position:absolute;left:50%;top:50%;margin:-3px 0 0 -5px;' +
    'width:0;height:0;border-left:5px solid transparent;border-right:5px solid transparent;' +
    'border-top:7px solid #000;';
  btn.appendChild(tri);
  ov.appendChild(btn);

  var list = document.createElement('div');
  list.id = 'ddlist';
  list.style.cssText = 'position:absolute;left:__DMX__px;top:__DDLY__px;width:__DDW__px;' +
    'background:#FFFFFF;border:1px solid #7F9DB9;box-sizing:border-box;display:none;' +
    'pointer-events:auto;z-index:20;box-shadow:0 2px 6px rgba(0,0,0,0.45);';
  var items = [];
  for (var m = 0; m < TX.length; m++) {
    (function (idx) {
      var it = document.createElement('div');
      it.id = 'ddit' + idx;
      it.textContent = TX[idx].month;
      it.style.cssText = 'height:32px;line-height:32px;padding-left:9px;font-size:20px;' +
        'color:#1A1A1A;cursor:pointer;';
      it.onmouseenter = function () { it.style.background = '#DCE8F5'; };
      it.onmouseleave = function () { hint(idx); };
      it.onclick = function (e) { e.stopPropagation(); closeList(); play(idx); };
      list.appendChild(it);
      items.push(it);
    })(m);
  }
  ov.appendChild(list);

  function hint(sel) {
    for (var i = 0; i < items.length; i++) {
      items[i].style.background = (i === sel) ? '#DCE8F5' : 'transparent';
    }
  }

  function openList() { list.style.display = 'block'; hint(cur); }
  function closeList() { list.style.display = 'none'; }
  dd.onclick = function (e) {
    e.stopPropagation();
    list.style.display === 'block' ? closeList() : openList();
  };
  btn.onclick = dd.onclick;
  document.addEventListener('click', closeList);

  /* ---- 状态切换 ---- */
  function apply(k, dst) {
    var f = FR[dst][k];
    Plotly.restyle(gd, {theta: [f.t], width: [f.w], customdata: [f.cd]}, [0]);
  }

  function setText(m) {
    title.textContent = TX[m].title;
    mtxt.textContent = TX[m].month;
    for (var i = 0; i < nums.length; i++) nums[i].textContent = TX[m].nums[i];
    hint(m);
  }

  function play(dst) {
    if (busy || dst === cur) return;
    busy = true;
    var k = 1;
    (function tick() {
      apply(k, dst);
      if (k >= N - 1) { setText(dst); cur = dst; busy = false; return; }
      k += 1;
      setTimeout(tick, STEP);
    })();
  }

  document.addEventListener('keydown', function (e) {
    var n = parseInt(e.key, 10);
    if (n >= 1 && n <= TX.length) play(n - 1);
  });
})();
"""

CSS_FIX = ("<style>html,body{margin:0;padding:0;overflow:hidden;background:" + BG + ";}"
           ".plotly-graph-div{overflow:visible !important;}</style>")


def build_html(values, cur):
    fig = build_figure(values)
    script = (_POST_SCRIPT
              .replace("__PLOT_ID__", DIV_ID)
              .replace("__FRAMES__", json.dumps(frame_payloads(cur), ensure_ascii=False))
              .replace("__TEXTS__", json.dumps(texts(), ensure_ascii=False))
              .replace("__DEPTS__", json.dumps(DEPTS, ensure_ascii=False))
              .replace("__W__", str(W)).replace("__H__", str(H))
              .replace("__FONT__", json.dumps(FONT_TEXT))
              .replace("__TX__", str(TITLE_X)).replace("__TY__", str(TITLE_Y))
              .replace("__FST__", str(FS_TITLE)).replace("__CT__", C_TITLE)
              .replace("__NOTE__", json.dumps("*注：数据来源公司人力资源系统", ensure_ascii=False))
              .replace("__NX__", str(NOTE_X)).replace("__NY__", str(NOTE_Y))
              .replace("__FSN__", str(FS_NOTE)).replace("__CN__", C_NOTE)
              .replace("__ROWS__", json.dumps(LABEL_ROWS))
              .replace("__NAME_W__", str(NAME_W)).replace("__NAME_GAP__", str(NAME_GAP))
              .replace("__LX0__", str(LABEL_X0)).replace("__LSX__", str(LABEL_SX))
              .replace("__LDY__", str(LABEL_DY))
              .replace("__LDY_N__", str(LABEL_DY_N))
              .replace("__LDY_V__", str(LABEL_DY_V))
              .replace("__NUM_W__", str(NUM_W))
              .replace("__FSL__", str(FS_LABEL)).replace("__CL__", C_LABEL)
              .replace("__DDX__", str(DD_X)).replace("__DDY__", str(DD_Y))
              .replace("__DDW__", str(DD_W)).replace("__DDH__", str(DD_H))
              .replace("__DDBW__", str(DD_BW)).replace("__DDB__", C_DD_BORDER)
              .replace("__DTX__", str(DD_TEXT_X)).replace("__DTY__", str(DD_TEXT_Y))
              .replace("__FSM__", str(FS_MONTH))
              .replace("__DBX__", str(DD_BTN_X)).replace("__DBY__", str(DD_BTN_Y))
              .replace("__DBTN__", str(DD_BTN))
              .replace("__DBB2__", C_DD_BTN_BORDER)
              .replace("__DMX__", str(DD_MONTH_X)).replace("__DDLY__", str(DD_MONTH_Y))
              .replace("__GRIPX__", str(DD_GRIP_X)).replace("__GRIPY__", str(DD_GRIP_Y))
              .replace("__CUR__", str(cur)))
    html = fig.to_html(include_plotlyjs=True, div_id=DIV_ID, config=CONFIG,
                       full_html=True, post_script=script)
    return html.replace("</head>", CSS_FIX + "</head>")


# ==========================================================================
# 8. 导出
# ==========================================================================
def save_html(path):
    html = build_html([VALUES[i][CUR0] for i in range(len(DEPTS))], CUR0)
    with open(path, "w", encoding="utf-8") as fh:
        fh.write(html)
    print("[ok] %s  (%.1f MB)" % (path, os.path.getsize(path) / 1048576))


def save_png(html_path, out_path):
    prof = tempfile.mkdtemp(prefix="chrome_")
    try:
        subprocess.run([CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars",
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
    html_path = os.path.join(OUT_DIR, "2 动态跑道图_plotly.html")
    save_html(html_path)

    # 四个月各截一张：临时页面把初始月份改成对应值
    for m, name in enumerate(MONTHS):
        p = os.path.join(OUT_DIR, "2 动态跑道图-%s_plotly.png" % name)
        if m == CUR0:
            save_png(html_path, p)
            continue
        tmp_html = os.path.join(OUT_DIR, "_tmp_%s.html" % name)
        with open(tmp_html, "w", encoding="utf-8") as fh:
            fh.write(build_html([VALUES[i][m] for i in range(len(DEPTS))], m))
        save_png(tmp_html, p)
        os.remove(tmp_html)


if __name__ == "__main__":
    main()
