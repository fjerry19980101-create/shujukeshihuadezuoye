# -*- coding: utf-8 -*-
"""
第四章「销售看板」—— plotly 可交互复现
=====================================

数据源：`第四章 销售看板参考.xlsx`
    · 销售明细（A1:J8565，8564 条）
    · 成本明细（A1:C49，12 个月 × 4 类成本）
输出：同目录下
    · 第四章 销售看板_复现.html   可交互看板
    · 第四章 销售看板_复现.png    1600×925 静态图

原件结构
--------
「数据大屏」工作表上浮着 11 个图表 + 2 个矩形（标题带/底带）+ 2 个文本框（主标题、成本费用）
+ 1 个日历图标 + 1 个窗体下拉框（月份）。本脚本按原件的图形锚点（pt）换算成设计像素
（画布 1600×925，比例 1.79775 px/pt，原点 = 原件图表区左上角 218.2,19.1）逐个还原，
再用导出图（1600×925）逐块标定圆心/半径/条宽/字号/文字基线。

11 个图表
---------
 产品类别销量  环形图（3 类，孔 85%）      销量环比    横向条形（本月/上月）
 各月销售额    渐变面积图（12 个月）       商品销售额Top3  隐形堆叠条 + 手动标签（改为 HTML 表）
 区域销量      柱形（6 区域）              快递公司销量  柱形（6 家）
 成本费用      2×2 四个小环图              利润率      环形图 + 药丸按钮
（成本费用的四个小环、利润率的环，data label 与标题都在图表里的 userShapes，这里一律用 DOM 文字层覆盖。）

颜色（取自 chart XML + 导出图采样）
----------------------------------
 画布 #1A1E43（卡片间隙露出）· 卡片 #2C3560 · 标题带 #2C355F
 青 #7BBDD5 · 粉 #FC6C9D · 靛 #4A5BD1 · 橙 #F6A576 · 文字 #F2F2F2
 环的底槽 #545B7D · 面积渐变 #4A5BD1 → #333D79 · 药丸描边 #4C5FDC

交互
----
右上角「月份」下拉（对应原件的窗体组合框）：切换后所有卡片联动重算，条长/弧长补间，
数值、标题、Top3 榜单同步刷新；图表支持 hover 提示。键盘 ↑↓ 换月、1-9/0 直选。

关于「销量环比」的一处偏差（原件的缺陷，非本脚本）
------------------------------------------------
原件里 8 月那条的长度只有 9 月的 48%，但两个值是 714 / 751（应几乎等长）。实测把 8 月
改成 1502 后 Excel 才按比例重画，说明原图是**旧缓存**没刷新。本脚本按真实比例绘制，
本月条长与原件一致（751 单 → 322px），上月条按同一基准轴换算，因此会比原件长。

其它两处刻意的处理
------------------
1. 成本小卡里「214.67」在原件的窄文本框中折成「214.6 / 7」、产品卡折成「10.80 / %」，
   这里一律排成一行。
2. 饼图外圈的 54% / 21% / 25% 三个标签，原件是手工摆在扇区旁的（Excel「最佳位置」），
   本脚本改为按扇区中点自动定位，切月时随数据一起移动 —— 9 月与原件相差不超过 6px。
"""

import base64
import io
import json
import os
import subprocess
import tempfile
from urllib.parse import quote

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from PIL import Image
from scipy import ndimage

# ==========================================================================
# 0. 路径与画布
# ==========================================================================
BASE = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\作业"
SRC = os.path.join(BASE, "第四章 销售看板参考.xlsx")
OUT_DIR = os.path.join(BASE, "《Excel数据可视化——从图表到数据大屏》", "第四章 销售看板参考")
HTML_PATH = os.path.join(OUT_DIR, "第四章 销售看板_复现.html")
PNG_PATH = os.path.join(OUT_DIR, "第四章 销售看板_复现.png")
CHROME = r"C:\Program Files\Google\Chrome\Application\chrome.exe"

S = 1600.0 / 890.0          # 1pt = 1.79775 设计像素
OX, OY = 218.2, 19.1        # 原件图表区左上角（pt）
W, H = 1600, 925


def X(pt):
    return (pt - OX) * S


def Y(pt):
    return (pt - OY) * S


# ==========================================================================
# 1. 配色
# ==========================================================================
C_BG = "#1A1E43"        # 画布（卡片缝隙）
C_CARD = "#2C3560"      # 卡片
C_BAND = "#2C355F"      # 标题带 / 底带
C_TEXT = "#F2F2F2"      # bg1 95%
C_TEAL = "#7BBDD5"
C_PINK = "#FC6C9D"
C_INDIGO = "#4A5BD1"
C_ORANGE = "#F6A576"
C_TRACK = "#545B7D"     # 环的底槽
C_GRAD0 = "#333D79"     # 面积渐变底部
C_PILL = "#4C5FDC"      # 药丸描边

# ==========================================================================
# 2. 数据
# ==========================================================================
MONTHS = ["%d月" % i for i in range(1, 13)]
DEFAULT_MONTH = "9月"                      # 原件组合框的默认值

CAT_ORDER = ["办公用品", "家具产品", "数码电子"]
CAT_COLORS = [C_PINK, C_INDIGO, C_ORANGE]
REGION_ORDER = ["华北", "华南", "东北", "西北", "西南", "华东"]
EXPRESS_ORDER = ["顺丰", "韵达", "中通", "申通", "圆通", "EMS"]
COST_ORDER = ["推广", "人工", "产品", "其他"]      # 2×2 卡片的排列顺序（c1/c3/c4/c2）


def load():
    sale = pd.read_excel(SRC, sheet_name="销售明细")
    cost = pd.read_excel(SRC, sheet_name="成本明细")
    sale["月"] = sale["月份_销售"].astype(str)
    cost["月"] = cost["月份_成本"].astype(str)

    year = [round(float(sale.loc[sale["月"] == m, "订单额"].sum()) / 1e4, 4) for m in MONTHS]

    data = {}
    for i, m in enumerate(MONTHS):
        df = sale[sale["月"] == m]
        amt = float(df["订单额"].sum())
        pro = float(df["利润额"].sum())
        cst = cost[cost["月"] == m].set_index("类别")["成本"].to_dict()
        cst_vals = [round(float(cst.get(k, 0.0)), 4) for k in COST_ORDER]
        cst_sum = sum(cst_vals) or 1.0
        top = (df.groupby("产品名称")
                 .agg(金额=("订单额", "sum"), 数量=("订单额", "size"))
                 .sort_values("金额", ascending=False).head(3))
        prev = MONTHS[i - 1] if i > 0 else MONTHS[-1]
        data[m] = dict(
            qty=int(len(df)),
            prev=prev,
            prevQty=int((sale["月"] == prev).sum()),
            amt=round(amt / 1e4, 2),
            profit=round(pro / 1e4, 2),
            margin=round(pro / amt, 4) if amt else 0.0,
            cat=[int((df["产品类别"] == k).sum()) for k in CAT_ORDER],
            region=[int((df["区域"] == k).sum()) for k in REGION_ORDER],
            express=[int((df["快递公司"] == k).sum()) for k in EXPRESS_ORDER],
            cost=cst_vals,
            share=[round(v / cst_sum, 4) for v in cst_vals],
            costSum=round(cst_sum, 2),
            top3=[[r.Index, round(float(r.金额) / 1e4, 2), int(r.数量)] for r in top.itertuples()],
        )
    return data, year


DATA, YEAR_SALES = load()
QTY_MAX = max(d["qty"] for d in DATA.values())
# 「销量环比」的横轴满量程：751 单 → 322px（绘图区 362px），即 844 单满格
HB_XMAX = max(844.0, QTY_MAX * 1.02)

# ==========================================================================
# 3. 卡片几何（px）：左/上/宽/高（由导出图的卡片缝隙实测）
# ==========================================================================
CARDS = {
    "pie9":   (0.0, Y(75.7), 525, 329),
    "huanbi": (X(516.8), Y(75.7), 525, 204),
    "area":   (X(815.6), Y(75.7), 526, 281),
    "top3":   (X(516.8), Y(195.4), 525, 255),
    "region": (X(815.6), Y(238.5), 526, 229),
    "kuaidi": (X(815.6), Y(373.1), 526, 229),
    "cost":   (0.0, Y(268.1), 525, 417),
    "profit": (X(516.8), Y(344.7), 525, 280),
    # 成本费用里的 2×2 小卡（推广/人工/产品/其他）
    "c1":     (X(225.7), Y(299.2), 234, 168),
    "c3":     (X(369.0), Y(300.0), 234, 168),
    "c4":     (X(225.8), Y(400.9), 234, 168),
    "c2":     (X(368.6), Y(401.6), 234, 168),
}

# 小环的圆心（卡内 px）与半径 —— 取自 chart1~4 的 manualLayout，并对导出图核对
GAUGE = {"c1": (70.4, 102.1), "c3": (65.0, 105.7), "c2": (63.2, 105.7), "c4": (70.4, 111.1)}
GAUGE_R = 40.7

# 四个小卡的文字锚点（卡内 px）：标题中心 x / 说明文字左端 x
GAUGE_TXT = {
    "c1": (70.0, 132.5, "推广"),
    "c3": (46.7, 133.0, "人工"),
    "c4": (71.4, 131.4, "产品"),
    "c2": (53.2, 122.7, "其他"),
}

# 区域/快递柱形：绘图区上沿、基线、值域上限
REGION_TOP, REGION_BASE, REGION_MAX = 100.6, 182.0, 300
KUAIDI_TOP, KUAIDI_BASE, KUAIDI_MAX = 73.6, 182.0, 400
BAR_W, BAR_PITCH = 24.0, 80.8

# 面积图：12 个点横向中心 75 + 37.727*i；值轴 0 在 232.3，200 在 70.3
AREA_X0, AREA_DX = 75.0, 37.727
AREA_Y0, AREA_Y200 = 232.3, 70.3

# 利润率环
PROFIT_C, PROFIT_R, PROFIT_HOLE = (153.0, 123.7), 90.0, 0.9222

# ==========================================================================
# 4. 字号（1pt = 1.79775px）
# ==========================================================================
FS_TITLE = 21.6       # 12pt 卡片标题
FS_8 = 14.4           # 8pt
FS_9 = 16.2           # 9pt
FS_10 = 18.0          # 10pt
FS_11 = 19.8          # 11pt
FS_12 = 21.6          # 12pt
FS_NUM_BIG = 44.0     # 饼图中心的销量数（宋体）
FS_PCT_BIG = 36.5     # 利润率大号百分比


def txt(sid, text, left, top, size, color=C_TEXT, anchor="center", weight=400,
        family="Microsoft YaHei", extra=""):
    """anchor: 'center' 以 left 为中心；'left' 左对齐；'right' 右对齐（均按垂直中心对齐）"""
    tr = {"center": "translate(-50%,-50%)", "left": "translate(0,-50%)",
          "right": "translate(-100%,-50%)"}[anchor]
    fam = "Arial, 'Microsoft YaHei'" if family == "latin" else family   # 西文优先，中文回落雅黑
    fam = fam if "," in fam else "'%s'" % fam
    return (f'<div class="t" id="{sid}" style="left:{left:.1f}px;top:{top:.1f}px;'
            f'font-size:{size:.1f}px;color:{color};font-weight:{weight};'
            f'font-family:{fam};transform:{tr};{extra}">{text}</div>')


# ==========================================================================
# 5. 文字层（DOM）—— 坐标均为卡片内 px，逐块对齐导出图
# ==========================================================================
def html_texts_pie9():
    h = [txt("pie9_title", "产品类别销量", 262, 32.4, FS_TITLE)]
    h.append(txt("pie9_total", "751", 182, 191.3, FS_NUM_BIG, family="SimSun"))
    h.append(txt("pie9_sub", "总销量", 178, 241.3, FS_10))
    h.append(txt("pie9_p1", "54%", 317, 220.3, FS_8))
    h.append(txt("pie9_p2", "21%", 72, 273.3, FS_8))
    h.append(txt("pie9_p3", "25%", 84, 103.3, FS_8))
    for i, (nm, col) in enumerate(zip(CAT_ORDER, CAT_COLORS)):
        cy = 143.5 + i * 55.5
        h.append(f'<div class="t" id="pie9_lg{i}" style="left:397px;top:{cy - 4:.1f}px;'
                 f'width:8px;height:8px;background:{col}"></div>')
        h.append(txt("pie9_lt%d" % i, nm, 410, cy, FS_8, anchor="left"))
    return "\n".join(h)


def html_texts_huanbi():
    h = [txt("hb_title", "销量环比", 262.2, 32.4, FS_TITLE)]
    h.append(txt("hb_m1", "9月", 34, 99.8, FS_8, anchor="right"))
    h.append(txt("hb_m2", "8月", 34, 152.3, FS_8, anchor="right"))
    h.append(txt("hb_v1", "751", 385, 101.0, FS_8, anchor="left"))
    h.append(txt("hb_v2", "714", 217, 153.0, FS_8, anchor="left"))
    h.append(txt("hb_pct", "5.18%", 411, 100.2, FS_10, weight=700, anchor="left"))
    h.append(txt("hb_yoy", "同比", 453.5, 149.3, FS_9))
    return "\n".join(h)


def html_texts_area():
    h = [txt("ar_title", "各月销售额(万)", 264, 32.5, FS_TITLE)]
    for i, v in enumerate(["200.00", "150.00", "100.00", "50.00", "0.00"]):
        h.append(txt("ar_y%d" % i, v, 60, AREA_Y200 + i * 40.5, FS_8, anchor="right"))
    for i, m in enumerate(MONTHS):
        h.append(txt("ar_x%d" % i, m, AREA_X0 + AREA_DX * i, 256.3, FS_9))
    return "\n".join(h)


TOP3_ROWS = (120.0, 167.5, 215.0)


def html_texts_top3():
    h = [txt("t3_title", "商品销售额Top3", 261.2, 33.2, FS_TITLE)]
    h.append(txt("t3_h1", "产品", 179.5, 81, FS_8))
    h.append(txt("t3_h2", "销售额(万)", 359, 81, FS_8))
    h.append(txt("t3_h3", "销量", 421, 81, FS_8))
    for i, y in enumerate(TOP3_ROWS):
        h.append(f'<img class="medal" id="t3_m{i}" src="" style="left:57px;top:{y - 18:.1f}px">')
        h.append(f'<div class="chip" id="t3_c{i}" style="left:327px;top:{y - 14.5:.1f}px;'
                 f'width:48px;height:29px"></div>')
        h.append(txt("t3_n%d" % i, "", 154, y - 1.5, FS_8, anchor="left", family="latin"))
        h.append(txt("t3_s%d" % i, "", 350.5, y, FS_8))
        h.append(txt("t3_q%d" % i, "", 419, y, FS_8))
    return "\n".join(h)


def _cat_texts(prefix, names, title, title_x, label_y):
    h = [txt("%s_title" % prefix, title, title_x, 32.5, FS_TITLE)]
    for i, nm in enumerate(names):
        cx = BAR_PITCH * (i + 0.5) + 21.0
        h.append(txt("%s_c%d" % (prefix, i), nm, cx, label_y, FS_9))
        h.append(txt("%s_v%d" % (prefix, i), "", cx, 120, FS_8))
    return h


def html_texts_region():
    return "\n".join(_cat_texts("rg", REGION_ORDER, "区域销量", 264, 205.5))


def html_texts_kuaidi():
    return "\n".join(_cat_texts("kd", EXPRESS_ORDER, "快递公司销量", 262, 204.5))


def html_texts_cost():
    return txt("cost_title", "成本费用", 262.5, 26.9, FS_TITLE)


def html_texts_gauge(key):
    tcx, lx, title = GAUGE_TXT[key]
    rcx = GAUGE[key][0]
    h = [txt("%s_title" % key, title, tcx, 27.0, FS_10)]
    h.append(txt("%s_l1" % key, title + "费", lx, 60.0, FS_8, anchor="left"))
    h.append(txt("%s_l2" % key, "用合计", lx, 83.5, FS_8, anchor="left"))
    h.append(txt("%s_pct" % key, "0.00%", rcx + 2, 103.5, FS_9))
    h.append(txt("%s_val" % key, "0.00", rcx + 59, 119, FS_10, anchor="left"))
    return "\n".join(h)


def html_texts_profit():
    h = [txt("pf_pct", "12.19%", 151, 112.0, FS_PCT_BIG)]
    h.append(txt("pf_sub", "利润率", 146.5, 148.7, 16.0))
    h.append(txt("pf_l1", "利润额(万)", 298, 97.2, FS_8, anchor="left"))
    h.append(txt("pf_l2", "销售额(万)", 298, 148.5, FS_8, anchor="left"))
    h.append(txt("pf_slash", "/", 290, 123, FS_9))
    h.append(txt("pf_v1", "16.37", 419, 97.2, FS_12, color=C_PINK, weight=700))
    h.append(txt("pf_v2", "134.29", 419, 148.5, FS_12, color=C_PINK, weight=700))
    h.append('<div class="pill" id="pf_pill"></div>')
    h.append(txt("pf_pilltxt", "利润额占比销售额", 260, 241.2, FS_12))
    return "\n".join(h)


# ==========================================================================
# 6. 组装 HTML
# ==========================================================================
def art_b64(path, thr=26):
    """原件里的奖牌/日历 PNG 是**纯黑底**，Excel 用「设置透明色」把黑抠掉。
    这里做同样的事：从四边泛洪找到与边相连的黑色区域 → alpha=0，
    奖牌内部被包住的深色（如数字「1」）保持不透明。"""
    a = np.asarray(Image.open(path).convert("RGB")).astype(int)
    dark = a.max(axis=2) <= thr
    lab, n = ndimage.label(dark)
    edge = set(lab[0].tolist()) | set(lab[-1].tolist())
    edge |= set(lab[:, 0].tolist()) | set(lab[:, -1].tolist())
    edge.discard(0)
    bg = np.isin(lab, list(edge)) if edge else np.zeros(dark.shape, bool)
    rgba = np.dstack([a.astype(np.uint8), np.where(bg, 0, 255).astype(np.uint8)])
    buf = io.BytesIO()
    Image.fromarray(rgba, "RGBA").save(buf, "PNG")
    return "data:image/png;base64," + base64.b64encode(buf.getvalue()).decode()


def build_html(plotly_js, medals, cal):
    def card(key, inner="", cls=""):
        l, t, w, h = CARDS[key]
        return (f'<div class="card{cls}" id="card_{key}" style="left:{l:.1f}px;top:{t:.1f}px;'
                f'width:{w}px;height:{h}px">{inner}</div>')

    def canvas(key):
        return f'<div class="cv" id="cv_{key}"></div>'

    cards = [card("pie9", canvas("pie9") + html_texts_pie9()),
             card("huanbi", canvas("huanbi") + html_texts_huanbi()),
             card("area", canvas("area") + html_texts_area()),
             card("top3", html_texts_top3()),
             card("region", canvas("region") + html_texts_region()),
             card("kuaidi", canvas("kuaidi") + html_texts_kuaidi()),
             card("cost", html_texts_cost())]
    for k in ("c1", "c3", "c4", "c2"):
        cards.append(card(k, canvas(k) + html_texts_gauge(k), cls=" box"))
    cards.append(card("profit", canvas("profit") + html_texts_profit()))

    geo = dict(
        cards=CARDS, gauge=GAUGE, gaugeR=GAUGE_R,
        region=[REGION_TOP, REGION_BASE, REGION_MAX],
        kuaidi=[KUAIDI_TOP, KUAIDI_BASE, KUAIDI_MAX],
        areaX=[AREA_X0, AREA_DX, 12],            # 首点 x / 步距 / 点数
        profit=[PROFIT_C[0], PROFIT_C[1], PROFIT_R, PROFIT_HOLE],
        size=dict(pie9=[525, 329], huanbi=[525, 204], area=[526, 281],
                  region=[526, 229], kuaidi=[526, 229], profit=[525, 280],
                  gauge=[234, 168]),
        barw=[BAR_W, BAR_PITCH], qtyMax=HB_XMAX,
    )

    html = TEMPLATE
    html = html.replace("__PLOTLY__", plotly_js)
    html = html.replace("__DATA__", json.dumps(DATA, ensure_ascii=False))
    html = html.replace("__YEAR__", json.dumps(YEAR_SALES))
    html = html.replace("__CARDS__", "\n".join(cards))
    html = html.replace("__GEO__", json.dumps(geo, ensure_ascii=False))
    html = html.replace("__MEDALS__", json.dumps(medals))
    html = html.replace("__CAL__", cal)
    html = html.replace("__COLORS__", json.dumps(dict(
        card=C_CARD, text=C_TEXT, teal=C_TEAL, pink=C_PINK, indigo=C_INDIGO,
        orange=C_ORANGE, track=C_TRACK, grad0=C_GRAD0, cat=CAT_COLORS)))
    html = html.replace("__MONTHS__", json.dumps(MONTHS, ensure_ascii=False))
    html = html.replace("__DEFMONTH__", DEFAULT_MONTH)
    return html


TEMPLATE = r"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>第四章 销售看板（plotly 复现）</title>
<script>__PLOTLY__</script>
<style>
  html,body{margin:0;padding:0;background:#fff;}
  #dash{position:relative;width:1600px;height:925px;background:#1A1E43;overflow:hidden;
        font-family:'Microsoft YaHei','Microsoft YaHei UI',sans-serif;
        -webkit-font-smoothing:antialiased;}
  .card{position:absolute;background:#2C3560;}
  .box{border:1px solid rgba(140,180,220,.55);border-radius:8px;box-sizing:border-box;}
  .cv{position:absolute;left:0;top:0;}
  .t{position:absolute;white-space:pre;line-height:1;pointer-events:none;
     font-feature-settings:'tnum';}
  #topband{position:absolute;left:3.4px;top:0;width:1596px;height:91.7px;background:#2C355F;}
  #botband{position:absolute;left:0;top:874.3px;width:1600px;height:50.9px;background:#2C355F;}
  #mainTitle{position:absolute;left:22px;top:42px;font-size:33.1px;font-weight:700;
             color:#F2F2F2;line-height:1;transform:translate(0,-50%);}
  #cal{position:absolute;left:1388px;top:46px;width:35px;height:35px;}
  #sel{position:absolute;left:1449px;top:48px;width:86px;height:31px;background:#fff;
       border:1px solid #A6A6A6;box-sizing:border-box;display:flex;align-items:center;
       cursor:pointer;user-select:none;}
  #selTxt{flex:1 1 auto;padding-left:7px;font-size:15px;color:#000;line-height:1;}
  #selArrow{flex:0 0 auto;width:20px;text-align:center;font-size:9px;color:#404040;}
  #selList{position:absolute;left:1449px;top:79px;width:86px;background:#fff;
           border:1px solid #A6A6A6;box-sizing:border-box;display:none;z-index:50;
           max-height:300px;overflow:auto;}
  #selList div{padding:4px 7px;font-size:15px;color:#000;cursor:pointer;line-height:1.2;}
  #selList div:hover{background:#DCE6F1;}
  .medal{position:absolute;width:36px;height:36px;pointer-events:none;}
  .chip{position:absolute;background:#4A5BD1;pointer-events:none;}
  .pill{position:absolute;left:151px;top:214.7px;width:207px;height:52px;
        border:1.5px solid #4C5FDC;border-radius:26px;box-sizing:border-box;}
</style>
</head>
<body>
<div id="dash">
  <div id="topband"></div>
  <div id="botband"></div>
  <div id="mainTitle">公司销售数据可视化看板</div>
  <img id="cal" src="__CAL__">
  <div id="sel"><div id="selTxt">9月</div><div id="selArrow">▼</div></div>
  <div id="selList"></div>
__CARDS__
</div>
<script>
var DATA = __DATA__, YEAR = __YEAR__, CO = __COLORS__, G = __GEO__, MONTHS = __MONTHS__;
var MEDALS = __MEDALS__;
var curMonth = "__DEFMONTH__";
var CFG = {displayModeBar:false, responsive:false, staticPlot:false, scrollZoom:false};

window.__err = "";
window.onerror = function (m, s, l, c) {
  window.__err += m + " @" + l + ":" + c + "\n";
  document.body.setAttribute("data-err", window.__err);
};
window.addEventListener("unhandledrejection", function (e) {
  window.__err += "REJ:" + ((e.reason && e.reason.message) || e.reason) + "\n";
  document.body.setAttribute("data-err", window.__err);
});

document.querySelectorAll('.medal').forEach(function (img, i) { img.src = MEDALS[i]; });

/* ---------------- 公共 ---------------- */
function baseLayout(w, h) {
  return {margin: {l: 0, r: 0, t: 0, b: 0}, width: w, height: h,
          paper_bgcolor: CO.card, plot_bgcolor: CO.card, showlegend: false,
          transition: {duration: 320, easing: 'cubic-in-out'},
          font: {family: 'Microsoft YaHei', size: 14, color: CO.text},
          hoverlabel: {bgcolor: '#11142E', bordercolor: CO.indigo,
                       font: {family: 'Microsoft YaHei', size: 13, color: CO.text}}};
}
var HIDEAX = {visible: true, showgrid: false, zeroline: false, showline: false,
              ticks: '', showticklabels: false, fixedrange: true};

function pieRing(values, colors, hole, cx, cy, r, w, h, hover) {
  var t = {type: 'pie', labels: ['', ' '], values: values, sort: false,
           direction: 'clockwise', rotation: 0, hole: hole,
           domain: {x: [(cx - r) / w, (cx + r) / w], y: [1 - (cy + r) / h, 1 - (cy - r) / h]},
           marker: {colors: colors, line: {width: 0}}, textinfo: 'none',
           hoverinfo: hover ? 'label+percent' : 'skip'};
  return t;
}

/* ---------------- 各图 ---------------- */
function fPie9(d) {
  return {data: [{type: 'pie', labels: ['办公用品', '家具产品', '数码电子'],
    values: d.cat, sort: false, direction: 'clockwise', rotation: 0, hole: 0.85,
    domain: {x: [0.1486, 0.5486], y: [0.0837, 0.7181]},
    marker: {colors: CO.cat, line: {width: 0}}, textinfo: 'none',
    hovertemplate: '%{label}<br>销量 %{value} 单<br>占比 %{percent}<extra></extra>'}],
    layout: baseLayout(525, 329)};
}
function fHuanbi(d) {
  var cur = d.qty, prev = d.prevQty;
  var xmax = G.qtyMax;                                    // 751 单 ≈ 322px（绘图区 362px）
  var L = baseLayout(525, 204);
  L.xaxis = Object.assign({}, HIDEAX, {range: [0, xmax], domain: [0.0971, 0.7867]});
  L.yaxis = Object.assign({}, HIDEAX, {domain: [0.1042, 0.6483],
                                       categoryorder: 'array', categoryarray: [d.prev, '本月']});
  L.bargap = 0.64;
  return {data: [{type: 'bar', orientation: 'h', x: [prev, cur], y: [d.prev, '本月'],
    marker: {color: [CO.indigo, CO.teal], line: {width: 0}},
    hovertemplate: '%{y} 销量 %{x:.0f} 单<extra></extra>'}], layout: L};
}
function fArea() {
  var L = baseLayout(526, 281);
  L.xaxis = Object.assign({}, HIDEAX, {domain: [0.1426, 0.9316],
                                       categoryorder: 'array', categoryarray: MONTHS});
  L.yaxis = Object.assign({}, HIDEAX, {range: [0, 200], domain: [0.1733, 0.7498]});
  return {data: [{type: 'scatter', mode: 'lines', x: MONTHS, y: YEAR,
    line: {width: 0, shape: 'linear'}, fill: 'tozeroy',
    fillgradient: {type: 'vertical', colorscale: [[0, CO.grad0], [1, CO.indigo]]},
    hovertemplate: '%{x}：%{y:.2f} 万<extra></extra>'}], layout: L};
}
function fBar(names, vals, color, top, base, ymax) {
  var L = baseLayout(526, 229);
  L.xaxis = Object.assign({}, HIDEAX, {domain: [21 / 526, 507 / 526],
                                       categoryorder: 'array', categoryarray: names});
  L.yaxis = Object.assign({}, HIDEAX, {range: [0, ymax], domain: [1 - base / 229, 1 - top / 229]});
  L.bargap = 1 - G.barw[0] / G.barw[1];
  return {data: [{type: 'bar', x: names, y: vals, marker: {color: color, line: {width: 0}},
    hovertemplate: '%{x}<br>销量 %{y} 单<extra></extra>'}], layout: L};
}
function fProfit(d) {
  var w = 525, h = 280, t = pieRing([d.margin, Math.max(1 - d.margin, 1e-9)],
      [CO.pink, CO.track], G.profit[3], G.profit[0], G.profit[1], G.profit[2], w, h, true);
  t.hovertemplate = '利润率 %{percent}<extra></extra>';
  return {data: [t], layout: baseLayout(w, h)};
}
function fGauge(k, val, total) {
  var cx = G.gauge[k][0], cy = G.gauge[k][1], r = G.gaugeR, w = 234, h = 168;
  return {data: [pieRing([val, Math.max(total - val, 1e-9)], [CO.teal, CO.track],
      0.85, cx, cy, r, w, h, false)], layout: baseLayout(w, h)};
}

/* ---------------- 渲染 ---------------- */
var drawn = false;
function render() {
  var d = DATA[curMonth];
  var jobs = [
    ['cv_pie9', fPie9(d)],
    ['cv_huanbi', fHuanbi(d)],
    ['cv_area', fArea()],
    ['cv_region', fBar(['华北', '华南', '东北', '西北', '西南', '华东'], d.region, CO.teal,
                       G.region[0], G.region[1], G.region[2])],
    ['cv_kuaidi', fBar(['顺丰', '韵达', '中通', '申通', '圆通', 'EMS'], d.express, CO.pink,
                       G.kuaidi[0], G.kuaidi[1], G.kuaidi[2])],
    ['cv_profit', fProfit(d)]
  ];
  var costSum = d.cost.reduce(function (a, b) { return a + b; }, 0);
  ['c1', 'c3', 'c4', 'c2'].forEach(function (k, i) {
    jobs.push(['cv_' + k, fGauge(k, d.cost[i], costSum)]);
  });
  jobs.forEach(function (j) {
    if (drawn) Plotly.react(j[0], j[1].data, j[1].layout, CFG);
    else Plotly.newPlot(j[0], j[1].data, j[1].layout, CFG);
  });
  drawn = true;
  updateTexts(d);
  document.body.setAttribute('data-ready', curMonth);
}

function setText(id, v) { var e = document.getElementById(id); if (e) e.textContent = v; }
function pct(v) { return (v * 100).toFixed(2) + '%'; }
function num(v) { return (Math.round(v * 100) / 100).toFixed(2); }

function updateTexts(d) {
  var sum = d.cat[0] + d.cat[1] + d.cat[2], acc = 0;
  setText('pie9_total', d.qty);
  /* 三个百分比标签跟随扇区中点（环心 183,198.3，标签半径 137），随数据联动 */
  for (var i = 0; i < 3; i++) {
    var share = d.cat[i] / sum;
    var mid = (acc + share / 2) * 2 * Math.PI;        // 自 12 点顺时针
    acc += share;
    var e = document.getElementById('pie9_p' + (i + 1));
    e.textContent = Math.round(share * 100) + '%';
    e.style.left = (183 + 137 * Math.sin(mid)).toFixed(1) + 'px';
    e.style.top = (198.3 - 137 * Math.cos(mid)).toFixed(1) + 'px';
  }

  setText('hb_m1', curMonth);
  setText('hb_m2', d.prev);
  var cur = d.qty, prev = d.prevQty;
  var pxPerUnit = 362 / G.qtyMax;                   // 绘图区宽 362px
  setText('hb_v1', cur);
  document.getElementById('hb_v1').style.left = (51 + cur * pxPerUnit + 12).toFixed(1) + 'px';
  setText('hb_v2', prev);
  document.getElementById('hb_v2').style.left = (51 + prev * pxPerUnit + 12).toFixed(1) + 'px';
  var p = document.getElementById('hb_pct');
  var dlt = (cur - prev) / prev;
  p.textContent = (dlt >= 0 ? '+' : '-') + Math.abs(dlt * 100).toFixed(2) + '%';
  p.style.left = (51 + cur * pxPerUnit + 39).toFixed(1) + 'px';

  for (var i = 0; i < 3; i++) {
    var t = d.top3[i] || ['—', 0, 0];
    setText('t3_n' + i, t[0]);
    setText('t3_s' + i, num(t[1]));
    setText('t3_q' + i, t[2]);
    document.getElementById('t3_c' + i).style.background = t[1] > 0 ? '#4A5BD1' : 'transparent';
  }
  var rTop = G.region[0], rBase = G.region[1], rMax = G.region[2];
  var kTop = G.kuaidi[0], kBase = G.kuaidi[1], kMax = G.kuaidi[2];
  for (var j = 0; j < 6; j++) {
    var y1 = rBase - (d.region[j] / rMax) * (rBase - rTop);
    setText('rg_v' + j, d.region[j]);
    document.getElementById('rg_v' + j).style.top = (y1 - 21) + 'px';
    var y2 = kBase - (d.express[j] / kMax) * (kBase - kTop);
    setText('kd_v' + j, d.express[j]);
    document.getElementById('kd_v' + j).style.top = (y2 - 21) + 'px';
  }

  setText('pf_pct', pct(d.margin));
  setText('pf_v1', num(d.profit));
  setText('pf_v2', num(d.amt));

  var keys = ['c1', 'c3', 'c4', 'c2'];
  keys.forEach(function (k, i) {
    setText(k + '_pct', pct(d.share[i]));
    setText(k + '_val', num(d.cost[i]));
  });
}

/* ---------------- 月份选择 ---------------- */
var selTxt = document.getElementById('selTxt'), selList = document.getElementById('selList');
MONTHS.forEach(function (m) {
  var e = document.createElement('div');
  e.textContent = m;
  e.onclick = function () { pick(m); };
  selList.appendChild(e);
});
function pick(m) {
  curMonth = m; selTxt.textContent = m; selList.style.display = 'none';
  render();
}
document.getElementById('sel').onclick = function (e) {
  selList.style.display = (selList.style.display === 'block') ? 'none' : 'block';
  e.stopPropagation();
};
document.body.addEventListener('click', function () { selList.style.display = 'none'; });
document.addEventListener('keydown', function (e) {
  var i = MONTHS.indexOf(curMonth);
  if (e.key === 'ArrowDown') pick(MONTHS[(i + 1) % 12]);
  else if (e.key === 'ArrowUp') pick(MONTHS[(i + 11) % 12]);
  else if (/^[1-9]$/.test(e.key)) pick(MONTHS[+e.key - 1]);
  else if (e.key === '0') pick(MONTHS[9]);
});

/* ---------------- 启动 ---------------- */
pick("__DEFMONTH__");
</script>
</body>
</html>
"""


# ==========================================================================
# 7. 截图
# ==========================================================================
def save_png(html_path, png_path):
    prof = tempfile.mkdtemp()
    subprocess.run([
        CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--disable-lcd-text",
        "--force-device-scale-factor=1", "--user-data-dir=" + prof,
        "--window-size=%d,%d" % (W, H),
        "--run-all-compositor-stages-before-draw", "--virtual-time-budget=6000",
        "--screenshot=" + png_path, "file:///" + quote(html_path.replace("\\", "/")),
    ], capture_output=True, timeout=300)
    import shutil
    shutil.rmtree(prof, ignore_errors=True)


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wd = os.path.dirname(os.path.abspath(__file__))
    medals = [art_b64(os.path.join(wd, n)) for n in ("徽章_金.png", "徽章_银.png", "徽章_铜.png")]
    cal = art_b64(os.path.join(wd, "日历.png"))
    from plotly.offline import get_plotlyjs
    html = build_html(get_plotlyjs(), medals, cal)
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print("HTML", HTML_PATH, round(os.path.getsize(HTML_PATH) / 1e6, 2), "MB")
    save_png(HTML_PATH, PNG_PATH)
    print("PNG ", PNG_PATH)
    d = DATA[DEFAULT_MONTH]
    print("校验 %s：销量 %d / 销售额 %.4f 万 / 利润额 %.4f 万 / 利润率 %.6f"
          % (DEFAULT_MONTH, d["qty"], d["amt"], d["profit"], d["margin"]))
    print("       成本 %s 合计 %.4f 占比 %s" % (d["cost"], d["costSum"], d["share"]))
    print("       Top3 %s" % (d["top3"],))


if __name__ == "__main__":
    main()
