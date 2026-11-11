# -*- coding: utf-8 -*-
"""
嘉兴大学 2026 级新生人数 · 旭日图（pyecharts 复现版）
================================================================
与参考图完全同款：pyecharts Sunburst 默认样式，两层「学院 → 专业」，
点击学院下钻、tooltip 显示 人数 / 占本层 / 占总体。

运行方式
    1) 直接运行本文件（Spyder 里 F5、或命令行 python sunburst_repro.py）
       → 生成 旭日图_复现.html，自动用浏览器打开
    2) 想和参考图一样在 JupyterLab 里交互 → 打开同目录的 旭日图.ipynb

数据
    data/新生人数.csv（列：学院,专业,人数），改完数字直接重跑即可
"""

import csv
import re
import webbrowser
from pathlib import Path

from pyecharts import options as opts
from pyecharts.charts import Sunburst
from pyecharts.commons.utils import JsCode

# ---------------------------------------------------------------- 路径
BASE = Path(__file__).resolve().parent
CSV_PATH = BASE / "data" / "新生人数.csv"
HTML_PATH = BASE / "旭日图_复现.html"
ECHARTS_LIB = BASE.parent / "assets" / "echarts.min.js"   # 内联用，缺省则走 CDN

ROOT_NAME = "2026级总览"

# 每个学院一个色系，专业用同色系浅色（与参考图的观感一致）
PALETTE = [
    "#5470c6", "#91cc75", "#fac858", "#ee6666", "#73c0de",
    "#3ba272", "#fc8452", "#9a60b4", "#ea7ccc", "#b6a2de",
    "#5ab1ef", "#ffb980", "#d87a80", "#8d98b3", "#e5cf0d",
]


# ---------------------------------------------------------------- 数据
def load_data(path=CSV_PATH):
    """读 CSV → [(学院, [(专业, 人数), ...]), ...]，按学院人数降序。"""
    buckets = {}
    with open(path, encoding="utf-8-sig") as f:
        for row in csv.DictReader(f):
            college = row["学院"].strip()
            major = row["专业"].strip()
            num = int(row["人数"])
            buckets.setdefault(college, []).append((major, num))

    data = [(c, sorted(v, key=lambda x: -x[1])) for c, v in buckets.items()]
    data.sort(key=lambda x: -sum(n for _, n in x[1]))
    return data


def mix(color: str, other: str, t: float) -> str:
    """把 color 往 other 混 t（0~1），用来生成同色系的浅色。"""
    a = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(other[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


def build_tree(data, root_name=ROOT_NAME):
    """拼成 ECharts sunburst 的树：根节点 = 全校，第一层 = 学院，第二层 = 专业。"""
    total = sum(n for _, majors in data for _, n in majors)
    children = []
    for i, (college, majors) in enumerate(data):
        base = PALETTE[i % len(PALETTE)]
        kids = [{
            "name": major,
            "value": num,
            "itemStyle": {"color": mix(base, "#ffffff", 0.62 + 0.05 * (j % 3))},
        } for j, (major, num) in enumerate(majors)]

        children.append({
            "name": college,
            "id": college,                       # 下钻定位用
            # 「机械工程学院（机器人工程学院）」这类带括号的长名，图上只写主名
            "short": re.sub(r"[（(].*?[)）]", "", college),
            "value": sum(n for _, n in majors),
            "itemStyle": {"color": base},
            "children": kids,
        })

    root = {
        "name": root_name,
        "value": total,
        "itemStyle": {"color": "#ffffff"},   # 中心留白，文字压在圆心
        "label": {"rotate": 0},              # 圆心文字必须水平
        "children": children,
    }
    return [root], total


# ---------------------------------------------------------------- JS
def label_js(total):
    """标签：名称一行、人数一行；太小的扇区不标，避免压字。"""
    return """function (p) {
    var share = p.value / %d;
    if (share < 0.004) { return ''; }
    var nm = (p.data && p.data.short) ? p.data.short : p.name;
    return nm + '\\n' + p.value;
}""" % total


def tooltip_js(total):
    """与参考图字段完全一致：人数 / 占本层 / 占总体（原始比值，不取整）。"""
    return """function (p) {
    var path = p.treePathInfo || [];
    var parent = path.length > 1 ? path[path.length - 2].value : p.value;
    return p.name
        + '<br/>人数: ' + p.value
        + '<br/>占本层: ' + (p.value / parent)
        + '<br/>占总体: ' + (p.value / %d);
}""" % total


# ---------------------------------------------------------------- 画图
def drill_js(cid):
    """点击下钻：用被点的学院重画一棵两层树；点圆心回到总览。"""
    return """
setTimeout(function () {
    var OPT = option_%s;
    var ROOT = JSON.parse(JSON.stringify(OPT.series[0].data[0]));
    var chart = chart_%s;
    var current = ROOT;

    function render(node) {
        var opt = {}, series = {};
        for (var k in OPT) { opt[k] = OPT[k]; }
        for (var j in OPT.series[0]) { series[j] = OPT.series[0][j]; }
        series.data = [node];
        opt.series = [series];
        chart.setOption(opt, true);
    }

    chart.on('click', function (p) {
        var id = p.data && p.data.id;
        if (!id) { return; }
        if (id === current.id) {              /* 点圆心：退回总览 */
            current = ROOT;
            render(ROOT);
            return;
        }
        for (var i = 0; i < ROOT.children.length; i++) {
            if (ROOT.children[i].id === id) {  /* 点学院：下钻 */
                current = ROOT.children[i];
                var node = JSON.parse(JSON.stringify(current));
                node.label = {rotate: 0};
                render(node);
                return;
            }
        }
    });
}, 0);
""" % (cid, cid)


def make_chart(data=None, root_name=ROOT_NAME):
    data = data or load_data()
    tree, total = build_tree(data, root_name)

    chart = Sunburst(
        init_opts=opts.InitOpts(
            width="1040px",
            height="940px",
            bg_color="#ffffff",
            page_title="嘉兴大学2026级新生人数旭日图",
        )
    )
    chart.add(
        series_name="",
        data_pair=tree,
        radius=["0%", "90%"],
        center=["50%", "50%"],
        sort_="desc",
        is_clockwise=True,
        start_angle=90,
        label_opts=opts.LabelOpts(
            is_show=True,
            rotate="radial",
            color="#333333",
            font_size=10.5,
            overflow="truncate",
            formatter=JsCode(label_js(total)),
        ),
        itemstyle_opts=opts.ItemStyleOpts(
            border_color="#ffffff",
            border_width=1.5,
        ),
    )

    # 标签宽度（超过则截断），LabelOpts 没有 width 字段，用 update 补上
    chart.options["series"][0]["label"].update(width=110, overflow="truncate")
    chart.options["animation"] = False

    # 点击下钻：以被点学院为根重画一棵两层树。
    # （不用内置的 nodeClick='rootToNode'：它会把「全校」那层也留在圆心，
    #   下钻后变成三层，而参考图是两层、学院名直接在圆心。）
    chart.add_js_funcs(drill_js(chart.chart_id))

    chart.set_global_opts(
        title_opts=opts.TitleOpts(is_show=False),
        legend_opts=opts.LegendOpts(is_show=False),
        tooltip_opts=opts.TooltipOpts(
            trigger="item",
            background_color="rgba(255,255,255,0.96)",
            border_color="#dddddd",
            border_width=1,
            textstyle_opts=opts.TextStyleOpts(color="#333333", font_size=12),
            formatter=JsCode(tooltip_js(total)),
        ),
    )
    return chart, total


def inline_echarts(html_path: Path) -> bool:
    """把 echarts.min.js 内联进 HTML，断网也能双击打开。"""
    if not ECHARTS_LIB.exists():
        return False
    lib = ECHARTS_LIB.read_text(encoding="utf-8", errors="ignore")
    if "</script" in lib:            # 内联会截断 HTML，放弃
        return False

    html = html_path.read_text(encoding="utf-8")
    pattern = re.compile(r'<script[^>]*src="[^"]*echarts[^"]*"[^>]*>\s*</script>')
    if not pattern.search(html):
        return False
    html_path.write_text(
        pattern.sub(lambda _m: "<script>%s</script>" % lib, html, count=1),
        encoding="utf-8",
    )
    return True


# ---------------------------------------------------------------- 主流程
def _pad(text, width):
    """按显示宽度补空格（中文占两格），让控制台表格对齐。"""
    w = sum(2 if ord(ch) > 127 else 1 for ch in text)
    return text + " " * max(0, width - w)


def main():
    data = load_data()
    chart, total = make_chart(data)

    chart.render(str(HTML_PATH))
    offline = inline_echarts(HTML_PATH)

    # 控制台核对表：每院小计 + 占比
    print("=" * 50)
    print(_pad("学院", 30) + "%6s %8s %5s" % ("人数", "占全校", "专业数"))
    print("-" * 50)
    for college, majors in data:
        sub = sum(n for _, n in majors)
        print(_pad(college, 30) + "%6d %7.2f%% %5d"
              % (sub, sub / total * 100, len(majors)))
    print("-" * 50)
    print(_pad("合计", 30) + "%6d %7.2f%% %5d" % (total, 100.0, len(data)))
    print("=" * 50)
    print("HTML:", HTML_PATH)
    print("ECharts 已内联：%s" % ("是（可离线）" if offline else "否（走 CDN）"))

    try:
        webbrowser.open(HTML_PATH.as_uri())
    except Exception:
        pass
    return chart


if __name__ == "__main__":
    main()
