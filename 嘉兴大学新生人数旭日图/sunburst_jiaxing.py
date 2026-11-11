# -*- coding: utf-8 -*-
"""
嘉兴大学 2026 级新生规模 · 交互式旭日图（Sunburst）生成脚本
============================================================
数据：data/新生人数.csv（学院, 专业, 人数；UTF-8-BOM，可直接用 Excel 改）
输出：嘉兴大学新生人数旭日图.html（单文件，已内联 ECharts，断网双击可用）

用法：
    python sunburst_jiaxing.py

设计要点：
1. pyecharts Sunburst 负责图形结构，option 导出后注入 HTML 模板；
2. label / tooltip 用 JsCode 自定义：显示“人数 / 占本层% / 占总体%”，
   占总体固定以全校口径计算（下钻后依然显示占全校比例，与参考图一致）；
3. nodeClick='rootToNode'：点击学院扇区即下钻为该学院的根节点；
4. 扇区颜色按学院分色系，专业用同色系浅一档，视觉上从属关系清晰。
"""

from __future__ import annotations

import csv
import json
import re
from collections import OrderedDict
from pathlib import Path

from pyecharts import options as opts
from pyecharts.charts import Sunburst
from pyecharts.commons.utils import JsCode

BASE = Path(__file__).resolve().parent
CSV_PATH = BASE / "data" / "新生人数.csv"
TPL_PATH = BASE / "templates" / "page.html"
ASSETS_PATH = BASE / "assets" / "echarts.min.js"
OUT_PATH = BASE / "嘉兴大学新生人数旭日图.html"

ROOT_NAME = "2026级总览"

# 图上放的短名（官方全名仍在 tooltip 中显示，避免长名超出环宽）
COLLEGE_ALIAS = {
    "机械工程学院（机器人工程学院）": "机械工程",
    "信息科学与工程学院": "信息科学",
    "生物与化学工程学院": "生物与化学",
    "材料与纺织工程学院": "材料与纺织",
    "马克思主义学院（中共党史党建学院）": "马克思主义",
    "平湖师范学院（教育学院）": "平湖师范",
}
MAJOR_ALIAS = {
    "机械设计制造及其自动化": "机械设计制造",
    "材料成型及控制工程": "材料成型",
    "建筑环境与能源应用工程": "建筑环境",
    "数据科学与大数据技术": "数据科学(大)",
    "信息管理与信息系统": "信息管理",
    "新能源材料与器件": "新能源材料",
    "非织造材料与工程": "非织造材料",
    "计算机科学与技术": "计算机科学",
    "国际经济与贸易": "国际经济贸易",
    "服装设计与工程": "服装设计工程",
    "视觉传达设计": "视觉传达",
    "电气工程及其自动化": "电气工程",
    "智能制造工程": "智能制造",
    "机器人工程": "机器人工程",
    "跨境电子商务": "跨境电商",
    "化学工程与工艺": "化学工程",
    "高分子材料与工程": "高分子材料",
    "数据科学与大数据": "数据科学(大)",
}

# 15 套学院色系（柔和但可区分，浅色底上文字用深色）
PALETTE = [
    "#5C86C4", "#4FA089", "#E0A657", "#9A85C6", "#D98A83",
    "#7FA96B", "#CDB35F", "#63A0B4", "#C08FAE", "#5FA79E",
    "#D2926A", "#9AA763", "#8B84C6", "#C39A9A", "#7B9AA6",
]


# --------------------------------------------------------------------------
# 1. 读数与校验
# --------------------------------------------------------------------------
def load_rows(path: Path) -> "OrderedDict[str, list]":
    """读取 CSV，保持学院首次出现顺序。"""
    groups: "OrderedDict[str, list]" = OrderedDict()
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            college = (row.get("学院") or "").strip()
            major = (row.get("专业") or "").strip()
            if not college or not major:
                continue
            groups.setdefault(college, []).append((major, int(str(row["人数"]).strip())))
    return groups


def validate(groups) -> dict:
    """校验并打印核对表：每院小计 = 专业之和；全校合计 = 4638。"""
    total = sum(n for items in groups.values() for _, n in items)
    n_major = sum(len(v) for v in groups.values())
    rows = sorted(((c, sum(n for _, n in v)) for c, v in groups.items()), key=lambda x: -x[1])

    print(f"{'学院':<24}{'人数':>6}{'占比':>8}   专业构成")
    print("-" * 96)
    for c, s in rows:
        detail = "  ".join(f"{m}{n}" for m, n in groups[c])
        print(f"{c:<24}{s:>6}{s / total * 100:>7.2f}%   {detail}")
    print("-" * 96)
    print(f"{'合计':<24}{total:>6}{100.0:>7.2f}%   {len(groups)} 个学院 / {n_major} 个专业")
    assert total == 4638, f"总人数 {total} != 4638，请检查数据"
    return {"total": total, "colleges": len(groups), "majors": n_major, "rows": rows}


# --------------------------------------------------------------------------
# 2. 构建旭日图数据树
# --------------------------------------------------------------------------
def _mix(hex_color: str, other: str = "#ffffff", t: float = 0.4) -> str:
    """把颜色向白色混合，t 越大越浅。"""
    a = [int(hex_color[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(other[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


def build_tree(groups):
    colleges = sorted(groups.items(), key=lambda kv: -sum(n for _, n in kv[1]))
    total = sum(n for _, items in colleges for _, n in items)
    children = []
    for i, (college, items) in enumerate(colleges):
        base = PALETTE[i % len(PALETTE)]
        majors = sorted(items, key=lambda x: -x[1])
        kids = []
        for j, (major, n) in enumerate(majors):
            kids.append({
                "name": major,
                "displayName": MAJOR_ALIAS.get(major, major),
                "value": n,
                "id": f"{college}/{major}",
                "itemStyle": {"color": _mix(base, "#ffffff", 0.30 + 0.07 * (j % 3))},
            })
        sub = sum(n for _, n in majors)
        short = COLLEGE_ALIAS.get(college) or re.sub(r"（.*?）", "", college)
        children.append({
            "name": college,
            "displayName": short,
            "value": sub,
            "id": college,
            "itemStyle": {"color": base, "borderColor": "#ffffff", "borderWidth": 2.5},
            "children": kids,
            "share": round(sub / total * 100, 2),
        })
    root = {
        "name": ROOT_NAME,
        "value": total,
        "id": "root",
        "itemStyle": {"color": "#ffffff", "borderColor": "#e8eef7", "borderWidth": 1},
        "children": children,
    }
    return root


# --------------------------------------------------------------------------
# 3. 标签 / 提示框交互（JS）
# --------------------------------------------------------------------------
LABEL_JS = """function (p) {
    /* treePathInfo 第 0 项是 series 自身，去掉后才是真实层级 */
    var path = p.treePathInfo || [];
    var d = Math.max(0, path.length - 1);
    var share = p.value / TOTAL;
    var nm = (p.data && p.data.displayName) ? p.data.displayName : p.name;
    if (d <= 1) { return ''; }
    if (d === 2) {
        if (share < 0.012) { return ''; }
        return '{c|' + nm + '}\\n{v|' + p.value + ' 人}';
    }
    if (share < 0.008) { return ''; }
    return '{m|' + nm + '  ' + p.value + '}';
}"""

TOOLTIP_JS = """function (p) {
    var path = p.treePathInfo || [];
    var parent = path.length > 2 ? path[path.length - 2].value : TOTAL;
    var pct = function (a, b) { return b ? (a / b * 100).toFixed(2) : '0.00'; };
    var hasKids = !!(p.data && p.data.children && p.data.children.length);
    var kind = hasKids ? (p.name === '2026级总览' ? '全校' : '学院') : '专业';
    var rows = [['人数', p.value + ' 人'], ['占本层', pct(p.value, parent) + '%'], ['占总体', pct(p.value, TOTAL) + '%']];
    var html = '<div class="tt-h">' + p.name + '<span class="tt-k">' + kind + '</span></div>';
    for (var i = 0; i < rows.length; i++) {
        html += '<div class="tt-r"><span>' + rows[i][0] + '</span><b>' + rows[i][1] + '</b></div>';
    }
    html += '<div class="tt-f">' + (hasKids && kind === '学院' ? '点击查看该学院专业构成' : '') + '</div>';
    return html;
}"""

LABEL_RICH = {
    "c": {"fontSize": 11, "fontWeight": 600, "color": "#1f3049", "lineHeight": 15},
    "v": {"fontSize": 10, "color": "#4a5872", "lineHeight": 14},
    "m": {"fontSize": 9.5, "color": "#33445f"},
}


# --------------------------------------------------------------------------
# 4. 生成图表配置
# --------------------------------------------------------------------------
def build_chart(root: dict) -> str:
    chart = Sunburst(
        init_opts=opts.InitOpts(
            width="100%", height="100%", bg_color="#ffffff",
            page_title="嘉兴大学 2026 级新生规模分布 · 旭日图",
        )
    )
    chart.add(
        series_name="2026级新生",
        data_pair=[root],
        radius=["18%", "93%"],
        center=["50%", "50%"],
        node_click="rootToNode",
        highlight_policy="ancestor",
        sort_="desc",
        is_clockwise=True,
        start_angle=90,
        label_opts={
            "show": True,
            "color": "#33445f",
            "fontSize": 9.5,
            "width": 140,
            "overflow": "truncate",
            "ellipsis": "…",
            "formatter": JsCode(LABEL_JS),
            "rich": LABEL_RICH,
        },
        label_layout_opts={"hideOverlap": True},
        itemstyle_opts={"borderColor": "#ffffff", "borderWidth": 2},
        emphasis_opts={
            "focus": "ancestor",
            "itemStyle": {"borderColor": "#2f6fed", "borderWidth": 2.5},
            "label": {"show": True, "fontWeight": 700},
        },
        # 注意：ECharts 的 levels[i] 实际作用于「第 i-1 层」，
        # 因此 levels[0] 留空，1=根节点（中心）、2=学院环、3=专业环
        levels=[
            {},
            {"r0": "0%", "r": "18%",
             "label": {"rotate": 0, "position": "inside", "width": 110},
             "itemStyle": {"color": "#ffffff", "borderColor": "#e8eef7", "borderWidth": 1}},
            {"r0": "18%", "r": "50%",
             "label": {"rotate": "radial", "fontSize": 10.5, "width": 96},
             "itemStyle": {"borderColor": "#ffffff", "borderWidth": 2.5}},
            {"r0": "50%", "r": "93%",
             "label": {"rotate": "radial", "fontSize": 9.5, "width": 122},
             "itemStyle": {"borderColor": "#ffffff", "borderWidth": 1.6}},
        ],
    )

    series = chart.options["series"][0]
    series["blur"] = {"itemStyle": {"opacity": 0.28}, "label": {"opacity": 0.35}}
    series["breadcrumb"] = {
        "show": False,
        "left": "center",
        "bottom": 6,
        "height": 24,
        "itemStyle": {"color": "#f2f6fd", "borderColor": "#e6edf7", "borderWidth": 1,
                      "textStyle": {"color": "#5a6a85", "fontSize": 12}},
        "emphasis": {"itemStyle": {"color": "#e2ecff", "borderColor": "#cfe0ff",
                                   "textStyle": {"color": "#2f6fed", "fontSize": 12}}},
    }
    series["animationDurationUpdate"] = 620
    series["animationEasingUpdate"] = "cubicOut"

    chart.set_global_opts(
        tooltip_opts=opts.TooltipOpts(
            is_show=True,
            trigger="item",
            formatter=JsCode(TOOLTIP_JS),
            background_color="#ffffff",
            border_color="#e6edf7",
            border_width=1,
            padding=[10, 14],
            textstyle_opts=opts.TextStyleOpts(color="#16233a", font_size=12),
            extra_css_text=(
                ".tt-h{font-size:13px;font-weight:700;margin-bottom:6px;color:#16233a;"
                "display:flex;align-items:center;gap:8px}"
                ".tt-k{font-size:10.5px;font-weight:500;color:#2f6fed;background:#eef4ff;"
                "border-radius:4px;padding:1px 6px}"
                ".tt-r{display:flex;justify-content:space-between;gap:22px;line-height:21px;color:#6b7a93}"
                ".tt-r b{color:#16233a;font-weight:700}"
                ".tt-f{margin-top:6px;font-size:11px;color:#a3aec2}"
            ),
        )
    )
    return chart.dump_options()


# --------------------------------------------------------------------------
# 5. 注入 HTML 模板
# --------------------------------------------------------------------------
def _rail_items(meta: dict, top_n: int = 15) -> str:
    """右侧「学院规模 TOP N」排行条。"""
    total = meta["total"]
    ranked = sorted(
        ((c, sum(n for _, n in items)) for c, items in meta["groups"].items()),
        key=lambda x: -x[1],
    )[:top_n]
    peak = ranked[0][1]
    out = []
    for i, (college, n) in enumerate(ranked):
        short = COLLEGE_ALIAS.get(college) or re.sub(r"（.*?）", "", college)
        color = PALETTE[i % len(PALETTE)]
        out.append(
            f'<div class="rank-item">'
            f'<span class="nm"><i class="dot" style="background:{color}"></i>{short}</span>'
            f'<span class="vl">{n}<i>{n / total * 100:.1f}%</i></span>'
            f'<span class="bar"><i style="width:{n / peak * 100:.1f}%;background:{color}"></i></span>'
            f'</div>'
        )
    return "".join(out)


def render(option_json: str, meta: dict) -> str:
    html = TPL_PATH.read_text(encoding="utf-8")

    if ASSETS_PATH.exists():
        js = ASSETS_PATH.read_text(encoding="utf-8", errors="ignore")
        html = html.replace("<script>__ECHARTS_JS__</script>", f"<script>{js}</script>")
    else:  # 没带 assets 时退回 CDN（需要联网）
        html = html.replace(
            "<script>__ECHARTS_JS__</script>",
            '<script src="https://assets.pyecharts.org/assets/v5/echarts.min.js"></script>',
        )

    total = meta["total"]
    top_college, top_college_n = meta["rows"][0]
    top_major_name, top_major_n = max(
        ((m, n) for items in meta["groups"].values() for m, n in items), key=lambda x: x[1]
    )

    repl = {
        "__OPTION_JSON__": option_json,
        "__RAIL_ITEMS__": _rail_items(meta),
        "__TOTAL__": str(total),
        "__TOTAL_TEXT__": f"{total:,}",
        "__N_COLLEGE__": str(meta["colleges"]),
        "__N_MAJOR__": str(meta["majors"]),
        "__TOP_COLLEGE_NAME__": top_college,
        "__TOP_COLLEGE_N__": str(top_college_n),
        "__TOP_COLLEGE_PCT__": f"{top_college_n / total * 100:.2f}",
        "__TOP_MAJOR_NAME__": top_major_name,
        "__TOP_MAJOR_N__": str(top_major_n),
        "__TOP_MAJOR_PCT__": f"{top_major_n / total * 100:.2f}",
    }
    for k, v in repl.items():
        html = html.replace(k, v)
    return html


def main() -> None:
    groups = load_rows(CSV_PATH)
    meta = validate(groups)
    meta["groups"] = groups
    root = build_tree(groups)
    option_json = build_chart(root)
    html = render(option_json, meta)
    OUT_PATH.write_text(html, encoding="utf-8")
    size = OUT_PATH.stat().st_size / 1024
    print(f"\n已生成：{OUT_PATH}  ({size:.0f} KB，已内联 ECharts，可离线打开)")


if __name__ == "__main__":
    main()
