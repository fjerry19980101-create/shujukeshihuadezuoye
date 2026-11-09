# -*- coding: utf-8 -*-
"""
第三章 动态图表 —— 3 动态南丁格尔圆环图（可交互复刻）
====================================================================

复刻目标（截图 944×716）
    深藏蓝底 + 白色粗体标题 + 六层同心环带构成的南丁格尔玫瑰/年轮图，
    右上角两枚灰色单选按钮（近7天 / 近30天）用于切换数据。

交互方式（弹出的窗口里）
    · 鼠标点击右上角「近7天 / 近30天」按钮 → 环形与标签以补间动画平滑过渡
    · 键盘 ← → 或数字键 1 / 2 → 同样切换
    · 鼠标移到某个扇区上 → 该渠道高亮、标签变纯白
    · 关闭窗口即退出

输出（保存在 图/第三章/）
    3 动态南丁格尔圆环图_近7天.png
    3 动态南丁格尔圆环图_近30天.png
    3 动态南丁格尔圆环图.gif      ← 两个状态来回变形的动图

说明
    · 数据直接读同目录的 xlsx（第 3~5 行 B:G 两组占比），改 Excel 再运行即可更新；
      Excel 缺失或读不到时，回退到脚本内置的 DATA。
    · 图形直接使用 944×716 像素坐标系（原点左上、y 向下），所有坐标量自截图，
      改动时按像素微调即可。
    · 环带"高度"是设计固定的六级台阶（6/5/4/3/2 层），随数据变化的只有圆心角。
    · 若运行环境拿不到图形界面（如 Spyder 的 Inline 绘图），脚本仍会正常出图，
      只是不弹交互窗口；此时可在 IPython 控制台先执行 `%matplotlib qt` 再运行，
      命令行运行 `python "3 动态南丁格尔圆环图.py" --no-gui` 则强制只出图。
"""

import os
import sys

import numpy as np

# --------------------------------------------------------------------------
# 0. 选择后端：优先 Qt / Tk 这类可交互后端，失败再退回默认（通常为 Agg）
# --------------------------------------------------------------------------
import matplotlib

SHOW_INTERACTIVE = True         # 改成 False 就只出图、不弹交互窗口
_FORCE_NO_GUI = ("--no-gui" in sys.argv) or ("-n" in sys.argv)
if SHOW_INTERACTIVE and not _FORCE_NO_GUI:
    for _name in ("QtAgg", "Qt5Agg", "TkAgg"):
        try:
            matplotlib.use(_name)
            break
        except Exception:
            continue

import matplotlib.pyplot as plt
from matplotlib.animation import FuncAnimation, PillowWriter
from matplotlib.font_manager import FontProperties
from matplotlib.patches import Circle, Rectangle, Wedge

plt.rcParams["axes.unicode_minus"] = False


def _backend_is_gui():
    """当前后端是否真的能弹窗交互。"""
    name = matplotlib.get_backend().lower()
    return not any(k in name for k in ("agg", "inline", "pdf", "svg", "ps", "cairo", "template"))


# --------------------------------------------------------------------------
# 1. 全局样式参数
# --------------------------------------------------------------------------
W, H = 944, 716                 # 画布尺寸（像素）
DPI = 100                       # 100 dpi → 1 英寸 = 100 像素

BG = "#1A1E43"                  # 背景深藏蓝
C_TITLE = "#FFFFFF"             # 标题白
C_LABEL = "#F2F2F2"             # 数据标签（略灰于纯白）
C_NOTE = "#D9D9D9"              # 脚注浅灰
C_BTN = "#8E8F9A"               # 按钮底色
C_RADIO_RING = "#646464"        # 单选圆点的灰色外环
C_TEXT_DARK = "#000000"         # 按钮文字 / 选中圆点

# 环形几何（量自截图）
CX, CY = 477.0, 393.0           # 圆心
R_IN = 109.0                    # 内半径（中间挖空）
R_OUT = 210.0                   # 外半径
RINGS = 6                       # 同心环带层数
RING_W = (R_OUT - R_IN) / RINGS  # 单层宽度 ≈ 16.83 px

DOT_R = 5.0                     # 引线端点圆点半径
LINE_W = 3.0 * 72 / DPI         # 引线 3px（对齐半像素后可得到硬边）

# 字号：pt = px * 72 / DPI
FS_TITLE = 47.8 * 72 / DPI
FS_LABEL = 24.3 * 72 / DPI
FS_NOTE = 22.0 * 72 / DPI
FS_BTN = 24.0 * 72 / DPI

# 标题 / 脚注锚点（左端基线）
TITLE_XY = (62.0, 85.5)
NOTE_XY = (53.0, 672.0)

# --------------------------------------------------------------------------
# 2. 字体：标题=微软雅黑 Bold，正文=等线（与截图的字宽最吻合）
# --------------------------------------------------------------------------
_FONT_DIR = r"C:\Windows\Fonts"


def _load_font(candidates, fallback="Microsoft YaHei"):
    for name in candidates:
        path = os.path.join(_FONT_DIR, name)
        if os.path.exists(path):
            try:
                return FontProperties(fname=path)
            except Exception:
                continue
    return FontProperties(family=fallback)


FP_TITLE = _load_font(["msyhbd.ttc", "simhei.ttf"])
FP_TEXT = _load_font(["Deng.ttf", "msyhl.ttc", "msyh.ttc"])

# --------------------------------------------------------------------------
# 3. 数据与配色
# --------------------------------------------------------------------------
TITLE_TEXT = "2022年6月30日流量来源分布"
NOTE_TEXT = "*注: 数据来源公司网站"

# 五个渠道：颜色 + 环带层数（个人主页内层黄、外层绿，是三层的双色处理）
#   line_color：引线与端点圆点的颜色，默认同 color；个人主页的引线在原图里是黄色
SEGMENTS = [
    dict(name="首页推荐", color="#4A5BD1", rings=6),
    dict(name="关注页面", color="#6073F3", rings=5),
    dict(name="搜索", color="#11A7AD", rings=4),
    dict(name="个人主页", color="#75BD42", rings=3, line_color="#FFC000",
         ring_colors=["#FFC000", "#75BD42", "#75BD42"]),
    dict(name="其他来源", color="#E66B4C", rings=2),
]

# 占比（顺序即从 12 点方向顺时针绘制）
# 优先从同目录的 xlsx（B3:G5，两组数据）读取；读不到就用这里的备份值
DATA = {
    "近7天": [0.44, 0.17, 0.15, 0.08, 0.16],
    "近30天": [0.38, 0.21, 0.15, 0.09, 0.17],
}


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
    except Exception as exc:          # 文件被占用 / 缺 openpyxl / 格式变了
        print(f"[提示] 未能从 xlsx 读取数据（{exc}），使用脚本内置数值。")
        return None


_XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                     "3 动态南丁格尔圆环图.xlsx")
if os.path.exists(_XLSX):
    _loaded = load_data_from_xlsx(_XLSX)
    if _loaded:
        DATA = _loaded

# 两个状态的标签与引线排版（量自两张截图）
#   text：文字左端 x、基线 y
#   line：水平段外端 → 拐点 → 端点圆点
LAYOUT = {
    "近7天": {
        "首页推荐": dict(text=(734.5, 256.0), line=[(860, 240.5), (720, 240.5), (681, 287)]),
        "关注页面": dict(text=(524.5, 631.5), line=[(711, 662.5), (576, 662.5), (551, 619)]),
        "搜索": dict(text=(168.5, 523.5), line=[(144, 577.5), (285, 577.5), (336, 535)]),
        "个人主页": dict(text=(158.5, 296.0), line=[(135, 330.5), (277, 330.5), (304, 376)]),
        "其他来源": dict(text=(208.5, 209.0), line=[(201, 221.5), (348, 221.5), (378, 267)]),
        "active": "近7天",
    },
    "近30天": {
        "首页推荐": dict(text=(720.5, 220.0), line=[(861, 240.5), (721, 240.5), (682, 287)]),
        "关注页面": dict(text=(576.5, 633.5), line=[(712, 662.5), (577, 662.5), (552, 619)]),
        "搜索": dict(text=(180.5, 546.5), line=[(145, 577.5), (286, 577.5), (337, 535)]),
        "个人主页": dict(text=(154.5, 315.0), line=[(136, 331.5), (278, 331.5), (305, 376)]),
        "其他来源": dict(text=(204.5, 213.0), line=[(202, 222.5), (349, 222.5), (379, 268)]),
        "active": "近30天",
    },
}

# 按钮矩形：(x, y, w, h)
BTN_RECTS = {"近7天": (639, 121, 127, 55), "近30天": (783, 120, 127, 55)}


def darken(hex_color, factor=0.86):
    """压暗颜色，用于环带之间的细分隔线。"""
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i:i + 2], 16) for i in (0, 2, 4))
    return (r * factor / 255, g * factor / 255, b * factor / 255)


def ease(t):
    """缓入缓出，让补间动画不呆板。"""
    return t * t * (3.0 - 2.0 * t)


# --------------------------------------------------------------------------
# 4. 图表对象：一次建好所有图元，之后只改参数（这样动画才不闪烁）
# --------------------------------------------------------------------------
class RoseChart:
    """南丁格尔圆环图。p=0 对应近7天，p=1 对应近30天，中间值即过渡状态。"""

    def __init__(self, fig):
        self.fig = fig
        self.progress = 0.0
        self.anim = None
        self.hover = -1
        self.btn_override = None   # 点击后按钮立刻切到目标态，不等动画过半

        fig.patch.set_facecolor(BG)
        ax = fig.add_axes([0, 0, 1, 1])
        ax.set_facecolor(BG)
        ax.set_xlim(0, W)
        ax.set_ylim(H, 0)          # y 轴向下，与图片行号一致
        ax.set_aspect("equal")
        ax.axis("off")
        self.ax = ax

        self.wedges = []           # 每个渠道的若干层环带
        self.seps = []             # 环带之间的虚线分隔弧
        self.hovers = []           # 悬停高亮层
        self.lines = []            # 引线
        self.dots = []             # 端点圆点
        self.labels = []           # 文字标签

        for seg in SEGMENTS:
            colors = seg.get("ring_colors") or [seg["color"]] * seg["rings"]
            rings = []
            for k in range(seg["rings"]):
                w = Wedge((CX, CY), R_IN + (k + 1) * RING_W, 0.0, 1.0,
                          width=RING_W, facecolor=colors[k],
                          edgecolor="none", linewidth=0, zorder=3)
                ax.add_patch(w)
                rings.append(w)
            self.wedges.append(rings)

            # 层与层之间是虚线弧（原图如此），颜色比填充色压暗约 15%
            seps = []
            for k in range(1, seg["rings"]):
                r = R_IN + k * RING_W
                arc, = ax.plot([], [], color=darken(colors[k], 0.85),
                               linewidth=0.72, linestyle=(0, (2.2, 2.2)),
                               solid_capstyle="butt", dash_capstyle="butt",
                               zorder=3.2)
                seps.append((r, arc))
            self.seps.append(seps)

            hov = Wedge((CX, CY), R_IN + seg["rings"] * RING_W, 0.0, 1.0,
                        width=seg["rings"] * RING_W,
                        facecolor="#FFFFFF", edgecolor="none",
                        alpha=0.0, zorder=3.6)
            ax.add_patch(hov)
            self.hovers.append(hov)

            line, = ax.plot([], [], color=seg.get("line_color", seg["color"]),
                            linewidth=LINE_W,
                            solid_capstyle="butt", solid_joinstyle="miter",
                            zorder=5)
            self.lines.append(line)

            dot = Circle((0, 0), DOT_R, facecolor=seg.get("line_color", seg["color"]),
                         edgecolor="none", zorder=6)
            ax.add_patch(dot)
            self.dots.append(dot)

            txt = ax.text(0, 0, "", fontproperties=FP_TEXT, fontsize=FS_LABEL,
                          color=C_LABEL, ha="left", va="baseline", zorder=7)
            self.labels.append(txt)

        # 标题与脚注
        ax.text(TITLE_XY[0], TITLE_XY[1], TITLE_TEXT,
                fontproperties=FP_TITLE, fontsize=FS_TITLE, color=C_TITLE,
                ha="left", va="baseline", zorder=7)
        ax.text(NOTE_XY[0], NOTE_XY[1], NOTE_TEXT,
                fontproperties=FP_TEXT, fontsize=FS_NOTE, color=C_NOTE,
                ha="left", va="baseline", zorder=7)

        # 右上角单选按钮
        self.btn_inner = {}
        for label, (x, y, w, h) in BTN_RECTS.items():
            ax.add_patch(Rectangle((x, y), w, h, facecolor=C_BTN,
                                   edgecolor="none", zorder=8))
            rx, ry = x + 17.5, y + h / 2.0
            ax.add_patch(Circle((rx, ry), 11.0, facecolor=C_RADIO_RING,
                                edgecolor="none", zorder=9))
            ax.add_patch(Circle((rx, ry), 7.0, facecolor="#FFFFFF",
                                edgecolor="none", zorder=10))
            inner = Circle((rx, ry), 5.0, facecolor=C_TEXT_DARK,
                           edgecolor="none", zorder=11)
            ax.add_patch(inner)
            self.btn_inner[label] = inner
            ax.text(x + 41, ry + 1, label, fontproperties=FP_TEXT,
                    fontsize=FS_BTN, color=C_TEXT_DARK,
                    ha="left", va="center", zorder=12)

        self.set_progress(0.0)

    # -- 把 p 应用到所有图元上 ------------------------------------------------
    def set_progress(self, p):
        p = float(np.clip(p, 0.0, 1.0))
        self.progress = p

        v0, v1 = DATA["近7天"], DATA["近30天"]
        vals = [a + (b - a) * p for a, b in zip(v0, v1)]

        phi = 0.0                          # 从 12 点方向起，顺时针累计角度
        for i, seg in enumerate(SEGMENTS):
            sweep = vals[i] * 360.0
            # 绘图像素坐标 y 向下，所以"钟面角 φ"换算成 matplotlib 角度是 φ-90
            t1, t2 = phi - 90.0, phi + sweep - 90.0
            for w in self.wedges[i]:
                w.set_theta1(t1)
                w.set_theta2(t2)
            for r, arc in self.seps[i]:
                n = max(2, int(sweep * 2))
                ang = np.deg2rad(np.linspace(phi, phi + sweep, n))
                arc.set_data(CX + r * np.sin(ang), CY - r * np.cos(ang))
            self.hovers[i].set_theta1(t1)
            self.hovers[i].set_theta2(t2)

            a, b = LAYOUT["近7天"][seg["name"]], LAYOUT["近30天"][seg["name"]]
            pts = [(x0 + (x1 - x0) * p, y0 + (y1 - y0) * p)
                   for (x0, y0), (x1, y1) in zip(a["line"], b["line"])]
            self.lines[i].set_data([q[0] for q in pts], [q[1] for q in pts])
            self.dots[i].set_center(pts[-1])

            tx = a["text"][0] + (b["text"][0] - a["text"][0]) * p
            ty = a["text"][1] + (b["text"][1] - a["text"][1]) * p
            self.labels[i].set_position((tx, ty))
            self.labels[i].set_text(f"{seg['name']}, {vals[i] * 100:.0f}%")

            phi += sweep

        # 选中态：黑心只在对应按钮里出现
        sel = p if self.btn_override is None else self.btn_override
        self.btn_inner["近7天"].set_visible(sel < 0.5)
        self.btn_inner["近30天"].set_visible(sel >= 0.5)
        if self.btn_override is not None and abs(p - self.btn_override) < 1e-9:
            self.btn_override = None

    # -- 补间动画 -------------------------------------------------------------
    def animate_to(self, target, frames=16, interval=24):
        end = 0.0 if target == "近7天" else 1.0
        start = self.progress
        if abs(end - start) < 1e-6:
            return
        if self.anim is not None:            # 上一次动画还没播完就先掐掉
            try:
                self.anim.event_source.stop()
            except Exception:
                pass
        self.btn_override = end
        values = [start + (end - start) * ease(t)
                  for t in np.linspace(0.0, 1.0, frames)]
        self.anim = FuncAnimation(self.fig, self.set_progress, frames=values,
                                  interval=interval, repeat=False, blit=False)
        self.fig.canvas.draw_idle()

    # -- 鼠标 / 键盘 ----------------------------------------------------------
    def install_events(self):
        c = self.fig.canvas
        c.mpl_connect("button_press_event", self._on_click)
        c.mpl_connect("motion_notify_event", self._on_move)
        c.mpl_connect("key_press_event", self._on_key)
        try:
            self.fig.canvas.manager.set_window_title(
                "3 动态南丁格尔圆环图 — 点击「近7天 / 近30天」切换")
        except Exception:
            pass

    def _to_data(self, event):
        if event.x is None or event.y is None:
            return None
        return self.ax.transData.inverted().transform((event.x, event.y))

    def _on_click(self, event):
        pt = self._to_data(event)
        if pt is None:
            return
        x, y = pt
        for label, (bx, by, bw, bh) in BTN_RECTS.items():
            if bx <= x <= bx + bw and by <= y <= by + bh:
                self.animate_to(label)
                return

    def _on_key(self, event):
        if event.key in ("right", "2"):
            self.animate_to("近30天")
        elif event.key in ("left", "1"):
            self.animate_to("近7天")

    def _on_move(self, event):
        pt = self._to_data(event)
        if pt is None:
            return
        x, y = pt
        on_btn = any(bx <= x <= bx + bw and by <= y <= by + bh
                     for bx, by, bw, bh in BTN_RECTS.values())
        try:
            self.fig.canvas.set_cursor(1 if on_btn else 0)   # 1=手型, 0=默认
        except Exception:
            pass

        # 悬停扇区：落在环带上就高亮对应渠道
        dx, dy = x - CX, y - CY
        r = np.hypot(dx, dy)
        idx = -1
        if R_IN <= r <= R_OUT:
            phi = (np.degrees(np.arctan2(dx, -dy))) % 360.0   # 从 12 点顺时针
            acc = 0.0
            v0, v1 = DATA["近7天"], DATA["近30天"]
            vals = [a + (b - a) * self.progress for a, b in zip(v0, v1)]
            for i, v in enumerate(vals):
                acc += v * 360.0
                if phi <= acc:
                    # 只有落在该渠道真实层高以内才算命中
                    if r <= R_IN + SEGMENTS[i]["rings"] * RING_W + 1.0:
                        idx = i
                    break
        if idx != self.hover:
            if 0 <= self.hover:
                self.hovers[self.hover].set_alpha(0.0)
                self.labels[self.hover].set_color(C_LABEL)
            self.hover = idx
            if idx >= 0:
                self.hovers[idx].set_alpha(0.14)
                self.labels[idx].set_color("#FFFFFF")
            self.fig.canvas.draw_idle()


# --------------------------------------------------------------------------
# 5. 出图
# --------------------------------------------------------------------------
OUT_DIR = r"C:\Users\fjwsm\Desktop\学习资料\大三上 第一学期\数据可视化\新建文件夹\图\第三章"


def make_figure():
    fig = plt.figure(figsize=(W / DPI, H / DPI), dpi=DPI, facecolor=BG)
    return fig, RoseChart(fig)


def save_png(period, path):
    fig, chart = make_figure()
    chart.set_progress(0.0 if period == "近7天" else 1.0)
    fig.savefig(path, dpi=DPI, facecolor=BG)
    plt.close(fig)
    return path


def save_gif(path, fps=20, hold=7, morph=16):
    """两个状态来回变形的动图：停留 → 变形 → 停留 → 变形回…（首尾相接循环）"""
    fig, chart = make_figure()
    seq = ([0.0] * hold
           + [ease(t) for t in np.linspace(0.0, 1.0, morph)]
           + [1.0] * hold
           + [1.0 - ease(t) for t in np.linspace(0.0, 1.0, morph)])
    ani = FuncAnimation(fig, chart.set_progress, frames=seq, interval=1000 / fps)
    ani.save(path, writer=PillowWriter(fps=fps))
    plt.close(fig)
    return path


def show_interactive():
    """弹出可交互窗口：点按钮 / 按方向键切换，鼠标悬停高亮。"""
    if not _backend_is_gui():
        print("[提示] 当前后端为", matplotlib.get_backend(),
              "，无法弹窗交互。可在命令行运行本文件，或把绘图后端切到 Qt5Agg。")
        return None
    fig, chart = make_figure()
    chart.install_events()
    fig.canvas.draw_idle()
    plt.show()
    return chart


if __name__ == "__main__":
    os.makedirs(OUT_DIR, exist_ok=True)

    for period in ("近7天", "近30天"):
        out = os.path.join(OUT_DIR, f"3 动态南丁格尔圆环图_{period}.png")
        print("[OK] 已保存", save_png(period, out))

    gif = os.path.join(OUT_DIR, "3 动态南丁格尔圆环图.gif")
    print("[OK] 已保存", save_gif(gif))

    if SHOW_INTERACTIVE and not _FORCE_NO_GUI:
        show_interactive()
