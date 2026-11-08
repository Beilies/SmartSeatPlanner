# -*- coding: utf-8 -*-
"""
导出模块（exporter.py）
------------------------------------------------------------------
1. 导出 Excel 座位表：名单 Sheet + 座位示意图 Sheet（带颜色，可直接打印）
2. 导出座位画布为 PNG 图片：用 matplotlib 重绘，只含座位信息（讲台 / 排号 / 座位方格），
   不含标题、图例、行列数、导出时间、特殊标记小字等任何附加说明；
   过道按原始位置留白（不画方框、不写字），方便直接打印 / 发家长群
3. 导出方案统计报表：多方案横向对比 + 分区分布 + 提示与冲突

所有函数失败时抛 ExportError，由 GUI 转成中文弹窗。
"""
from __future__ import annotations

import os
from datetime import datetime
from typing import List, Optional

import matplotlib
matplotlib.use("Agg")                      # 不需要弹窗，只生成图片
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Rectangle
from openpyxl import Workbook
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from evaluate import compare_rows
from layout import BLOCK, PODIUM, SEAT, ZONE_NAMES, ClassroomLayout
from models import COLOR, Student
from seating import Plan


class ExportError(Exception):
    """导出失败（权限、文件被占用、缺字体等），GUI 会友好提示。"""


# 座位图是给家长看 / 打印张贴的，光栅化分辨率不能省：
# 180 DPI 放大到 A4 时笔画会发虚，统一按 300 DPI 输出。
FIGURE_DPI = 200                  # 画布内部渲染分辨率
PNG_DPI = 300                     # savefig 输出分辨率（打印级）


# ----------------------------------------------------------------------
# 中文字体（Windows 一般有微软雅黑 / 黑体；找不到时降级，不会崩溃）
# ----------------------------------------------------------------------
def setup_chinese_font() -> str:
    """设置 matplotlib 中文字体，返回实际使用的字体名（供提示用）。"""
    candidates = ["Microsoft YaHei", "SimHei", "SimSun", "KaiTi", "FangSong",
                  "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Zen Hei",
                  "Arial Unicode MS", "DengXian"]
    try:
        available = {f.name for f in font_manager.fontManager.ttflist}
    except Exception:
        available = set()
    chosen = next((n for n in candidates if n in available), "DejaVu Sans")
    plt.rcParams["font.sans-serif"] = [chosen, "DejaVu Sans"]
    plt.rcParams["axes.unicode_minus"] = False
    # 高分辨率渲染 + 字体抗锯齿：文字笔画不再是一圈糊边
    plt.rcParams["figure.dpi"] = FIGURE_DPI
    plt.rcParams["savefig.dpi"] = PNG_DPI
    plt.rcParams["text.antialiased"] = True
    return chosen


setup_chinese_font()

def _xl(color: str) -> str:
    """界面配色是 #RRGGBB，openpyxl 只接受不带 # 的 6 位（或 8 位 aRGB）文本。"""
    return color.lstrip("#")


THIN = Side(style="thin", color="9AA6B2")
GRID_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="4472C4")
HEAD_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
CELL_FONT = Font(name="微软雅黑", size=10, color=_xl(COLOR["text"]))
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _ensure_ext(path: str, ext: str) -> str:
    root, cur = os.path.splitext(path)
    return path if cur.lower() == ext.lower() else root + ext


def _stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def _cell_style(value: str, layout_value: str, student: Optional[Student]):
    """返回 (显示文本, 填充色, 边框色)，Excel 与 PNG 共用同一套配色口径。"""
    if layout_value == BLOCK:
        return "过道", COLOR["block"], COLOR["block_border"]
    if layout_value == PODIUM:
        return "讲台", COLOR["podium"], COLOR["podium"]
    if student is None:
        return "空位", COLOR["empty"], COLOR["empty_border"]
    tag = student.short_tag()
    text = student.name + ("\n(" + tag + ")" if tag else "")
    return (text, COLOR["focus"], COLOR["focus_border"]) if student.is_focus() \
        else (text, COLOR["normal"], COLOR["normal_border"])


# ----------------------------------------------------------------------
# 一、Excel 座位表
# ----------------------------------------------------------------------
def export_seating_excel(path: str, plan: Plan, layout: ClassroomLayout,
                         students: List[Student]) -> str:
    """导出「名单 + 座位示意图」两个 Sheet 的 xlsx，返回实际路径。"""
    try:
        path = _ensure_ext(path, ".xlsx")
        wb = Workbook()

        # ---- Sheet 1：学生名单 ----
        ws = wb.active
        ws.title = "学生名单"
        headers = ["学号", "姓名", "性别", "总分", "特殊标记", "禁止同桌名单",
                   "方案", "区域", "第几排", "第几列", "同桌姓名", "备注"]
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill, cell.font, cell.alignment, cell.border = HEAD_FILL, HEAD_FONT, CENTER, GRID_BORDER
        ws.row_dimensions[1].height = 26
        ws.freeze_panes = "A2"

        pos_of = {s.name: p for p, s in plan.assignment.items()}
        desk = {}
        for a, b in layout.desk_pairs():
            desk[a], desk[b] = b, a
        for r, s in enumerate(students, start=2):
            p = pos_of.get(s.name)
            mate = ""
            if p is not None and p in desk:
                m = plan.assignment.get(desk[p])
                mate = m.name if m else ""
            values = [
                s.sid, s.name, s.gender, "" if s.score is None else s.score,
                s.tag, s.forbid_text(), plan.name,
                layout.zone(p[0]) if p else "未安排",
                (p[0] + 1) if p else "", (p[1] + 1) if p else "",
                mate, "未入座，请扩大教室行列数" if p is None else "",
            ]
            for c, v in enumerate(values, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font, cell.alignment, cell.border = CELL_FONT, CENTER, GRID_BORDER
                if s.is_focus() and c <= 5:
                    cell.fill = PatternFill("solid", fgColor=_xl(COLOR["focus"]))
        widths = [14, 12, 8, 10, 18, 24, 18, 10, 10, 10, 12, 26]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}1"

        # ---- Sheet 2：座位示意图（前排在上，颜色与界面一致） ----
        pic = wb.create_sheet("座位示意图")
        pic.sheet_view.showGridLines = False
        total_cols = layout.cols + 1               # A 列放行号
        pic.merge_cells(start_row=1, start_column=1, end_row=1, end_column=total_cols)
        title = pic.cell(row=1, column=1,
                         value=f"{plan.name}｜座位示意图（前排在最上方，导出时间 {_stamp()}）")
        title.font = Font(name="微软雅黑", size=14, bold=True, color="1F2933")
        title.alignment = CENTER
        pic.row_dimensions[1].height = 30

        # 讲台横条
        pic.merge_cells(start_row=2, start_column=2, end_row=2, end_column=total_cols)
        podium = pic.cell(row=2, column=2, value="讲　台　（前方）")
        podium.fill = PatternFill("solid", fgColor=_xl(COLOR["podium"]))
        podium.font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        podium.alignment = CENTER
        pic.row_dimensions[2].height = 26

        for r in range(layout.rows):
            rr = r + 3
            label = pic.cell(row=rr, column=1, value=f"第{r + 1}排")
            label.font = Font(name="微软雅黑", size=10, bold=True, color=_xl(COLOR["subtext"]))
            label.alignment = CENTER
            pic.row_dimensions[rr].height = 38
            for c in range(layout.cols):
                cell_value = layout.get(r, c)
                stu = plan.assignment.get((r, c))
                text, fill, border_color = _cell_style("", cell_value, stu)
                cell = pic.cell(row=rr, column=c + 2, value=text)
                cell.fill = PatternFill("solid", fgColor=_xl(fill))
                thin = Side(style="thin", color=_xl(border_color))
                cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
                cell.alignment = CENTER
                if cell_value in (BLOCK, PODIUM):
                    cell.font = Font(name="微软雅黑", size=9,
                                     color="FFFFFF" if cell_value == PODIUM else "5A6B7B")
                else:
                    cell.font = Font(name="微软雅黑", size=10, bold=bool(stu),
                                     color=_xl(COLOR["text"]))
        pic.column_dimensions["A"].width = 10
        for c in range(layout.cols):
            pic.column_dimensions[get_column_letter(c + 2)].width = 13

        # 图例
        legend_row = layout.rows + 4
        pic.cell(row=legend_row, column=1, value="图例：")
        legend = [("浅红=重点关注", COLOR["focus"]), ("白色=普通学生", COLOR["normal"]),
                  ("灰色=过道/禁用", COLOR["block"]), ("浅灰=空座位", COLOR["empty"])]
        for i, (text, color) in enumerate(legend, start=2):
            cell = pic.cell(row=legend_row, column=i, value=text)
            cell.fill = PatternFill("solid", fgColor=color.lstrip("#"))
            cell.font = Font(name="微软雅黑", size=9)
            cell.alignment = CENTER

        folder = os.path.dirname(os.path.abspath(path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        wb.save(path)
        return path
    except PermissionError:
        raise ExportError("文件正被 Excel 打开，请先关闭它，或另存为别的文件名")
    except Exception as exc:
        raise ExportError(f"导出 Excel 座位表失败：{str(exc)[:100]}")


# ----------------------------------------------------------------------
# 二、座位画布 PNG（matplotlib 重绘，可直接发家长群 / 打印）
#
# 【输出范围约定】图片里只出现「座位信息」：
#   · 讲　台（方位）、1排/2排…（排号，属于座位定位）
#   · 学生姓名方格，以及末尾几个「空位」方格
# 刻意不画的「其它信息」：
#   · 方案标题、颜色图例、行列数、前排说明、导出时间
#   · 过道 / 禁用格：按原始列位置空出来（不画灰色方框、不写「过道」），
#     座位之间自然出现空白纵向通道，家长一眼能看出教室的真实布局
#   · 老师对学生的重点关照信息：不写「重点关注 / 视力 / 学困」等小字，
#     也不再用浅红底色区分——所有学生方格一律白色，家长看到的是同一张表
# 【注意】主界面画布仍保留浅红标记与灰色过道（老师排座时需要看），
#         只在这个「对外导出」的函数里做脱敏与简化；Excel 座位表也仍保留教师信息。
# ----------------------------------------------------------------------
def png_columns(layout: ClassroomLayout) -> List[int]:
    """返回 PNG 网格保留的原始列号：**全部列，含过道列**。

    早期版本会把整列过道压缩删掉，结果三段座位紧挨在一起、看不出通道；
    现在按原始列号排布，过道列照常占一格宽度但完全留白（画座位时跳过非 SEAT 格），
    既不留灰色方框，也不会把教室布局挤变形。
    """
    return list(range(layout.cols))


def png_aisle_columns(layout: ClassroomLayout) -> List[int]:
    """返回 PNG 中应当整列留白的原始列号（自检 / 测试用它做回归断言）。"""
    return [c for c in range(layout.cols)
            if all(layout.get(r, c) != SEAT for r in range(layout.rows))]


def build_seat_figure(plan: Plan, layout: ClassroomLayout):
    """画出「纯座位信息」画布，返回 (fig, ax)。

    单独抽成函数是为了让自检脚本能直接检查图上到底画了哪些文字 / 颜色，
    防止以后改代码时又把标题、图例、过道方框或重点关注标记加回来。
    """
    rows = layout.rows
    draw_cols = png_columns(layout)
    out_cols = max(1, len(draw_cols))              # 极端容错：全无座位也不让坐标崩
    # 没有标题、页脚，画布按实际网格收紧；过道列也算进宽度，用来留白
    fig_w = max(8.0, out_cols * 1.18 + 0.9)
    fig_h = max(4.0, rows * 0.78 + 1.15)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=FIGURE_DPI)
    fig.patch.set_facecolor("white")
    # 坐标紧贴内容：左边留一点给排号，上边留给讲台，下边不再留页脚空间
    ax.set_xlim(-0.60, out_cols + 0.10)
    ax.set_ylim(-0.06, rows + 1.00)
    ax.set_aspect("equal")
    ax.axis("off")
    fig.subplots_adjust(left=0.01, right=0.99, bottom=0.01, top=0.99)

    # 讲台横条（属于座位方位信息，保留），居中于整个座位网格
    ax.add_patch(Rectangle((out_cols / 2 - 1.35, rows + 0.32), 2.7, 0.6,
                          facecolor=COLOR["podium"], edgecolor=COLOR["podium"]))
    ax.text(out_cols / 2, rows + 0.62, "讲　台", ha="center", va="center",
            fontsize=13, color="white", fontweight="bold")

    # 字号按「一行实际有多少个座位」决定，过道列变多也不该把姓名挤小
    seat_per_row = max((sum(1 for c in draw_cols if layout.get(r, c) == SEAT)
                        for r in range(rows)), default=1)
    name_size = 11 if seat_per_row <= 10 else 9
    for r in range(rows):
        y = rows - 1 - r                           # 第 0 排画在最上方
        # 排号：帮助老师/家长定位第几排，属于座位信息
        ax.text(-0.18, y + 0.3, f"{r + 1}排", ha="right", va="center",
                fontsize=9, color=COLOR["subtext"])
        for x, c in enumerate(draw_cols):
            if layout.get(r, c) != SEAT:
                continue                           # 过道 / 禁用格：整格留白，不画框、不写字
            stu = plan.assignment.get((r, c))
            if stu is None:
                fill, border = COLOR["empty"], COLOR["empty_border"]
                label, size, color = "空位", name_size - 1, "#9AA6B2"
            else:
                # 统一白色：不泄露「重点关注」等老师内部标记
                fill, border = COLOR["normal"], COLOR["normal_border"]
                label, size, color = stu.name, name_size, COLOR["text"]
            ax.add_patch(Rectangle((x + 0.04, y + 0.04), 0.92, 0.56,
                                  facecolor=fill, edgecolor=border, linewidth=1.1))
            ax.text(x + 0.5, y + 0.32, label, ha="center", va="center",
                    fontsize=size, color=color,
                    fontweight="bold" if stu is not None else "normal")
    return fig, ax


def export_canvas_png(path: str, plan: Plan, layout: ClassroomLayout) -> str:
    """把 build_seat_figure 画好的座位图存成 PNG，返回实际路径。"""
    try:
        path = _ensure_ext(path, ".png")
        fig, _ = build_seat_figure(plan, layout)
        folder = os.path.dirname(os.path.abspath(path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        # tight + 很小的留白：图片边缘干净，打印不会被裁到字
        fig.savefig(path, dpi=PNG_DPI, bbox_inches="tight", pad_inches=0.06,
                    facecolor="white")
        plt.close(fig)
        return path
    except Exception as exc:
        raise ExportError(f"导出 PNG 图片失败：{str(exc)[:100]}")


# ----------------------------------------------------------------------
# 三、方案统计报表
# ----------------------------------------------------------------------
def export_report_xlsx(path: str, plans: List[Plan], layout: ClassroomLayout,
                       students: List[Student], notices: Optional[List[str]] = None) -> str:
    try:
        path = _ensure_ext(path, ".xlsx")
        wb = Workbook()
        ws = wb.active
        ws.title = "方案对比"
        columns, rows = compare_rows(plans)
        for c, h in enumerate(columns, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill, cell.font, cell.alignment, cell.border = HEAD_FILL, HEAD_FONT, CENTER, GRID_BORDER
        for r, row in enumerate(rows, start=2):
            for c, v in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.font, cell.alignment, cell.border = CELL_FONT, CENTER, GRID_BORDER
                if r % 2 == 0:
                    cell.fill = PatternFill("solid", fgColor="F4F7FA")
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "C2"
        widths = [20, 40, 12, 12, 12, 14, 11, 12, 14, 16, 14, 13, 12, 46]
        for c, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(c)].width = w
        # 「区均分最大差」加数据条，越小越好，一眼看出哪套更均衡
        if rows:
            ws.conditional_formatting.add(
                f"F2:F{len(rows) + 1}",
                DataBarRule(start_type="num", start_value=0, end_type="max",
                            color="7FB3E8", showValue=True))

        # 分区分布明细
        detail = wb.create_sheet("分区分布")
        detail_cols = ["方案", "区域", "入座人数", "平均分", "优生人数", "学困生人数"]
        for c, h in enumerate(detail_cols, start=1):
            cell = detail.cell(row=1, column=c, value=h)
            cell.fill, cell.font, cell.alignment = HEAD_FILL, HEAD_FONT, CENTER
        rr = 2
        for p in plans:
            st = p.stats or {}
            for z in ZONE_NAMES:
                vals = [p.name, z, st.get("zone_counts", {}).get(z, 0),
                        round(st.get("zone_avgs", {}).get(z, 0), 1),
                        st.get("high_by_zone", {}).get(z, 0),
                        st.get("low_by_zone", {}).get(z, 0)]
                for c, v in enumerate(vals, start=1):
                    cell = detail.cell(row=rr, column=c, value=v)
                    cell.font, cell.alignment, cell.border = CELL_FONT, CENTER, GRID_BORDER
                rr += 1
        for c, w in enumerate([20, 10, 12, 12, 12, 14], start=1):
            detail.column_dimensions[get_column_letter(c)].width = w

        # 提示 / 冲突 / 降级记录
        note = wb.create_sheet("提示与冲突")
        note.cell(row=1, column=1, value="类型").fill = HEAD_FILL
        note.cell(row=1, column=2, value="内容").fill = HEAD_FILL
        for c in (1, 2):
            note.cell(row=1, column=c).font = HEAD_FONT
            note.cell(row=1, column=c).alignment = CENTER
        rr = 2
        lines = [(f"生成时间 {_stamp()}", "信息")]
        lines += [(n, "降级提示") for n in (notices or [])]
        for p in plans:
            lines += [(f"{p.name}：{c}", "约束冲突") for c in p.conflicts]
            if p.unassigned:
                lines.append((f"{p.name} 未入座：" + "、".join(s.name for s in p.unassigned),
                              "未入座名单"))
        for text, kind in lines:
            note.cell(row=rr, column=1, value=kind).font = CELL_FONT
            cell = note.cell(row=rr, column=2, value=text)
            cell.font, cell.alignment = CELL_FONT, Alignment(vertical="center", wrap_text=True)
            rr += 1
        note.column_dimensions["A"].width = 16
        note.column_dimensions["B"].width = 100

        folder = os.path.dirname(os.path.abspath(path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        wb.save(path)
        return path
    except PermissionError:
        raise ExportError("报表文件正被 Excel 打开，请先关闭或换个文件名")
    except Exception as exc:
        raise ExportError(f"导出统计报表失败：{str(exc)[:100]}")
