# -*- coding: utf-8 -*-
"""
自检脚本（selfcheck.py，可选运行，不需要图形界面）
------------------------------------------------------------------
自动造一份「很脏」的 Excel：表头不在第一行、合并单元格、干扰行、
空行、底部合计行、缺考、N/A、异常分数、重名学生、空姓名……
然后跑完 导入 → 预览表往返 → 布局 → 5 套方案 → 评估 → 三种导出，
把结果打印出来并放到「自检输出」目录。

运行： python src/selfcheck.py
      （或在项目根目录用 pytest 跑 tests/ 里的回归用例）
"""
from __future__ import annotations

import os
import sys

_SRC_DIR = os.path.dirname(os.path.abspath(__file__))
if _SRC_DIR not in sys.path:
    sys.path.insert(0, _SRC_DIR)
from app_paths import BASE_DIR, OUTPUT_DIRNAME
OUT_DIR = os.path.join(BASE_DIR, OUTPUT_DIRNAME)

from excel_io import import_excel, students_from_table, students_to_table, download_template
from layout import BLOCK, SEAT, ClassroomLayout
from models import COLOR
from seating import SeatingOptions, generate_plans, apply_swap
from evaluate import attach_stats
from exporter import (export_seating_excel, export_canvas_png, export_report_xlsx,
                      build_seat_figure, png_columns, png_aisle_columns)


def make_dirty_excel(path: str) -> None:
    """构造一份故意不规整的成绩表，用来检验容错能力。"""
    from openpyxl import Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = "期中成绩"
    ws.merge_cells("A1:I1")                       # 合并单元格大标题
    ws["A1"] = "六年级三班期中成绩表（表头在第 4 行）"
    ws["A3"] = "制表：教务处"
    ws["C3"] = "2026-11-07"
    headers = ["学号", "姓名", "总分", "性别", "特殊标记", "禁止同桌名单",
               "语文", "数学", "英语"]
    for c, h in enumerate(headers, start=1):
        ws.cell(row=4, column=c, value=h)

    r = 5
    for i in range(1, 37):
        score = 180 + (i * 7) % 110
        row = [f"2026{i:03d}", f"学生{i:02d}", score,
               "男" if i % 2 else "女", "", "",
               score // 3, score // 3, score - 2 * (score // 3)]
        if i == 3:
            row[4] = "重点关注"          # 浅红 + 优先靠前
        if i == 8:
            row[4] = "视力需照顾"
        if i == 5:
            row[5] = "学生06、学生07"     # 禁止同桌
        if i == 6:
            row[5] = "学生05"
        if i == 11:
            row[2] = "缺考"              # 总分缺考，但科目有分 -> 自动求和
            row[6] = 82
        if i == 12:
            row[2] = "—"                 # 缺失值
            row[7] = "N/A"
        if i == 13:
            row[2] = ""                  # 总分留空 -> 自动求和
        if i == 14:
            row[6] = "缺考"              # 单科缺考，不按 0 分
        if i == 20:
            row[2] = "缺考"              # 总分 + 三科全部缺考 -> 无总分，仍不按 0 分
            row[6] = "缺考"
            row[7] = "—"
            row[8] = "N/A"
        if i == 15:
            row[7] = -5                  # 异常分数
        if i == 16:
            row[8] = "随便填的"           # 非数值
        if i == 17:
            row[1] = ""                  # 空姓名 -> 预览表标红
        if i == 18:
            row[0] = "2026001"           # 重复学号
        if i == 19:
            row[1] = "学生01"            # 重复姓名
        for c, v in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1

    r += 1                                  # 空行
    ws.cell(row=r + 1, column=1, value="合计")
    ws.cell(row=r + 1, column=3, value=9999)   # 底部合计行
    ws.cell(row=r + 1, column=7, value=9999)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    wb.save(path)


def main() -> int:
    os.makedirs(OUT_DIR, exist_ok=True)
    dirty = os.path.join(OUT_DIR, "脏数据测试表.xlsx")
    make_dirty_excel(dirty)
    print(f"[1] 已构造脏 Excel：{dirty}")

    result = import_excel(dirty)
    kinds = {}
    for w in result.warnings:
        kinds[w.kind] = kinds.get(w.kind, 0) + 1
    print(f"[2] 导入成功：表头第 {result.header_row} 行，学生 {len(result.students)} 人，"
          f"忽略 {result.ignored_rows} 行，告警 {len(result.warnings)} 条 {kinds}")
    assert len(result.students) >= 34, "学生行数明显不对"
    assert result.ignored_rows >= 1, "底部合计行没有被忽略"
    assert all(s.name != "合计" and s.sid != "合计" for s in result.students), \
        "底部合计行混进了学生名单"
    assert any(s.score is None for s in result.students), "缺考没有被当成缺失"
    assert kinds.get("bad_score", 0) >= 2, "异常分数没有被告警"
    assert kinds.get("duplicate", 0) >= 1, "重复学生没有被告警"
    assert kinds.get("empty_name", 0) >= 1, "空姓名没有被告警"

    # 预览表往返（模拟老师在预览窗口改数据）
    rows = students_to_table(result.students)
    back, _ = students_from_table(rows)
    print(f"[3] 预览表双向转换：{len(rows)} 行 -> {len(back)} 名学生，正常")
    assert len(back) == len(rows)

    # 空姓名行由预览窗口负责提示，排座时会安全忽略
    clean = [s for s in back if s.name.strip()]
    # 默认教室布局（老师指定）：7 行 × 11 列，第 4 / 8 列是纵向过道
    default_layout = ClassroomLayout.default()
    assert default_layout.rows == 7 and default_layout.cols == 11, "默认布局行列数不对"
    aisles = [c for c in range(default_layout.cols)
              if default_layout.get(0, c) == BLOCK]
    assert aisles == [3, 7], f"默认布局过道列不对：{aisles}"
    print(f"[4] 默认教室布局：{default_layout.summary()}（过道在第 4、8 列）")
    layout = ClassroomLayout.default(6, 9)   # 自检继续用 6×9 覆盖小教室场景

    plans, notices = generate_plans(clean, layout, SeatingOptions(plan_count=5, spread_weight=0.7))
    attach_stats(plans, layout, clean)
    print(f"[5] 生成 {len(plans)} 套方案，全局提示 {len(notices)} 条：")
    signatures = set()
    for p in plans:
        st = p.stats
        za = st["zone_avgs"]
        sig = tuple(sorted((s.name, pos) for pos, s in p.assignment.items()))
        signatures.add(sig)
        print(f"    {p.name}: 入座 {len(p.assignment)} 人，"
              f"前/中/后排均分 {za['前排']:.1f}/{za['中排']:.1f}/{za['后排']:.1f}，"
              f"黑名单相邻 {st['ban_hits']}，结论：{st['verdict']}")
    assert len(plans) == 5
    assert len(signatures) == 5, "方案之间出现了完全相同的座位表"
    assert len(plans[0].assignment) == min(len(clean), len(layout.seats()))

    # 拖拽换座 + 撤销（用算法层接口模拟界面动作）
    a = list(plans[0].assignment.keys())[0]
    b = list(plans[0].assignment.keys())[1]
    before = (plans[0].assignment.get(a), plans[0].assignment.get(b))
    apply_swap(plans[0], a, b)
    apply_swap(plans[0], a, b)
    assert before == (plans[0].assignment.get(a), plans[0].assignment.get(b)), "换座撤销不一致"
    print("[6] 拖拽换座 / 撤销等价性校验：通过")

    # 三种导出
    xlsx = export_seating_excel(os.path.join(OUT_DIR, "示例_Excel座位表.xlsx"),
                                plans[0], layout, clean)
    png = export_canvas_png(os.path.join(OUT_DIR, "示例_座位图.png"),
                            plans[0], layout)
    report = export_report_xlsx(os.path.join(OUT_DIR, "示例_方案统计报表.xlsx"),
                                plans, layout, clean, notices)
    template = download_template(os.path.join(OUT_DIR, "示例_标准模板.xlsx"))
    for p in (xlsx, png, report, template):
        assert os.path.isfile(p) and os.path.getsize(p) > 1000, f"导出文件异常：{p}"
        print(f"[7] 导出成功：{os.path.basename(p)}（{os.path.getsize(p):,} 字节）")

    # 回归检查 ③：PNG 必须是打印级分辨率，不能退回 180 DPI 的糊图
    from PIL import Image
    from exporter import FIGURE_DPI, PNG_DPI
    assert PNG_DPI >= 300 and FIGURE_DPI >= 200, "座位图渲染 / 输出 DPI 被调低了"
    with Image.open(png) as im:
        pw, ph = im.size
    assert pw >= 1800, f"座位图像素宽度不足，打印会发虚：{pw} px"
    print(f"[7b] PNG 分辨率：{pw} × {ph} 像素（画布 {FIGURE_DPI} / 输出 {PNG_DPI} DPI）")

    # 回归检查 ①：空位不能夹在学生中间。
    # 已入座座位必须恰好是「从前往后、从左到右」的前 N 个座位。
    seat_order = layout.seats()
    for p in plans:
        occupied = set(p.assignment)
        assert occupied == set(seat_order[:len(occupied)]), \
            f"{p.name} 出现夹在学生中间的空位"
    print("[8] 空位检查：所有方案的空位都集中在最后几排末尾，没有夹在学生中间")

    # 回归检查 ②：PNG 只含座位信息，且不给家长看老师内部标记。
    # 过道列按原始位置留出空白（不画框、不写字），所以文字数仍是
    # 1 讲台 + 排号 + 座位数；同时要断言过道列真的没有任何方格 / 姓名。
    import math
    import matplotlib.pyplot as plt
    from matplotlib.colors import to_rgba
    draw_cols = png_columns(layout)
    aisle_cols = png_aisle_columns(layout)
    assert len(draw_cols) == layout.cols, "PNG 又把过道列压缩掉了"
    assert aisle_cols, "自检布局应至少有一条过道，无法验证留白"
    cells = sum(1 for r in range(layout.rows) for c in draw_cols
                if layout.get(r, c) == SEAT)
    expected_texts = 1 + layout.rows + cells
    fig, ax = build_seat_figure(plans[0], layout)
    assert fig._suptitle is None, "PNG 里又出现了方案标题"
    assert len(ax.texts) == expected_texts, (
        f"PNG 文字数量异常：应有 {expected_texts} 处，实际 {len(ax.texts)} 处")
    assert len(ax.patches) == 1 + cells, "PNG 多画了非座位方格（过道方框？）"
    # patches[0] 是讲台横条，其余全部是座位方格；方格中心所在的列不能是过道列
    seat_x = [math.floor(p.get_x() + p.get_width() / 2) for p in ax.patches[1:]]
    assert all(x not in aisle_cols for x in seat_x), "过道列被画上了座位方格"
    # 姓名 / 排号文字也一样；讲台横条上的「讲　台」两字居中，不计入
    text_x = [math.floor(t.get_position()[0]) for t in ax.texts
              if t.get_text() != "讲　台"]
    assert all(x not in aisle_cols for x in text_x), "过道列被写上了文字"
    forbidden = ("过道", "图例", "浅红", "导出时间", "座位示意图", "前排",
                 "重点关注", "普通学生", "行×", "重点", "视力", "学困", "讲台格")
    extra = [t.get_text() for t in ax.texts
             if any(k in t.get_text() for k in forbidden)]
    assert not extra, f"PNG 出现附加 / 敏感文字：{extra}"
    focus_rgba = to_rgba(COLOR["focus"])
    assert not any(to_rgba(patch.get_facecolor()) == focus_rgba
                   for patch in ax.patches), "PNG 仍用浅红底泄露重点关注学生"
    plt.close(fig)
    print(f"[9] PNG 内容检查：{cells} 个座位方格（第 "
          f"{'、'.join(str(c + 1) for c in aisle_cols)} 列过道整列留白）、"
          "无过道方框、无标题图例、无重点关注标记（颜色与文字都没有）")

    print("\n✅ 全部自检通过：导入容错、预览编辑、5 套差异化方案、评估、三种导出均正常。")
    print(f"   示例文件在：{OUT_DIR}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception as exc:
        import traceback
        traceback.print_exc()
        print(f"\n❌ 自检失败：{type(exc).__name__}: {exc}")
        sys.exit(1)
