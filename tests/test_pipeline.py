# -*- coding: utf-8 -*-
"""回归测试：只依赖纯业务模块，不需要图形界面，CI 里可以直接跑。

覆盖四条最容易回归的规则：
  1. 默认教室布局（7×11、第 4/8 列过道、63 个座位）
  2. 5 套方案互不相同，且空位只留在最后几排末尾
  3. 预览表往返 + 缺考绝不当 0 分
  4. 导出 PNG 脱敏：无过道格、无重点关注标记（文字与底色都没有）
"""
from __future__ import annotations

import sys
from pathlib import Path

# 让 `pytest` 从任意目录启动都能找到 src/ 下的平铺模块
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pytest
from matplotlib.colors import to_rgba

from excel_io import students_from_table, students_to_table
from exporter import build_seat_figure, png_columns
from layout import BLOCK, SEAT, ClassroomLayout
from models import COLOR, Student
from seating import SeatingOptions, generate_plans


@pytest.fixture(scope="module")
def bundle():
    """35 名学生 + 6×9（48 个座位）的小教室，留出空位便于检查空位规则。"""
    students = [
        Student(sid=f"2026{i:03d}", name=f"学生{i:02d}",
                score=180 + (i * 7) % 110,
                gender="男" if i % 2 else "女",
                tag="重点关注" if i in (3, 17) else "")
        for i in range(1, 36)
    ]
    layout = ClassroomLayout.default(6, 9)
    plans, _notices = generate_plans(
        students, layout, SeatingOptions(plan_count=5, spread_weight=0.7))
    return students, layout, plans


def test_default_layout():
    layout = ClassroomLayout.default()
    assert (layout.rows, layout.cols) == (7, 11)
    aisles = [c for c in range(layout.cols) if layout.get(0, c) == BLOCK]
    assert aisles == [3, 7], "用户视角第 4、8 列应为过道"
    assert len(layout.seats()) == 63


def test_five_plans_are_different(bundle):
    _, _, plans = bundle
    assert len(plans) == 5
    signatures = {
        tuple(sorted((s.name, pos) for pos, s in p.assignment.items()))
        for p in plans
    }
    assert len(signatures) == 5


def test_empty_seats_only_at_back(bundle):
    students, layout, plans = bundle
    seats = layout.seats()
    assert len(students) < len(seats), "这个用例需要学生少于座位才有意义"
    for plan in plans:
        occupied = set(plan.assignment)
        assert occupied == set(seats[:len(occupied)]), \
            f"{plan.name} 出现夹在学生中间的空位"


def test_preview_roundtrip(bundle):
    students, _, _ = bundle
    rows = students_to_table(students)
    back, warnings = students_from_table(rows)
    assert len(back) == len(students)
    assert not warnings
    assert all(s.name for s in back)


def test_missing_score_is_not_zero():
    rows = [["2026001", "学生01", "缺考", "男", "", "", ""],
            ["2026002", "学生02", "N/A", "女", "", "", ""]]
    students, _ = students_from_table(rows)
    assert len(students) == 2
    assert all(s.score is None for s in students), "缺考不能按 0 分处理"


def test_png_is_desensitized(bundle):
    _, layout, plans = bundle
    draw_cols = png_columns(layout)
    cells = sum(1 for r in range(layout.rows) for c in draw_cols
                if layout.get(r, c) == SEAT)
    fig, ax = build_seat_figure(plans[0], layout)
    try:
        assert fig._suptitle is None, "PNG 里不应出现方案标题"
        assert len(ax.texts) == 1 + layout.rows + cells, "PNG 文字数量异常"
        assert len(ax.patches) == 1 + cells, "PNG 不应多画过道/讲台方格"
        forbidden = ("过道", "重点关注", "重点", "视力", "图例",
                     "导出时间", "座位示意图", "浅红")
        assert not [t.get_text() for t in ax.texts
                    if any(k in t.get_text() for k in forbidden)]
        assert not any(to_rgba(p.get_facecolor()) == to_rgba(COLOR["focus"])
                       for p in ax.patches), "PNG 不能用浅红底泄露重点关注学生"
    finally:
        plt.close(fig)
