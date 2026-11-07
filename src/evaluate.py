# -*- coding: utf-8 -*-
"""
方案评估模块（evaluate.py）
------------------------------------------------------------------
把「这套座位好不好」变成老师一眼能看懂的两样东西：
    1. 前排 / 中排 / 后排 平均分（给主界面嵌入的柱状图用）
    2. 优生、学困生、重点关注学生分布 + 一句话中文结论（给文本面板用）

注意：缺考 / 无总分的学生不参与平均分计算（绝不按 0 分拉低均分），
      但在人数统计里仍然算「已入座」。
"""
from __future__ import annotations

from typing import Dict, List, Tuple

from layout import ClassroomLayout, ZONE_NAMES
from models import Student
from seating import Plan


def _thresholds(students: List[Student]) -> Tuple[float, float]:
    """优生 / 学困生分数阈值：有效成绩的前 25% / 后 25%。"""
    scores = sorted(s.score for s in students if s.score is not None)
    if not scores:
        return float("inf"), float("-inf")
    return (scores[int(0.75 * (len(scores) - 1))],
            scores[int(0.25 * (len(scores) - 1))])


def _ban_pairs(students: List[Student]) -> set:
    """重建禁止相邻名单（按姓名匹配，与排座算法口径一致）。"""
    names = {s.name.strip() for s in students}
    pairs = set()
    for s in students:
        for target in s.forbid:
            t = target.strip()
            if t in names and t != s.name.strip():
                pairs.add(frozenset((s.name.strip(), t)))
    return pairs


def evaluate_plan(plan: Plan, layout: ClassroomLayout,
                  students: List[Student]) -> dict:
    """计算单套方案的统计结果，返回结构化字典。"""
    assign = plan.assignment
    all_scores = [s.score for s in students if s.score is not None]
    overall = sum(all_scores) / len(all_scores) if all_scores else 0.0
    high_cut, low_cut = _thresholds(students)

    zone_scores: Dict[str, List[float]] = {z: [] for z in ZONE_NAMES}
    zone_counts: Dict[str, int] = {z: 0 for z in ZONE_NAMES}
    high_by_zone: Dict[str, int] = {z: 0 for z in ZONE_NAMES}
    low_by_zone: Dict[str, int] = {z: 0 for z in ZONE_NAMES}
    high_total = low_total = focus_total = focus_good = 0

    for pos, s in assign.items():
        z = layout.zone(pos[0])
        zone_counts[z] += 1
        if s.score is not None:
            zone_scores[z].append(s.score)
            if s.score >= high_cut:
                high_by_zone[z] += 1
                high_total += 1
            if s.score <= low_cut:
                low_by_zone[z] += 1
                low_total += 1
        if s.is_focus():
            focus_total += 1
            if z == "前排" or layout.is_near_aisle(pos):
                focus_good += 1

    zone_avgs = {z: (sum(v) / len(v) if v else overall) for z, v in zone_scores.items()}
    valid_avgs = list(zone_avgs.values())
    max_diff = (max(valid_avgs) - min(valid_avgs)) if valid_avgs else 0.0

    # 异性同桌比例
    hetero = same = 0
    for a, b in layout.desk_pairs():
        sa, sb = assign.get(a), assign.get(b)
        if sa is None or sb is None or not sa.gender or not sb.gender:
            continue
        if sa.gender != sb.gender:
            hetero += 1
        else:
            same += 1

    # 仍存在的黑名单相邻（拖拽换座后可能重新出现，所以每次都重算）
    bans = _ban_pairs(students)
    ban_hits = 0
    for a, b in layout.adjacent_pairs():
        sa, sb = assign.get(a), assign.get(b)
        if sa is not None and sb is not None and frozenset((sa.name.strip(), sb.name.strip())) in bans:
            ban_hits += 1

    # -- 一句话结论（给老师挑方案用） --
    if ban_hits:
        verdict = f"注意：还有 {ban_hits} 处黑名单相邻，建议手动拖拽或改用其它方案。"
    elif max_diff <= max(3.0, overall * 0.06):
        verdict = "各区成绩比较均衡，可以直接使用。"
    elif zone_avgs["前排"] > zone_avgs["后排"] + max(3.0, overall * 0.06):
        verdict = "前排平均分偏高，若想照顾后排可适当拖拽调整。"
    else:
        verdict = "后排平均分偏高，可看看重点关注学生是否都安排到了前排。"

    lines = [
        f"◆ {plan.name}｜{plan.description}",
        f"班级平均分 {overall:.1f} 分（有效成绩 {len(all_scores)} 人；"
        f"缺考/无总分 {len(students) - len(all_scores)} 人不计入平均分）",
        "前排 / 中排 / 后排 平均分：" +
        " / ".join(f"{zone_avgs[z]:.1f}" for z in ZONE_NAMES) +
        f"｜最大差 {max_diff:.1f} 分",
        f"优生 {high_total} 人：前排 {high_by_zone['前排']}、中排 {high_by_zone['中排']}、后排 {high_by_zone['后排']}",
        f"学困生 {low_total} 人：前排 {low_by_zone['前排']}、中排 {low_by_zone['中排']}、后排 {low_by_zone['后排']}",
        f"重点关注 {focus_total} 人，其中 {focus_good} 人坐在前排或靠过道",
        f"异性同桌 {hetero} 对 / 同性同桌 {same} 对｜黑名单相邻 {ban_hits} 处｜"
        f"未入座 {len(plan.unassigned)} 人",
        "结论：" + verdict,
    ]
    if plan.conflicts:
        lines.append("提示：" + plan.conflicts[0])

    return {
        "plan_name": plan.name,
        "overall_avg": overall,
        "zone_avgs": zone_avgs,
        "zone_counts": zone_counts,
        "max_diff": max_diff,
        "high_total": high_total, "low_total": low_total,
        "high_by_zone": high_by_zone, "low_by_zone": low_by_zone,
        "focus_total": focus_total, "focus_good": focus_good,
        "hetero": hetero, "same": same,
        "ban_hits": ban_hits,
        "unassigned": len(plan.unassigned),
        "verdict": verdict,
        "text": "\n".join(lines),
    }


def attach_stats(plans: List[Plan], layout: ClassroomLayout,
                 students: List[Student]) -> None:
    """生成方案或拖拽换座后调用：给每套方案刷新统计。"""
    for p in plans:
        p.stats = evaluate_plan(p, layout, students)


def compare_rows(plans: List[Plan]) -> Tuple[List[str], List[list]]:
    """所有方案的横向对比表（给「导出统计报表」用）。"""
    columns = ["方案", "策略说明", "前排均分", "中排均分", "后排均分",
               "区均分最大差", "优生人数", "学困生人数", "重点关注人数",
               "重点靠前/过道", "异性同桌对数", "黑名单相邻", "未入座人数", "结论"]
    rows = []
    for p in plans:
        st = p.stats or {}
        za = st.get("zone_avgs", {z: 0.0 for z in ZONE_NAMES})
        rows.append([
            p.name, p.description,
            round(za.get("前排", 0), 1), round(za.get("中排", 0), 1),
            round(za.get("后排", 0), 1), round(st.get("max_diff", 0), 1),
            st.get("high_total", 0), st.get("low_total", 0),
            st.get("focus_total", 0), st.get("focus_good", 0),
            st.get("hetero", 0), st.get("ban_hits", 0),
            st.get("unassigned", 0), st.get("verdict", ""),
        ])
    return columns, rows
