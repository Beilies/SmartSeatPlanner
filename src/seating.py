# -*- coding: utf-8 -*-
"""
排座算法模块（seating.py）
==================================================================
【单独说明算法逻辑，方便后续改规则】

一、整体思路
    本模块不用「一次性贪心排完就结束」的做法，而是：
        初始解（体现方案差异） -> 成本函数打分 -> 随机交换 + 模拟退火局部优化
    这样既能做出 3~5 套风格不同的方案，又能保证每套方案都尽量满足硬约束。

二、成本函数 cost（越小越好）—— 想改规则，改 _cost() 里对应项即可
    1. 分区均衡     ：前排/中排/后排的「标准化平均分」差异越大，罚分越大。
                      标准化 = (个人总分 - 班级均分) / 标准差，避免不同科目满分不同带来的量纲问题。
    2. 优生打散     ：相邻两个座位都是优生（分数前 25%）时罚分，权重由左侧滑块控制。
                      滑块拉到最大 = 坚决不让优生扎堆；拉到最小 = 基本不管扎堆。
    3. 强弱搭配/同水平：
                      help 策略奖励同桌分差大（一强一弱，方便帮扶）；
                      similar 策略惩罚同桌分差大（水平接近，方便分层教学）；
                      其余策略不干预同桌分差。
    4. 黑名单       ：禁止同桌名单里的两人出现在上下左右相邻座位时重罚（权重远大于其它项）。
                      同时额外做一次「强制修复」，实在修不好就记录冲突并给老师降级建议，不崩溃。
    5. 性别均衡     ：开关打开时，同桌同性别要给一点罚分（只约束同桌，不强行拆散整排）。
    6. 特殊学生位置 ：重点关注学生坐得越靠前、越靠过道，成本越低（相当于奖励）。
                      front/aisle 两个指标都由布局模块计算，改布局即自动生效。
    7. 缺考/无总分  ：这类学生标准化分取 0（班级平均水平），绝不按 0 分拉低班级均分。

三、5 套方案怎么做到「差异化」
    方案1 强弱帮扶   ：高低分交替 + 同桌对优先，优化时奖励大分差同桌
    方案2 同水平就近 ：按分数排序后轮转分到三个区，区内水平接近，分区均分又均衡
    方案3 均衡随机   ：随机初始解 + 强分区均衡，座位关系最"松"
    方案4 性别优先   ：男女交替初始解 + 加大的性别权重（性别开关关闭时退化为均衡随机）
    方案5 特殊学生优先：重点关注学生先占「靠前 + 靠过道」最佳位置，其余随机 + 均衡
    每套方案用不同随机种子；万一生成结果与前面某套完全一样，会自动换种子重跑。

四、约束冲突时的降级策略（全部转成中文提示，绝不抛异常给老师）
    - 学生数 > 座位数   ：优先安排有特殊标记的学生，其余放进「未安排名单」，提示扩大行列数
    - 学生数 < 座位数   ：只启用「从前往后、从左到右」的前 N 个座位，
                          空位统一留在最后几排的末尾，绝不夹在学生中间
    - 可用座位为 0      ：抛 SeatingError，主界面弹窗建议打开布局编辑器
    - 黑名单无解        ：先尝试强制交换修复；仍冲突则保留方案 + 列出冲突名单 + 给出放宽建议
    - 名单里的名字找不到：不算错误，只在提示里说明「未匹配到该学生」
    - 班级人数太少/分数全相同：标准差按 1 处理，避免除零

五、扩展预留（仅注释，暂不实现）
    【定时座位轮换】可在 Plan 上增加 effective_date / next_rotate 字段，
      按 layout 的列分组做整体左右平移，复用 _optimize() 做微调即可。
    【小组功能】给座位增加 group_id，成本函数里加一项「同组学生尽量相邻」
      （w_group * 同组相邻对数取负），再把小组名显示到方格角标即可。
"""
from __future__ import annotations

import math
import random
import statistics
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from layout import ClassroomLayout, Pos, ZONE_NAMES
from models import Student

# 单元格坐标别名
Assignment = Dict[Pos, Student]


class SeatingError(Exception):
    """排座无法进行（例如没有座位）。GUI 捕获后给中文提示。"""


@dataclass
class SeatingOptions:
    """左侧参数区的设置。"""
    gender_balance: bool = True        # 性别均衡开关
    use_blacklist: bool = True         # 启用同桌黑名单开关
    spread_weight: float = 0.5         # 优生打散权重 0~1（滑块 0~100 换算）
    plan_count: int = 3                # 生成方案数量 3~5
    seed: int = 20261107               # 固定种子，保证同一份数据每次生成结果一致


@dataclass
class Plan:
    """一套座位方案。stats 由 evaluate 模块填充。"""
    name: str
    strategy: str
    description: str
    assignment: Assignment = field(default_factory=dict)
    unassigned: List[Student] = field(default_factory=list)
    conflicts: List[str] = field(default_factory=list)
    stats: Optional[dict] = None

    def clone_assignment(self) -> Assignment:
        return dict(self.assignment)


# 每套方案的策略配置（改这里 = 改方案风格）
STRATEGIES: List[dict] = [
    {"name": "方案1 强弱帮扶", "key": "help",
     "desc": "一强一弱同桌、优生分散，适合想让同学互相带一带",
     "pair_mode": "help", "w_zone": 2.2, "w_pair": 0.9, "w_gender": 1.0, "w_special": 1.2},
    {"name": "方案2 同水平就近", "key": "similar",
     "desc": "水平接近的同学坐得近，方便分层提问与分组辅导",
     "pair_mode": "similar", "w_zone": 2.6, "w_pair": 1.0, "w_gender": 1.0, "w_special": 1.0},
    {"name": "方案3 均衡随机", "key": "random",
     "desc": "分区成绩最均衡、同桌关系最随机，适合常规月考座次",
     "pair_mode": "none", "w_zone": 3.4, "w_pair": 0.0, "w_gender": 0.8, "w_special": 0.9},
    {"name": "方案4 性别优先均衡", "key": "gender",
     "desc": "尽量男女同桌，同时保证各区平均分接近",
     "pair_mode": "none", "w_zone": 2.4, "w_pair": 0.0, "w_gender": 3.0, "w_special": 1.0},
    {"name": "方案5 特殊学生优先", "key": "front",
     "desc": "重点关注学生优先靠前、靠过道，方便老师随时观察",
     "pair_mode": "help", "w_zone": 2.2, "w_pair": 0.4, "w_gender": 0.8, "w_special": 4.0},
]


# ----------------------------------------------------------------------
# 上下文：把 repeated 计算预先做好，优化循环里只做纯数值运算
# ----------------------------------------------------------------------
class _Context:
    def __init__(self, students: List[Student], layout: ClassroomLayout,
                 options: SeatingOptions):
        self.layout = layout
        self.options = options
        self.seats = layout.seats()
        self.desk_pairs = layout.desk_pairs()
        self.adjacent_pairs = layout.adjacent_pairs()

        # 人比座位少时，只在前 N 个座位里排，空位统一留在后排末尾。
        # 优化器只在「已入座」的座位之间交换，所以空位位置一旦定下就不会再散开。
        self.used_seats = _front_seats(self.seats, len(students))
        used_set = set(self.used_seats)
        self.used_desk_pairs = [pair for pair in self.desk_pairs
                                if pair[0] in used_set and pair[1] in used_set]

        scores = [s.score for s in students if s.score is not None]
        self.mean = statistics.fmean(scores) if scores else 0.0
        try:
            self.std = statistics.pstdev(scores) if len(scores) > 1 else 1.0
        except statistics.StatisticsError:
            self.std = 1.0
        if self.std < 1e-6:
            self.std = 1.0  # 全班同分，避免除零

        # 优生 / 学困生阈值：分数前 25% / 后 25%（无分数者不算）
        ordered = sorted(scores)
        if ordered:
            self.high_cut = ordered[int(0.75 * (len(ordered) - 1))]
            self.low_cut = ordered[int(0.25 * (len(ordered) - 1))]
        else:
            self.high_cut, self.low_cut = float("inf"), float("-inf")

        self.norm: Dict[str, float] = {}       # 姓名 -> 标准化分
        self.high: Dict[str, bool] = {}
        self.low: Dict[str, bool] = {}
        self.focus: Dict[str, bool] = {}
        self.gender: Dict[str, str] = {}
        for s in students:
            key = s.name
            self.norm[key] = 0.0 if s.score is None else (s.score - self.mean) / self.std
            self.high[key] = s.score is not None and s.score >= self.high_cut
            self.low[key] = s.score is not None and s.score <= self.low_cut
            self.focus[key] = s.is_focus()
            self.gender[key] = s.gender

        # 黑名单：把「姓名对」放进集合，双向都能查
        self.ban_pairs = set()
        by_name: Dict[str, Student] = {}
        for s in students:
            by_name.setdefault(s.name.strip(), s)
        self.unmatched_bans: List[str] = []
        for s in students:
            for target in s.forbid:
                t = by_name.get(target.strip())
                if t is None or t.name == s.name:
                    self.unmatched_bans.append(f"{s.name} → {target}")
                    continue
                self.ban_pairs.add(frozenset((s.name.strip(), t.name.strip())))

        # 特殊学生位置的「理想程度」预计算（越靠前、越靠过道越接近 1）
        self.position_value: Dict[Pos, float] = {}
        for pos in self.seats:
            front = layout.front_ratio(pos[0])
            aisle = 1.0 if layout.is_near_aisle(pos) else 0.0
            self.position_value[pos] = 0.6 * front + 0.4 * aisle

    def is_banned(self, a: Student, b: Student) -> bool:
        if not self.options.use_blacklist:
            return False
        return frozenset((a.name.strip(), b.name.strip())) in self.ban_pairs


# ----------------------------------------------------------------------
# 成本函数
# ----------------------------------------------------------------------
def _cost(assign: Assignment, ctx: _Context, cfg: dict) -> float:
    """计算一套座位方案的「不满意程度」，数值越小越好。"""
    cost = 0.0
    spread_w = 0.3 + 3.0 * ctx.options.spread_weight   # 优生打散权重（滑块）

    # ① 分区成绩均衡：三个区的标准化均分尽量都接近 0（班级平均）
    zone_sum = {z: 0.0 for z in ZONE_NAMES}
    zone_cnt = {z: 0 for z in ZONE_NAMES}
    for pos, s in assign.items():
        z = ctx.layout.zone(pos[0])
        zone_sum[z] += ctx.norm.get(s.name, 0.0)
        zone_cnt[z] += 1
    for z in ZONE_NAMES:
        avg = zone_sum[z] / zone_cnt[z] if zone_cnt[z] else 0.0
        cost += cfg["w_zone"] * abs(avg)

    # ② 优生打散：上下左右相邻的两个座位都是优生 -> 罚分
    for a, b in ctx.adjacent_pairs:
        sa, sb = assign.get(a), assign.get(b)
        if sa is None or sb is None:
            continue
        if ctx.high.get(sa.name) and ctx.high.get(sb.name):
            cost += spread_w * 1.2

    # ③ 同桌搭配偏好（help 奖励分差 / similar 惩罚分差）
    mode = cfg["pair_mode"]
    if mode != "none" and cfg["w_pair"] > 0:
        for a, b in ctx.desk_pairs:
            sa, sb = assign.get(a), assign.get(b)
            if sa is None or sb is None:
                continue
            diff = min(3.0, abs(ctx.norm.get(sa.name, 0.0) - ctx.norm.get(sb.name, 0.0)))
            if mode == "help":
                cost -= cfg["w_pair"] * diff * 0.35     # 分差越大越划算（帮扶）
            elif mode == "similar":
                cost += cfg["w_pair"] * diff * 0.35     # 分差越小越划算（同水平）
            # ④ 黑名单：同桌/相邻重罚（权重远高于其它软约束）
            if ctx.is_banned(sa, sb):
                cost += 40.0
            # ⑤ 性别均衡：同桌同性别轻微罚分
            if ctx.options.gender_balance and sa.gender and sa.gender == sb.gender:
                cost += 1.4 * cfg["w_gender"]
    else:
        # 即使不评价分差，黑名单与性别约束仍然要检查
        for a, b in ctx.desk_pairs:
            sa, sb = assign.get(a), assign.get(b)
            if sa is None or sb is None:
                continue
            if ctx.is_banned(sa, sb):
                cost += 40.0
            if ctx.options.gender_balance and sa.gender and sa.gender == sb.gender:
                cost += 1.4 * cfg["w_gender"]

    # 相邻但不同桌的黑名单同样禁止（前后排相邻也算「相邻」）
    if ctx.options.use_blacklist and ctx.ban_pairs:
        for a, b in ctx.adjacent_pairs:
            sa, sb = assign.get(a), assign.get(b)
            if sa is None or sb is None:
                continue
            if ctx.is_banned(sa, sb):
                cost += 12.0     # 同桌对已在上面重罚，这里避免重复计满权重

    # ⑥ 特殊学生优先靠前、靠过道：位置越理想，成本越低
    w_special = cfg["w_special"]
    if w_special:
        for pos, s in assign.items():
            if ctx.focus.get(s.name):
                cost -= w_special * ctx.position_value.get(pos, 0.0)

    return cost


def _swap(assign: Assignment, a: Pos, b: Pos) -> None:
    """交换两个座位上的学生（允许其中一个是空座位）。"""
    va, vb = assign.get(a), assign.get(b)
    if va is None and vb is None:
        return
    if va is None:
        assign[a] = vb
        assign.pop(b, None)
    elif vb is None:
        assign[b] = va
        assign.pop(a, None)
    else:
        assign[a], assign[b] = vb, va


def _banned_conflicts(assign: Assignment, ctx: _Context) -> List[Tuple[Pos, Pos]]:
    """列出仍然违反黑名单的相邻座位对。"""
    out = []
    for a, b in ctx.adjacent_pairs:
        sa, sb = assign.get(a), assign.get(b)
        if sa is not None and sb is not None and ctx.is_banned(sa, sb):
            out.append((a, b))
    return out


def _repair_blacklist(assign: Assignment, ctx: _Context, cfg: dict,
                      rng: random.Random, max_rounds: int = 12) -> int:
    """
    黑名单强制修复：对每一对冲突，尝试把其中一人换到别处。
    返回仍存在的冲突数量。修不好也不抛错，由上层转成中文降级提示。
    """
    occupied = [p for p in ctx.seats if p in assign]
    if len(occupied) < 2 or not ctx.ban_pairs:
        return len(_banned_conflicts(assign, ctx))
    for _ in range(max_rounds):
        conflicts = _banned_conflicts(assign, ctx)
        if not conflicts:
            return 0
        a, b = conflicts[0]
        base = _cost(assign, ctx, cfg)
        best, best_cost = None, base
        candidates = rng.sample(occupied, min(len(occupied), 30))
        for other in candidates:
            if other in (a, b):
                continue
            for source in (a, b):
                _swap(assign, source, other)
                c = _cost(assign, ctx, cfg)
                _swap(assign, source, other)   # 还原
                if c < best_cost - 1e-9:
                    best, best_cost = (source, other), c
        if best is None:
            break
        _swap(assign, best[0], best[1])
    return len(_banned_conflicts(assign, ctx))


# ----------------------------------------------------------------------
# 初始解：不同策略给出不同「起点」，这是方案差异化的主要来源
# ----------------------------------------------------------------------
def _front_seats(seats: List[Pos], count: int) -> List[Pos]:
    """只取「从前往后、从左到右」的前 count 个座位。

    【为什么不做均匀留白】早期版本会等间隔抽座位，导致学生中间夹着一个个
    「空位」，看起来像漏排了人，老师也不好打印发家长。现在人少于座位时，
    统一把多出来的座位留在最后几排末尾：学生连成一片，空位集中且不显眼。
    seats 必须已按 (行, 列) 排好序（ClassroomLayout.seats() 天然如此）。
    """
    if count >= len(seats):
        return list(seats)
    if count <= 0:
        return []
    return list(seats[:count])


def _initial_assignment(strategy: str, students: List[Student],
                        seats: List[Pos], ctx: _Context,
                        rng: random.Random) -> Assignment:
    """生成初始解。想调整某套方案的编排习惯，改这里即可。"""
    assign: Assignment = {}
    n = len(students)
    if n == 0 or not seats:
        return assign

    by_score = sorted(students, key=lambda s: (s.score is None, -(s.score or 0)))

    if strategy == "help":
        # 高低分交替：一强一弱尽量落在同一同桌对上
        half = n // 2
        high = by_score[:half]
        low = by_score[half:][::-1]      # 低分组倒序，让最强配最弱
        paired: List[Student] = []
        for i in range(max(len(high), len(low))):
            if i < len(high):
                paired.append(high[i])
            if i < len(low):
                paired.append(low[i])
        order_seats: List[Pos] = []
        for a, b in ctx.used_desk_pairs:
            order_seats.extend([a, b])
        for p in seats:
            if p not in order_seats:
                order_seats.append(p)
        for p, s in zip(order_seats, paired):
            assign[p] = s
        return assign

    if strategy == "similar":
        # 按分数排序后轮转分到三个区：区内水平接近，区间平均分又均衡
        zone_seats: Dict[str, List[Pos]] = {z: [] for z in ZONE_NAMES}
        for p in seats:
            zone_seats[ctx.layout.zone(p[0])].append(p)
        zone_students: Dict[str, List[Student]] = {z: [] for z in ZONE_NAMES}
        for i, s in enumerate(by_score):
            zone_students[ZONE_NAMES[i % 3]].append(s)
        for z in ZONE_NAMES:
            for p, s in zip(zone_seats[z], zone_students[z]):
                assign[p] = s
        # 若某个区座位不够（例如过道把前排切掉很多），把剩余学生放到空座位
        used_ids = {id(s) for s in assign.values()}
        left = [s for s in students if id(s) not in used_ids]
        free = [p for p in seats if p not in assign]
        for p, s in zip(free, left):
            assign[p] = s
        return assign

    if strategy == "gender":
        # 男女交替；未知性别的排在后面
        boys = [s for s in students if s.gender == "男"]
        girls = [s for s in students if s.gender == "女"]
        unknown = [s for s in students if s.gender not in ("男", "女")]
        rng.shuffle(boys); rng.shuffle(girls); rng.shuffle(unknown)
        merged: List[Student] = []
        for i in range(max(len(boys), len(girls))):
            if i < len(girls):
                merged.append(girls[i])
            if i < len(boys):
                merged.append(boys[i])
        merged.extend(unknown)
        for p, s in zip(_front_seats(seats, len(merged)), merged):
            assign[p] = s
        return assign

    if strategy == "front":
        # 重点关注学生先占「靠前 + 靠过道」的最佳座位
        focus = [s for s in students if s.is_focus()]
        others = [s for s in students if not s.is_focus()]
        rng.shuffle(others)
        best_seats = sorted(seats, key=lambda p: (-ctx.position_value.get(p, 0.0), p))
        used = set()
        for s, p in zip(focus, best_seats):
            assign[p] = s
            used.add(p)
        free = [p for p in _front_seats([p for p in seats if p not in used], len(others))]
        for p, s in zip(free, others):
            assign[p] = s
        return assign

    # random（默认）：完全随机起点，交给优化器做分区均衡与优生打散
    shuffled = list(students)
    rng.shuffle(shuffled)
    for p, s in zip(_front_seats(seats, len(shuffled)), shuffled):
        assign[p] = s
    return assign


def _optimize(assign: Assignment, ctx: _Context, cfg: dict,
              rng: random.Random, iterations: int) -> None:
    """
    模拟退火局部优化：随机挑两个有人坐的座位交换，成本下降就接受，
    偶尔也接受一点变差的解以跳出局部最优。温度逐步下降。
    """
    occupied = [p for p in ctx.seats if p in assign]
    if len(occupied) < 2:
        return
    cur = _cost(assign, ctx, cfg)
    temp = 0.9
    cooling = (0.02 / temp) ** (1.0 / max(1, iterations))
    for _ in range(iterations):
        a, b = rng.sample(occupied, 2)
        _swap(assign, a, b)
        new = _cost(assign, ctx, cfg)
        delta = new - cur
        if delta <= 0 or rng.random() < math.exp(-delta / max(temp, 1e-6)):
            cur = new
        else:
            _swap(assign, a, b)     # 不接受就还原
        temp *= cooling


# ----------------------------------------------------------------------
# 对外主入口
# ----------------------------------------------------------------------
def generate_plans(students: List[Student], layout: ClassroomLayout,
                   options: SeatingOptions) -> Tuple[List[Plan], List[str]]:
    """
    生成 3~5 套座位方案。
    返回 (方案列表, 全局提示)。全局提示里包含所有降级说明，GUI 会用中文弹窗展示。
    """
    notices: List[str] = []
    clean = [s for s in students if (s.name or "").strip()]
    if not clean:
        raise SeatingError("还没有有效学生，请先导入 Excel 成绩表")

    seats = layout.seats()
    if not seats:
        raise SeatingError("当前教室布局里一个学生座位都没有，请先打开「教室布局编辑器」标记座位")

    # —— 降级 1：人比座位多 ——
    placed = list(clean)
    unassigned: List[Student] = []
    if len(clean) > len(seats):
        # 有特殊标记的学生优先入座，其余保持原表顺序
        priority = [s for s in clean if s.is_focus()]
        rest = [s for s in clean if not s.is_focus()]
        placed = (priority + rest)[:len(seats)]
        placed_ids = {id(s) for s in placed}
        unassigned = [s for s in clean if id(s) not in placed_ids]
        notices.append(
            f"可用座位只有 {len(seats)} 个，但学生有 {len(clean)} 人。"
            f"已优先安排重点关注学生，还有 {len(unassigned)} 人未入座"
            f"（{ '、'.join(s.name for s in unassigned[:6]) }{'…' if len(unassigned) > 6 else ''}）。"
            "建议在布局编辑器里增加行列数，或减少过道格子。")
    elif len(clean) < len(seats):
        notices.append(f"学生 {len(clean)} 人、座位 {len(seats)} 个，"
                       f"多出的 {len(seats) - len(clean)} 个空位统一留在最后几排末尾，"
                       "不会夹在学生中间；老师也可以直接拖拽调整。")

    ctx = _Context(placed, layout, options)
    if ctx.unmatched_bans:
        sample = "、".join(ctx.unmatched_bans[:5])
        notices.append(f"禁止同桌名单里有 {len(ctx.unmatched_bans)} 条没匹配到学生（{sample}…），"
                       "已跳过这些记录；请核对姓名是否与名单完全一致。")
    if options.gender_balance and not any(s.gender for s in placed):
        notices.append("表格里没有可用的性别信息，「性别均衡」开关这次不会起作用（不影响其它规则）。")

    count = max(3, min(5, int(options.plan_count)))
    iterations = min(3000, 700 + 30 * len(placed))
    plans: List[Plan] = []
    signatures = set()

    for i in range(count):
        cfg = dict(STRATEGIES[i])
        key = cfg["key"]
        # 性别开关关闭时，方案4 自动退化成均衡随机版本，避免出现"名不副实"的方案
        if key == "gender" and not options.gender_balance:
            cfg["pair_mode"] = "none"
            cfg["w_gender"] = 0.0
            cfg["w_zone"] = 3.0
        best_assign = None
        for attempt in range(3):       # 结果撞车就换种子重跑
            rng = random.Random(options.seed + i * 7919 + attempt * 104729)
            assign = _initial_assignment(key, placed, ctx.used_seats, ctx, rng)
            _optimize(assign, ctx, cfg, rng, iterations)
            signature = tuple(sorted((s.name, p[0], p[1]) for p, s in assign.items()))
            if signature not in signatures:
                signatures.add(signature)
                best_assign = assign
                break
            if best_assign is None:
                best_assign = assign
        assert best_assign is not None

        # —— 降级 2：黑名单冲突尽力修复，修不好就如实告知 ——
        conflicts: List[str] = []
        remain = _repair_blacklist(best_assign, ctx, cfg, random.Random(options.seed + i + 1))
        if remain:
            pairs = _banned_conflicts(best_assign, ctx)[:4]
            detail = "；".join(
                f"{best_assign[a].name} 与 {best_assign[b].name}" for a, b in pairs
                if a in best_assign and b in best_assign)
            conflicts.append(f"仍有 {remain} 处黑名单相邻（{detail}）。"
                             "建议：减少禁止同桌名单人数、在两人之间加一格过道，"
                             "或生成更多方案后手动拖拽调整。")
            notices.append(f"{cfg['name']}：{conflicts[-1]}")

        plans.append(Plan(
            name=cfg["name"], strategy=key, description=cfg["desc"],
            assignment=best_assign, unassigned=list(unassigned),
            conflicts=conflicts,
        ))

    return plans, notices


def apply_swap(plan: Plan, a: Pos, b: Pos) -> None:
    """供 GUI 拖拽换座调用：交换两个座位（允许与空座位交换）。"""
    _swap(plan.assignment, a, b)
    plan.stats = None      # 统计失效，界面会重新计算
