# -*- coding: utf-8 -*-
"""
数据模型模块（models.py）
------------------------------------------------------------------
只存放全局共用的数据结构与常量，不依赖任何其它业务模块，
这样 excel / 布局 / 算法 / 界面 / 导出 五个模块都可以安全地引用它。

设计要点：
1. 学生「总分」允许为 None（缺考、未填、全部科目缺考），
   任何地方都不得把 None 当成 0 分参与统计或排座。
2. 学生「特殊标记」是一段自由文本，只要包含预设关键词就视为
   「重点关注学生」，界面上用浅红色方格标出，方便老师一眼看到。
3. 所有对象都尽量提供 clone()/to_*() 方法，避免界面层直接改动
   算法层的数据（拖拽换座只改座位分配，不改学生本身）。
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

# ----------------------------------------------------------------------
# 一、全局常量
# ----------------------------------------------------------------------

# 这些文本一律视为「缺失值」，而不是 0 分。比较时会先 strip 再转小写。
MISSING_TOKENS = {
    "", "-", "－", "—", "–", "――", "/", "\\", "n/a", "na", "none",
    "null", "nan", "缺考", "缺", "缺席", "未考", "未参加", "无", "空",
    "暂无", "absent", "null值", "—", "?", "？",
}

# 特殊标记里出现以下任一关键词，即判定为「重点关注」学生（浅红色方格）
FOCUS_KEYWORDS = (
    "重点", "关注", "特殊", "学困", "需帮", "帮扶", "视力", "听力",
    "多动", "留级", "随班", "照顾", "违纪", "薄弱",
)

# 统一浅色配色（与浅色桌面主题一致，投影/打印也清晰）
COLOR = {
    "focus": "#FADBD8",          # 重点关注学生：浅红
    "focus_border": "#E3A19A",
    "normal": "#FFFFFF",         # 普通学生：白色
    "normal_border": "#B6C2CE",
    "empty": "#F2F4F6",          # 空座位：很浅的灰
    "empty_border": "#CBD3DB",
    "block": "#C7CDD3",          # 过道 / 禁用空位：灰色
    "block_border": "#AAB3BC",
    "podium": "#5A6B7B",         # 讲台：深灰蓝
    "podium_text": "#FFFFFF",
    "text": "#1F2933",
    "subtext": "#687582",
    "aisle_hint": "#DCE6F1",
    "ghost": "#FFCF5C",          # 拖拽虚影
    "select": "#3B82F6",         # 选中 / 框选高亮
    "board": "#F7F9FB",          # 画布底色
}

# 告警类别（导入预览窗口按类别分色显示）
WARN_KIND = {
    "empty_name": "空姓名",
    "bad_score": "异常分数",
    "duplicate": "重复学生",
    "absent": "缺考名单",
    "info": "导入说明",
    "conflict": "约束冲突",
}


@dataclass
class WarningItem:
    """一条给老师看的中文提示（不显示任何原始代码异常）"""
    kind: str                 # 见 WARN_KIND 的键
    message: str              # 中文说明
    row_no: int = 0           # 关联的原表行号，0 表示整体说明

    def display(self) -> str:
        prefix = WARN_KIND.get(self.kind, "提示")
        if self.row_no:
            return f"【{prefix}·第{self.row_no}行】{self.message}"
        return f"【{prefix}】{self.message}"


# ----------------------------------------------------------------------
# 二、学生对象
# ----------------------------------------------------------------------
@dataclass
class Student:
    """一名学生。所有字段都允许为空，界面与算法必须容忍脏数据。"""
    sid: str = ""                                   # 学号
    name: str = ""                                  # 姓名
    score: Optional[float] = None                   # 总分；None = 缺失，绝不按 0 计
    gender: str = ""                                # 男 / 女 / 空（未知）
    tag: str = ""                                   # 特殊标记（自由文本）
    forbid: List[str] = field(default_factory=list)  # 禁止同桌 / 相邻的学生姓名
    subjects: Dict[str, Optional[float]] = field(default_factory=dict)  # 各科成绩
    source_row: int = 0                             # 原始 Excel 行号，便于老师核对

    # -- 判定 ----------------------------------------------------------
    def is_focus(self) -> bool:
        """是否为重点关注学生（浅红色）。只看特殊标记文本，规则简单可预期。"""
        t = (self.tag or "").strip()
        return any(k in t for k in FOCUS_KEYWORDS)

    def short_tag(self) -> str:
        """方格内显示的简短标记，太长会遮挡姓名。"""
        t = (self.tag or "").strip()
        if not t:
            return ""
        for k in FOCUS_KEYWORDS:
            if k in t:
                return k
        return t[:3]

    def forbid_text(self) -> str:
        return "、".join(self.forbid) if self.forbid else ""

    def score_text(self) -> str:
        return "缺考/未填" if self.score is None else f"{self.score:g}"

    def identity(self) -> str:
        """用于去重 / 黑名单匹配的唯一标识，优先学号，其次姓名。"""
        return (self.sid or self.name).strip()

    def clone(self) -> "Student":
        return Student(
            sid=self.sid, name=self.name, score=self.score, gender=self.gender,
            tag=self.tag, forbid=list(self.forbid), subjects=dict(self.subjects),
            source_row=self.source_row,
        )

    def __hash__(self) -> int:
        # 允许把 Student 放进 set / 作为字典键使用；内容相同的两个对象视为不同人时
        # 由 identity() 区分，这里只按身份哈希，避免可变字段影响哈希值。
        return hash((self.sid, self.name, self.source_row))
