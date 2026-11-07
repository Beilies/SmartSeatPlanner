# -*- coding: utf-8 -*-
"""
教室布局模块（layout.py）
------------------------------------------------------------------
用二维字符网格描述教室：
    seat   = 学生座位（可以坐人）
    block  = 过道 / 垃圾桶 / 禁用空位（灰色，不能坐人）
    podium = 讲台所在格（深灰色，不能坐人；画布顶部另有整条讲台示意）

布局可以保存为 .json 模板，下次一键加载，老师不用重复点格子。

【扩展预留·仅注释，暂不实现】
    定时座位轮换：可在本模块增加 "rotation_group" 标记，
    把若干列划为同一轮换组，配合定时器整体平移 assignment 即可。
    小组功能：可给每个 seat 增加 group_id 字段，
    排座算法把「同组就近」加入成本函数即可，无需改动界面层。
"""
from __future__ import annotations

import json
import os
import re
from typing import Dict, Iterable, List, Tuple

# 单元格三种取值
SEAT = "seat"
BLOCK = "block"
PODIUM = "podium"
VALID_CELLS = (SEAT, BLOCK, PODIUM)

CELL_LABEL = {SEAT: "学生座位", BLOCK: "过道/禁用", PODIUM: "讲台"}
ZONE_NAMES = ("前排", "中排", "后排")

Pos = Tuple[int, int]  # (行, 列)，均从 0 开始


class LayoutError(Exception):
    """布局文件损坏 / 参数非法。上层 GUI 会捕获并转成中文弹窗。"""


class ClassroomLayout:
    """教室布局对象。所有坐标一律 (row, col)，从 0 开始。"""

    MIN_SIZE = 2
    MAX_SIZE = 30

    def __init__(self, rows: int = 6, cols: int = 9,
                 cells: List[List[str]] | None = None,
                 name: str = "默认教室"):
        self.name = name or "默认教室"
        self.rows = self._check_size(rows)
        self.cols = self._check_size(cols)
        if cells is None:
            self.cells = [[SEAT for _ in range(self.cols)] for _ in range(self.rows)]
        else:
            self.cells = self._normalize_cells(cells)
            self.rows, self.cols = len(self.cells), len(self.cells[0])

    # ------------------------------------------------------------------
    # 构造与校验
    # ------------------------------------------------------------------
    @staticmethod
    def _check_size(n: int) -> int:
        try:
            n = int(n)
        except (TypeError, ValueError):
            raise LayoutError("行列数必须是整数")
        if not (ClassroomLayout.MIN_SIZE <= n <= ClassroomLayout.MAX_SIZE):
            raise LayoutError(f"行列数请设置在 {ClassroomLayout.MIN_SIZE}~{ClassroomLayout.MAX_SIZE} 之间")
        return n

    def _normalize_cells(self, cells: List[List[str]]) -> List[List[str]]:
        """加载外部文件时做容错：非法字符一律降级为座位，行列不齐自动补齐。"""
        if not isinstance(cells, list) or not cells:
            raise LayoutError("布局文件内容为空")
        width = max((len(r) if isinstance(r, list) else 0) for r in cells)
        if width == 0:
            raise LayoutError("布局文件里没有有效的格子")
        width = min(max(width, self.MIN_SIZE), self.MAX_SIZE)
        out: List[List[str]] = []
        for row in cells[:self.MAX_SIZE]:
            safe = []
            for c in range(width):
                v = row[c] if isinstance(row, list) and c < len(row) else SEAT
                safe.append(v if v in VALID_CELLS else SEAT)
            out.append(safe)
        if not out:
            raise LayoutError("布局文件里没有有效的格子")
        return out

    @classmethod
    def default(cls, rows: int = 7, cols: int = 11) -> "ClassroomLayout":
        """
        默认教室布局（老师指定的常用教室）：
            7 行 × 11 列，3 个座位 | 过道 | 3 个座位 | 过道 | 3 个座位
            即第 4 列与第 8 列（从 1 数起）是纵向过道，共 7×9 = 63 个座位。
            画布顶部另有整条「讲台（前方）」示意，不占用网格格子。

        传入其它行列数时保持通用规则：宽教室 (>=6 列) 在中间留一条过道，
        这样旧模板 / 自检脚本指定 6×9 之类的尺寸时行为不变。
        """
        lay = cls(rows=rows, cols=cols, name="默认教室")
        if cols == 11:
            aisle_cols = (3, 7)          # 0 起算，对应界面上的第 4、8 列
        elif cols >= 6:
            aisle_cols = (cols // 2,)
        else:
            aisle_cols = ()
        for r in range(lay.rows):
            for c in aisle_cols:
                lay.cells[r][c] = BLOCK
        return lay

    def resize(self, rows: int, cols: int, preserve: bool = True) -> None:
        """改变行列数；preserve=True 时尽量保留左上角已有格子。"""
        rows, cols = self._check_size(rows), self._check_size(cols)
        old = [list(r) for r in self.cells]
        self.rows, self.cols = rows, cols
        self.cells = [[SEAT for _ in range(cols)] for _ in range(rows)]
        if preserve:
            for r in range(min(rows, len(old))):
                for c in range(min(cols, len(old[r]))):
                    self.cells[r][c] = old[r][c]

    def set_cell(self, row: int, col: int, value: str) -> None:
        if not (0 <= row < self.rows and 0 <= col < self.cols):
            return
        self.cells[row][col] = value if value in VALID_CELLS else SEAT

    def get(self, row: int, col: int) -> str:
        if 0 <= row < self.rows and 0 <= col < self.cols:
            return self.cells[row][col]
        return BLOCK

    def fill_all(self, value: str = SEAT) -> None:
        self.cells = [[value if value in VALID_CELLS else SEAT
                       for _ in range(self.cols)] for _ in range(self.rows)]

    # ------------------------------------------------------------------
    # 几何查询（供排座算法 / 导出 / 绘图共用）
    # ------------------------------------------------------------------
    def seats(self) -> List[Pos]:
        """按「从前往后、从左到右」返回所有可坐座位。"""
        return [(r, c) for r in range(self.rows) for c in range(self.cols)
                if self.cells[r][c] == SEAT]

    def desk_pairs(self) -> List[Tuple[Pos, Pos]]:
        """
        同桌对：同一行里连续两个 seat 组成一对，遇到过道/讲台就断开。
        例：座位 座位 过道 座位 座位 座位 座位 -> (0,1)、(3,4)、(5,6)
        这也是「强弱帮扶」「同水平就近」策略做初始配对时的依据。
        """
        pairs: List[Tuple[Pos, Pos]] = []
        for r in range(self.rows):
            run: List[int] = []
            for c in range(self.cols + 1):
                is_seat = c < self.cols and self.cells[r][c] == SEAT
                if is_seat:
                    run.append(c)
                else:
                    for i in range(0, len(run) - 1, 2):
                        pairs.append(((r, run[i]), (r, run[i + 1])))
                    run = []
        return pairs

    def deskmate(self, pos: Pos) -> Pos | None:
        """返回该座位的同桌座位（没有同桌返回 None）。"""
        for a, b in self.desk_pairs():
            if pos == a:
                return b
            if pos == b:
                return a
        return None

    def neighbors(self, pos: Pos) -> List[Pos]:
        """上下左右四邻中可坐人的座位；黑名单约束就检查这四个方向。"""
        r, c = pos
        out = []
        for dr, dc in ((-1, 0), (1, 0), (0, -1), (0, 1)):
            nr, nc = r + dr, c + dc
            if 0 <= nr < self.rows and 0 <= nc < self.cols and self.cells[nr][nc] == SEAT:
                out.append((nr, nc))
        return out

    def adjacent_pairs(self) -> List[Tuple[Pos, Pos]]:
        """所有共享边的座位对（每个对只出现一次）。"""
        pairs = []
        for r in range(self.rows):
            for c in range(self.cols):
                if self.cells[r][c] != SEAT:
                    continue
                if c + 1 < self.cols and self.cells[r][c + 1] == SEAT:
                    pairs.append(((r, c), (r, c + 1)))
                if r + 1 < self.rows and self.cells[r + 1][c] == SEAT:
                    pairs.append(((r, c), (r + 1, c)))
        return pairs

    def zone(self, row: int) -> str:
        """行号 -> 前排 / 中排 / 后排（按行数三等分）。"""
        if self.rows <= 0:
            return ZONE_NAMES[0]
        idx = min(2, (row * 3) // self.rows)
        return ZONE_NAMES[idx]

    def zone_rows(self) -> Dict[str, List[int]]:
        out = {z: [] for z in ZONE_NAMES}
        for r in range(self.rows):
            out[self.zone(r)].append(r)
        return out

    def front_ratio(self, row: int) -> float:
        """越靠前越接近 1，最后一排为 0。"""
        if self.rows <= 1:
            return 1.0
        return 1.0 - row / (self.rows - 1)

    def is_near_aisle(self, pos: Pos) -> bool:
        """是否靠过道：贴着教室左右边界，或左右相邻格是过道/讲台。"""
        r, c = pos
        if c == 0 or c == self.cols - 1:
            return True
        return (self.cells[r][c - 1] != SEAT) or (self.cells[r][c + 1] != SEAT)

    def summary(self) -> str:
        n_seat = len(self.seats())
        n_block = sum(v == BLOCK for row in self.cells for v in row)
        n_podium = sum(v == PODIUM for row in self.cells for v in row)
        return (f"{self.rows} 行 × {self.cols} 列｜可用座位 {n_seat} 个｜"
                f"过道/禁用 {n_block} 格｜讲台 {n_podium} 格｜同桌 {len(self.desk_pairs())} 对")

    # ------------------------------------------------------------------
    # JSON 模板存取
    # ------------------------------------------------------------------
    def to_dict(self) -> dict:
        return {"version": 1, "name": self.name, "rows": self.rows,
                "cols": self.cols, "cells": [list(r) for r in self.cells]}

    @classmethod
    def from_dict(cls, data: dict) -> "ClassroomLayout":
        if not isinstance(data, dict) or "cells" not in data:
            raise LayoutError("这不是有效的教室模板文件")
        rows = int(data.get("rows") or len(data["cells"]))
        cols = int(data.get("cols") or len(data["cells"][0]))
        return cls(rows=rows, cols=cols, cells=data["cells"],
                   name=str(data.get("name") or "导入的教室"))

    def save(self, path: str) -> None:
        folder = os.path.dirname(os.path.abspath(path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls, path: str) -> "ClassroomLayout":
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except (OSError, json.JSONDecodeError) as exc:
            raise LayoutError(f"模板读取失败或格式不正确（{type(exc).__name__}）")
        return cls.from_dict(data)

    @staticmethod
    def template_dir(base_dir: str) -> str:
        """统一的模板目录：<项目目录>/教室模板"""
        return os.path.join(base_dir, "教室模板")

    @staticmethod
    def list_templates(base_dir: str) -> List[str]:
        """列出历史模板（返回完整路径，按文件名排序）。"""
        folder = ClassroomLayout.template_dir(base_dir)
        if not os.path.isdir(folder):
            return []
        names = [n for n in os.listdir(folder) if n.lower().endswith(".json")]
        return [os.path.join(folder, n) for n in sorted(names)]

    @staticmethod
    def safe_filename(name: str) -> str:
        """把老师输入的模板名转成安全文件名。"""
        cleaned = re.sub(r'[\\/:*?"<>|\r\n\t]', "_", (name or "").strip())
        return (cleaned or "我的教室")[:40]
