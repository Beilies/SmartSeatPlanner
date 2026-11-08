# -*- coding: utf-8 -*-
"""
图形界面 · 主窗口与座位画布（gui_main.py）
------------------------------------------------------------------
左侧：控制面板（导入 / 布局 / 参数 / 生成 / 切换 / 导出 / 撤销）
右侧：上方可缩放、可拖拽的座位画布；下方嵌入 matplotlib 评估图表

交互约定：
    左键拖动学生方格  = 交换座位（可拖到空位）
    Ctrl + Z / 撤销按钮 = 撤销上一步换座
    鼠标悬浮方格        = 弹出学号、总分、标签、禁止同桌对象
    Ctrl + 滚轮 / 工具栏 = 缩放画布
    鼠标滚轮（左侧面板）= 上下滚动控制面板（按钮、文字、滚动条上都有效）
所有按钮都套了 safe_call，任何异常都只给中文弹窗，不崩溃。
"""
from __future__ import annotations

import os
import traceback
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from evaluate import attach_stats
from excel_io import download_template, import_excel, list_sheets
from exporter import (export_canvas_png, export_report_xlsx,
                      export_seating_excel)
from gui_dialogs import (ImportPreviewDialog, LayoutEditorDialog,
                         EvaluationPanel, SheetChoiceDialog, log_exception,
                         safe_call, setup_style, show_error, show_info_list)
from layout import (BLOCK, PODIUM, SEAT, ClassroomLayout, Pos)
from models import COLOR, Student, WarningItem
from seating import (Plan, SeatingOptions, apply_swap, generate_plans)

from app_paths import BASE_DIR, asset_path
from ui_scale import apply_tk_scaling, minimum_size, px, window_size


# ======================================================================
# 座位画布（核心图形区）
# ======================================================================
class SeatBoard(ttk.Frame):
    """彩色方格座位图：悬浮提示、拖拽换座、撤销、缩放。"""

    # 设计稿是 96 DPI 下的逻辑像素；高 DPI 屏按系统缩放放大，
    # 否则字体被 Tk 放大后，格子还是原来那么大，姓名会挤成一团。
    BASE_W, BASE_H = px(118), px(66)

    def __init__(self, master, on_change=None, on_status=None):
        super().__init__(master)
        self.on_change = on_change
        self.on_status = on_status or (lambda msg: None)
        self.layout = ClassroomLayout.default()
        self.plan: Optional[Plan] = None
        self.students: List[Student] = []
        self.scale = 1.0
        self.undo_stacks: Dict[str, List[Tuple[Pos, Pos]]] = {}
        self.drag: Optional[dict] = None
        self.hover_pos: Optional[Pos] = None
        self.tip_job: Optional[str] = None
        self.pad = px(16)

        # ---- 工具栏 ----
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=6, pady=(6, 4))
        ttk.Label(bar, text="座位画布", style="Title.TLabel").pack(side=tk.LEFT)
        ttk.Button(bar, text="−", width=3, command=self.zoom_out).pack(side=tk.LEFT, padx=(12, 2))
        ttk.Button(bar, text="+", width=3, command=self.zoom_in).pack(side=tk.LEFT)
        ttk.Button(bar, text="适应窗口", command=self.fit).pack(side=tk.LEFT, padx=4)
        ttk.Button(bar, text="撤销换座 (Ctrl+Z)", command=self.undo).pack(side=tk.LEFT, padx=4)
        ttk.Label(bar, text="左键拖动学生换座｜悬浮看详情｜Ctrl+滚轮缩放",
                  style="Hint.TLabel").pack(side=tk.RIGHT)

        # ---- 画布 + 滚动条 ----
        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self.canvas = tk.Canvas(body, bg=COLOR["board"], highlightthickness=1,
                                highlightbackground="#CBD3DB")
        hsb = ttk.Scrollbar(body, orient=tk.HORIZONTAL, command=self.canvas.xview)
        vsb = ttk.Scrollbar(body, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        body.rowconfigure(0, weight=1)
        body.columnconfigure(0, weight=1)

        self.canvas.bind("<ButtonPress-1>", self._press)
        self.canvas.bind("<B1-Motion>", self._motion)
        self.canvas.bind("<ButtonRelease-1>", self._release)
        self.canvas.bind("<Motion>", self._hover)
        self.canvas.bind("<Leave>", lambda e: self._hide_tip())
        self.canvas.bind("<MouseWheel>", self._wheel)
        self.canvas.bind("<Button-3>", self._right_menu)

        # ---- 悬浮提示窗口 ----
        self.tip = tk.Toplevel(self)
        self.tip.wm_overrideredirect(True)
        self.tip.withdraw()
        self.tip_label = tk.Label(self.tip, justify=tk.LEFT, anchor="w",
                                  font=("微软雅黑", 11), bg="#FFFDF0", fg=COLOR["text"],
                                  relief=tk.SOLID, bd=1, padx=10, pady=7)
        self.tip_label.pack()
        self.menu = tk.Menu(self, tearoff=0)
        self.menu.add_command(label="撤销上一步换座", command=self.undo)
        self.menu.add_command(label="查看这个座位的完整信息", command=self._show_seat_info)
        self.right_pos: Optional[Pos] = None
        self.draw()

    # ------------------------------------------------------------------
    # 尺寸与坐标
    # ------------------------------------------------------------------
    def _cell_size(self) -> Tuple[float, float]:
        return self.BASE_W * self.scale, self.BASE_H * self.scale

    def _pos_at(self, x, y) -> Optional[Pos]:
        w, h = self._cell_size()
        c = int(self.canvas.canvasx(x) // w) - 1
        r = int(self.canvas.canvasy(y) // h) - 1
        if 0 <= r < self.layout.rows and 0 <= c < self.layout.cols:
            return (r, c)
        return None

    def set_plan(self, plan: Optional[Plan], layout: ClassroomLayout,
                 students: List[Student]) -> None:
        self.layout = layout
        self.students = students
        self.plan = plan
        self.draw()

    def draw(self) -> None:
        """重绘整张座位图。"""
        self.canvas.delete("all")
        w, h = self._cell_size()
        pad = self.pad
        cols, rows = self.layout.cols, self.layout.rows

        if self.plan is None:
            self.canvas.create_text(30, 30, anchor="nw",
                                    text="欢迎使用智能学生排座位系统\n\n"
                                         "① 左侧「导入学生 Excel」→ 核对预览与告警 → 确认导入\n"
                                         "② 需要时打开「教室布局编辑器」画出过道与讲台\n"
                                         "③ 调好参数，点「生成座位方案」，这里就会出现彩色座位图\n"
                                         "④ 直接拖动方格换座，下方图表会实时更新\n\n"
                                         "浅红=重点关注　白色=普通学生　灰色=过道/禁用",
                                    font=("微软雅黑", 14), fill=COLOR["subtext"],
                                    justify=tk.LEFT)
            self.canvas.configure(scrollregion=(0, 0, px(900), px(500)))
            return

        # 讲台横条
        self.canvas.create_rectangle(pad + w * 0.18, pad,
                                     pad + w * (cols - 0.18), pad + h * 0.74,
                                     fill=COLOR["podium"], outline=COLOR["podium"])
        self.canvas.create_text(pad + w * cols / 2, pad + h * 0.37,
                                text="讲　台　（前方）", fill="white",
                                font=("微软雅黑", max(10, int(12 * self.scale)), "bold"))

        name_font = ("微软雅黑", max(9, int(12 * self.scale)), "bold")
        tag_font = ("微软雅黑", max(8, int(10 * self.scale)))
        small_font = ("微软雅黑", max(8, int(10 * self.scale)))

        for r in range(rows):
            y0 = pad + (r + 1) * h
            self.canvas.create_text(pad + w * 0.5, y0 + h / 2, text=f"{r + 1}排",
                                    fill=COLOR["subtext"], font=small_font)
            for c in range(cols):
                x0 = pad + (c + 1) * w
                value = self.layout.get(r, c)
                stu = self.plan.assignment.get((r, c))
                if value == BLOCK:
                    fill, outline = COLOR["block"], COLOR["block_border"]
                elif value == PODIUM:
                    fill, outline = COLOR["podium"], COLOR["podium"]
                elif stu is None:
                    fill, outline = COLOR["empty"], COLOR["empty_border"]
                elif stu.is_focus():
                    fill, outline = COLOR["focus"], COLOR["focus_border"]
                else:
                    fill, outline = COLOR["normal"], COLOR["normal_border"]

                self.canvas.create_rectangle(x0 + 3, y0 + 3, x0 + w - 3, y0 + h - 3,
                                            fill=fill, outline=outline, width=1.4,
                                            tags=("seatrect", f"rect_{r}_{c}"))
                cx, cy = x0 + w / 2, y0 + h / 2
                if value == BLOCK:
                    self.canvas.create_text(cx, cy, text="过道", fill="#5A6B7B", font=small_font)
                elif value == PODIUM:
                    self.canvas.create_text(cx, cy, text="讲台", fill="white", font=small_font)
                elif stu is None:
                    self.canvas.create_text(cx, cy, text="空位", fill="#9AA6B2", font=small_font)
                else:
                    self.canvas.create_text(cx, cy - h * 0.12, text=stu.name,
                                            fill=COLOR["text"], font=name_font)
                    tag = stu.short_tag()
                    sub = tag if tag else (f"{stu.score:g}分" if stu.score is not None else "缺考")
                    self.canvas.create_text(cx, cy + h * 0.24, text=sub,
                                            fill="#B0554D" if tag else COLOR["subtext"],
                                            font=tag_font)

        self.canvas.configure(scrollregion=(
            0, 0, pad * 2 + (cols + 1.6) * w, pad * 2 + (rows + 1.6) * h))

    # ------------------------------------------------------------------
    # 悬浮提示
    # ------------------------------------------------------------------
    def _detail_text(self, pos: Pos) -> str:
        if self.layout.get(*pos) != SEAT:
            return "过道 / 禁用位置\n这里不安排学生"
        stu = self.plan.assignment.get(pos) if self.plan else None
        if stu is None:
            return (f"空座位｜{self.layout.zone(pos[0])} 第{pos[0] + 1}排 第{pos[1] + 1}列\n"
                    "可以把别的学生拖到这里")
        lines = [
            f"学号：{stu.sid or '未填'}",
            f"姓名：{stu.name}",
            f"总分：{stu.score_text()}",
            f"性别：{stu.gender or '未填'}",
            f"标签：{'重点关注' if stu.is_focus() else (stu.tag or '普通学生')}",
            f"禁止同桌：{stu.forbid_text() or '无'}",
            f"位置：{self.layout.zone(pos[0])} 第{pos[0] + 1}排 第{pos[1] + 1}列"
            + ("（靠过道）" if self.layout.is_near_aisle(pos) else ""),
        ]
        if stu.subjects:
            subj = "、".join(f"{k}:{'缺考' if v is None else f'{v:g}'}"
                            for k, v in list(stu.subjects.items())[:5])
            lines.append(f"各科：{subj}")
        return "\n".join(lines)

    def _hover(self, event) -> None:
        if self.drag and self.drag.get("active"):
            self._hide_tip()
            return
        pos = self._pos_at(event.x, event.y)
        if pos == self.hover_pos:
            if pos and self.tip.winfo_viewable():
                self.tip.geometry(f"+{event.x_root + 18}+{event.y_root + 14}")
            return
        self.hover_pos = pos
        if self.tip_job:
            try:
                self.after_cancel(self.tip_job)
            except Exception:
                pass
            self.tip_job = None
        if pos is None:
            self._hide_tip()
            return
        self.tip_job = self.after(280, lambda: self._show_tip(pos, event.x_root + 18, event.y_root + 14))

    def _show_tip(self, pos: Pos, rx: int, ry: int) -> None:
        if self.hover_pos != pos or self.plan is None:
            return
        self.tip_label.configure(text=self._detail_text(pos))
        self.tip.geometry(f"+{rx}+{ry}")
        self.tip.deiconify()
        self.tip.lift()

    def _hide_tip(self) -> None:
        self.hover_pos = None
        if self.tip_job:
            try:
                self.after_cancel(self.tip_job)
            except Exception:
                pass
            self.tip_job = None
        try:
            self.tip.withdraw()
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 拖拽换座 + 撤销
    # ------------------------------------------------------------------
    def _press(self, event) -> None:
        self._hide_tip()
        if self.plan is None:
            return
        pos = self._pos_at(event.x, event.y)
        if pos is None or self.layout.get(*pos) != SEAT:
            return
        self.drag = {"start": pos, "x": event.x, "y": event.y, "active": False}

    def _motion(self, event) -> None:
        if not self.drag:
            return
        if not self.drag["active"] and \
                abs(event.x - self.drag["x"]) + abs(event.y - self.drag["y"]) > 8:
            self.drag["active"] = True
            self._hide_tip()
        if not self.drag["active"]:
            return
        self.canvas.delete("ghost")
        w, h = self._cell_size()
        src = self.drag["start"]
        stu = self.plan.assignment.get(src)
        x, y = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
        self.canvas.create_rectangle(x - w * 0.38, y - h * 0.34,
                                     x + w * 0.38, y + h * 0.34,
                                     fill=COLOR["ghost"], outline="#C79A20",
                                     width=2, tags="ghost")
        self.canvas.create_text(x, y, text=stu.name if stu else "空位",
                                font=("微软雅黑", max(9, int(11 * self.scale)), "bold"),
                                fill=COLOR["text"], tags="ghost")
        # 高亮即将放下的目标格
        target = self._pos_at(event.x, event.y)
        self.canvas.itemconfigure("seatrect", width=1.4)
        if target and self.layout.get(*target) == SEAT and target != src:
            self.canvas.itemconfigure(f"rect_{target[0]}_{target[1]}",
                                      outline=COLOR["select"], width=3)

    def _release(self, event) -> None:
        if not self.drag:
            return
        drag, self.drag = self.drag, None
        self.canvas.delete("ghost")
        self.canvas.itemconfigure("seatrect", width=1.4)
        if not drag["active"] or self.plan is None:
            return
        src = drag["start"]
        target = self._pos_at(event.x, event.y)
        if target is None or target == src or self.layout.get(*target) != SEAT:
            self.on_status("没有换座：请把学生拖到另一个学生座位或空位上")
            return
        self._push_undo(src, target)
        apply_swap(self.plan, src, target)
        self.draw()
        a = self.plan.assignment.get(src)
        b = self.plan.assignment.get(target)
        self.on_status(f"已换座：{src[0] + 1}排{src[1] + 1}列 ↔ "
                       f"{target[0] + 1}排{target[1] + 1}列"
                       f"（{(a.name if a else '空位')} / {b.name if b else '空位'}）｜Ctrl+Z 可撤销")
        if self.on_change:
            self.on_change()

    def _plan_key(self) -> str:
        return self.plan.name if self.plan else ""

    def _push_undo(self, a: Pos, b: Pos) -> None:
        stack = self.undo_stacks.setdefault(self._plan_key(), [])
        stack.append((a, b))
        if len(stack) > 60:
            stack.pop(0)

    def undo(self) -> None:
        """撤销上一步拖拽换座（每套方案各自记录）。"""
        if self.plan is None:
            return
        stack = self.undo_stacks.get(self._plan_key(), [])
        if not stack:
            self.on_status("这一步没有可撤销的换座操作")
            return
        a, b = stack.pop()
        apply_swap(self.plan, a, b)
        self.draw()
        self.on_status("已撤销上一步换座")
        if self.on_change:
            self.on_change()

    def _right_menu(self, event) -> None:
        self.right_pos = self._pos_at(event.x, event.y)
        if self.right_pos and self.layout.get(*self.right_pos) == SEAT:
            try:
                self.menu.tk_popup(event.x_root, event.y_root)
            finally:
                self.menu.grab_release()

    def _show_seat_info(self) -> None:
        """右键菜单：把悬浮提示的内容用弹窗再显示一次，方便投影讲解。"""
        if self.right_pos is None or self.plan is None:
            return
        try:
            messagebox.showinfo("座位信息", self._detail_text(self.right_pos), parent=self)
        except Exception:
            pass

    # ------------------------------------------------------------------
    # 缩放
    # ------------------------------------------------------------------
    def _set_scale(self, value: float) -> None:
        self.scale = min(1.8, max(0.55, value))
        self.draw()

    def zoom_in(self) -> None:
        self._set_scale(self.scale + 0.12)

    def zoom_out(self) -> None:
        self._set_scale(self.scale - 0.12)

    def fit(self) -> None:
        self.update_idletasks()
        cw = max(300, self.canvas.winfo_width())
        ch = max(240, self.canvas.winfo_height())
        sx = (cw - self.pad * 3) / ((self.layout.cols + 1.6) * self.BASE_W)
        sy = (ch - self.pad * 3) / ((self.layout.rows + 1.7) * self.BASE_H)
        self._set_scale(min(sx, sy))

    def _wheel(self, event) -> None:
        if event.state & 0x0004:      # Ctrl + 滚轮 -> 缩放
            self._set_scale(self.scale + (0.08 if event.delta > 0 else -0.08))
            return "break"
        self.canvas.yview_scroll(-1 if event.delta > 0 else 1, "units")


# ======================================================================
# 主窗口
# ======================================================================
class MainWindow(tk.Tk):
    """程序主界面：左控制面板 + 右座位画布 / 评估图表。"""

    def __init__(self):
        super().__init__()
        # 必须在创建任何控件之前：Tk 会按真实 DPI 重新光栅化所有点数字体，
        # 文字不再是「96 DPI 位图被 Windows 拉伸」的模糊效果。
        apply_tk_scaling(self)
        self.title("智能学生排座位系统　｜　made by Beilies")
        w, h = window_size(self, 1380, 880)
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        x, y = max(0, (sw - w) // 2 - 20), max(0, (sh - h) // 2 - 40)
        self.geometry(f"{w}x{h}+{x}+{y}")
        self.minsize(*minimum_size(self, 1080, 680))
        setup_style(self)
        self._setup_icon()          # 老师设计的图标：标题栏 / 任务栏 / 面板左上角
        self.report_callback_exception = self._tk_exception

        # ---- 业务状态 ----
        self.students: List[Student] = []
        self.layout = ClassroomLayout.default()
        self.plans: List[Plan] = []
        self.notices: List[str] = []
        self.current_index = 0
        self.import_warnings: List[WarningItem] = []

        # ---- 参数变量 ----
        self.var_gender = tk.BooleanVar(value=True)
        self.var_blacklist = tk.BooleanVar(value=True)
        self.var_spread = tk.IntVar(value=50)
        self.var_plan_count = tk.IntVar(value=3)
        self.var_plan = tk.StringVar()
        self.var_status = tk.StringVar(value="就绪：请先导入学生成绩表")
        self.var_students = tk.StringVar(value="尚未导入学生")
        self.var_layout = tk.StringVar(value=self.layout.summary())
        self.var_desc = tk.StringVar(value="生成方案后这里会显示该方案的说明")

        self._build_left()
        self._build_right()
        self._bind_keys()
        self.board.set_plan(None, self.layout, self.students)
        self.eval_panel.show_empty()
        self.after(300, self.board.fit)

    def _setup_icon(self) -> None:
        """
        设置程序图标（老师自己设计的座位图）。

        三处都用同一份设计稿：
          1) exe 文件图标      —— 打包时由 --icon assets/app.ico 写入，资源管理器可见
          2) 标题栏 / 任务栏  —— iconbitmap 用 ico；iconphoto 用 png 做高 DPI 兜底
          3) 左侧面板左上角    —— 同一个 PhotoImage 直接放进标题行

        图标属于「锦上添花」：文件缺失或 Tk 不支持 PNG 时静默跳过，
        绝不让程序因为一个图片打不开，全部用 try 包住。
        """
        self._app_icon = None       # 必须长期持有引用，否则 Tk 会把图片回收掉
        try:
            ico = asset_path("app.ico")
            if os.path.isfile(ico):
                self.iconbitmap(ico)                 # Windows 窗口左上角小图标
        except Exception:
            pass
        try:
            png = asset_path("app_logo.png")
            if os.path.isfile(png):
                self._app_icon = tk.PhotoImage(file=png)
                # 第二个参数 default=True：预览窗、布局编辑器等子窗口也继承
                self.iconphoto(True, self._app_icon)
        except Exception:
            self._app_icon = None

    # ------------------------------------------------------------------
    # 界面搭建
    # ------------------------------------------------------------------
    def _build_left(self) -> None:
        """左侧控制面板（可滚动，小窗口也不会挤掉按钮）。"""
        outer = ttk.Frame(self, width=px(326))
        outer.grid(row=0, column=0, sticky="nsew", padx=(px(8), px(4)), pady=px(8))
        outer.pack_propagate(False)
        # 把「滚一格」固定成 20 逻辑像素（高 DPI 下同步放大）：Tk 默认一格 =
        # 窗口高度的 1/10，大屏上一滚就跳一大截，小屏上又几乎不动。
        canvas = tk.Canvas(outer, highlightthickness=0, width=px(310),
                           yscrollincrement=px(20))
        scroll = ttk.Scrollbar(outer, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=scroll.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        panel = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=panel, anchor="nw", width=px(306))
        panel.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        # 供滚轮绑定和 --diagnose 回归检查使用
        self.left_outer = outer
        self.left_canvas = canvas
        self.left_panel = panel

        # 标题行：左边放老师设计的图标，右边是控制面板标题
        head = ttk.Frame(panel)
        head.pack(fill=tk.X, pady=(2, 4))
        if self._app_icon is not None:
            ttk.Label(head, image=self._app_icon).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(head, text="智能排座位 · 控制面板",
                  style="Title.TLabel").pack(side=tk.LEFT)
        # 给老师一句明确提示：这块面板是可以往下滚的
        ttk.Label(panel, text="鼠标滚轮可上下滚动本面板",
                  style="Hint.TLabel", wraplength=px(270)).pack(anchor="w", pady=(0, 6))
        self.columnconfigure(0, weight=0, minsize=px(330))
        self.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        # ① 学生数据
        box = ttk.LabelFrame(panel, text="① 学生数据", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Button(box, text="导入学生 Excel（Ctrl+O）", style="Accent.TButton",
                   command=self._import_students).pack(fill=tk.X, pady=2)
        ttk.Button(box, text="下载标准 Excel 模板",
                   command=self._download_template).pack(fill=tk.X, pady=2)
        ttk.Label(box, textvariable=self.var_students, style="Hint.TLabel",
                  wraplength=px(270)).pack(anchor="w", pady=(4, 0))

        # ② 教室布局
        box = ttk.LabelFrame(panel, text="② 教室布局", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Button(box, text="打开可视化布局编辑器",
                   command=self._open_layout).pack(fill=tk.X, pady=2)
        ttk.Label(box, textvariable=self.var_layout, style="Hint.TLabel",
                  wraplength=px(270)).pack(anchor="w", pady=(4, 0))

        # ③ 排座参数
        box = ttk.LabelFrame(panel, text="③ 排座参数", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Checkbutton(box, text="性别均衡（尽量男女同桌）",
                        variable=self.var_gender).pack(anchor="w", pady=2)
        ttk.Checkbutton(box, text="启用同桌黑名单（禁止相邻）",
                        variable=self.var_blacklist).pack(anchor="w", pady=2)
        ttk.Label(box, text="优生打散权重（越大越不允许优生扎堆）",
                  style="Hint.TLabel").pack(anchor="w", pady=(6, 0))
        scale_row = ttk.Frame(box)
        scale_row.pack(fill=tk.X)
        ttk.Scale(scale_row, from_=0, to=100, variable=self.var_spread,
                  orient=tk.HORIZONTAL, command=self._on_spread).pack(
                      side=tk.LEFT, fill=tk.X, expand=True)
        self.lbl_spread = ttk.Label(scale_row, text="50", width=4)
        self.lbl_spread.pack(side=tk.RIGHT, padx=4)
        cnt_row = ttk.Frame(box)
        cnt_row.pack(fill=tk.X, pady=(6, 0))
        ttk.Label(cnt_row, text="生成方案数量（3~5 套）").pack(side=tk.LEFT)
        ttk.Spinbox(cnt_row, from_=3, to=5, width=5,
                    textvariable=self.var_plan_count).pack(side=tk.RIGHT)

        # ④ 生成与切换
        box = ttk.LabelFrame(panel, text="④ 生成与切换方案", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Button(box, text="生成座位方案（F5）", style="Accent.TButton",
                   command=self._generate).pack(fill=tk.X, pady=2)
        self.combo = ttk.Combobox(box, textvariable=self.var_plan,
                                  state="readonly", values=[])
        self.combo.pack(fill=tk.X, pady=4)
        self.combo.bind("<<ComboboxSelected>>", lambda e: self._on_combo())
        nav = ttk.Frame(box)
        nav.pack(fill=tk.X)
        ttk.Button(nav, text="← 上一套", command=self._prev_plan).pack(side=tk.LEFT, expand=True, fill=tk.X)
        ttk.Button(nav, text="下一套 →", command=self._next_plan).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=4)
        ttk.Label(box, textvariable=self.var_desc, style="Hint.TLabel",
                  wraplength=px(270)).pack(anchor="w", pady=(6, 0))

        # ⑤ 导出
        box = ttk.LabelFrame(panel, text="⑤ 导出", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Button(box, text="导出 Excel 座位表（Ctrl+S）",
                   command=self._export_excel).pack(fill=tk.X, pady=2)
        ttk.Button(box, text="导出座位图为 PNG（可打印）",
                   command=self._export_png).pack(fill=tk.X, pady=2)
        ttk.Button(box, text="导出方案统计报表",
                   command=self._export_report).pack(fill=tk.X, pady=2)

        # ⑥ 常用操作
        box = ttk.LabelFrame(panel, text="⑥ 常用操作", style="Card.TLabelframe")
        box.pack(fill=tk.X, pady=4)
        ttk.Button(box, text="撤销上一步换座（Ctrl+Z）",
                   command=lambda: self.board.undo()).pack(fill=tk.X, pady=2)
        ttk.Button(box, text="查看生成提示与降级建议",
                   command=self._show_notices).pack(fill=tk.X, pady=2)
        ttk.Button(box, text="使用说明 / 常见问题",
                   command=self._show_help).pack(fill=tk.X, pady=2)

        # 状态栏
        ttk.Label(panel, textvariable=self.var_status, style="Hint.TLabel",
                  wraplength=px(270), foreground="#24527A").pack(anchor="w", pady=(10, 4))

        # 所有子控件都建好之后，再统一绑定滚轮（详见方法内说明）
        self._bind_panel_wheel()

    def _bind_panel_wheel(self) -> None:
        """
        让鼠标滚轮在左侧面板的「任意位置」都能上下滚动。

        原来只写了 canvas.bind("<MouseWheel>", ...)，这就是滚轮失灵的根因：
        Tk 把滚轮事件交给鼠标正下方的那个控件，而面板里全是按钮 / 标签 /
        输入框，它们的绑定链条是 [控件 → 控件类 → 顶层窗口 → all]，
        canvas 并不在链条里。于是指针停在按钮上、或停在右侧滚动条上时，
        滚轮事件根本到不了 canvas，看起来就像「滚动条拖得动、滚轮没反应」。

        两层保险：
          ① 递归给面板里每个控件（含滚动条）绑控件级事件，先于类绑定触发，
             顺带挡住 Spinbox / Combobox 被滚轮误改数值；
          ② 在主窗口（顶层绑定级别）再兜一次底，且先判断鼠标位置确实在
             左侧面板范围内、上方没有弹窗，避免影响右侧座位画布和预览表格。
        Windows 用 <MouseWheel>，Linux / macOS 触控板另兼容 <Button-4/5>。
        """
        outer, canvas = self.left_outer, self.left_canvas
        WHEEL_UNITS = 3            # 20 像素 × 3 = 一格滚轮约 60 像素

        def do_scroll(event) -> Optional[str]:
            # 弹窗（导入预览 / 布局编辑器）里的滚轮交给它们自己处理。
            # 注意：MouseWheel 事件的 num 不是数字，而是字符串 "??"，
            # 直接 int() 会抛异常让 Tk 静默丢掉整个绑定，必须先判类型。
            try:
                if event.widget.winfo_toplevel() is not self:
                    return None
                raw_num = getattr(event, "num", 0)
                num = raw_num if isinstance(raw_num, int) else 0
                try:
                    delta = int(getattr(event, "delta", 0) or 0)
                except (TypeError, ValueError):
                    delta = 0
                if num == 4:                      # Linux 向上滚
                    direction, notches = -1, 1
                elif num == 5:                    # Linux 向下滚
                    direction, notches = 1, 1
                elif delta:
                    direction = -1 if delta > 0 else 1
                    # Windows 一格 = 120；触控板的细碎滚动按一格算
                    notches = min(max(abs(delta) // 120, 1), 5)
                else:
                    return None
                canvas.yview_scroll(direction * WHEEL_UNITS * notches, "units")
                return "break"                    # 不让滚轮改到方案数量等输入值
            except (tk.TclError, AttributeError, TypeError, ValueError):
                return None

        def global_scroll(event) -> Optional[str]:
            """顶层兜底：只接管鼠标确实停在左侧控制面板上的滚轮事件。"""
            try:
                if event.widget.winfo_toplevel() is not self:
                    return None
                under = self.winfo_containing(event.x_root, event.y_root)
                if under is None or under.winfo_toplevel() is not self:
                    return None
                x0, y0 = outer.winfo_rootx(), outer.winfo_rooty()
                if not (x0 <= event.x_root < x0 + outer.winfo_width()
                        and y0 <= event.y_root < y0 + outer.winfo_height()):
                    return None
            except (tk.TclError, AttributeError, TypeError, ValueError):
                return None
            return do_scroll(event)

        def bind_tree(widget) -> None:
            for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
                try:
                    widget.bind(seq, do_scroll, add="+")
                except tk.TclError:
                    pass                # 个别平台不支持某个事件序列时跳过
            for child in widget.winfo_children():
                bind_tree(child)

        bind_tree(outer)
        for seq in ("<MouseWheel>", "<Button-4>", "<Button-5>"):
            self.bind(seq, global_scroll, add="+")

    def _build_right(self) -> None:
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.grid(row=0, column=1, sticky="nsew", padx=(4, 8), pady=8)
        top = ttk.Frame(paned)
        bottom = ttk.Frame(paned)
        paned.add(top, weight=3)
        paned.add(bottom, weight=2)
        self.board = SeatBoard(top, on_change=self._on_board_change,
                               on_status=self.set_status)
        self.board.pack(fill=tk.BOTH, expand=True)
        self.eval_panel = EvaluationPanel(bottom)
        self.eval_panel.pack(fill=tk.BOTH, expand=True)

    def _bind_keys(self) -> None:
        self.bind_all("<Control-o>", lambda e: self._import_students())
        self.bind_all("<Control-O>", lambda e: self._import_students())
        self.bind_all("<Control-z>", lambda e: self.board.undo())
        self.bind_all("<Control-Z>", lambda e: self.board.undo())
        self.bind_all("<Control-s>", lambda e: self._export_excel())
        self.bind_all("<Control-S>", lambda e: self._export_excel())
        self.bind_all("<F5>", lambda e: self._generate())

    def _tk_exception(self, exc, value, tb) -> None:
        """Tk 回调里的漏网异常也统一转成中文提示 + 日志。"""
        err = value if isinstance(value, BaseException) else Exception(str(value))
        detail = "".join(traceback.format_exception(exc, value, tb))
        log_exception(err, where="Tk 回调", detail=detail)
        show_error(self, err, where="Tk 回调")

    def set_status(self, text: str) -> None:
        self.var_status.set(text)

    def _on_spread(self, _=None) -> None:
        try:
            self.lbl_spread.configure(text=str(int(float(self.var_spread.get()))))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # ① 导入
    # ------------------------------------------------------------------
    @safe_call
    def _import_students(self) -> None:
        path = filedialog.askopenfilename(
            parent=self, title="选择学生成绩 Excel",
            filetypes=[("Excel 文件", "*.xlsx *.xls"), ("所有文件", "*.*")])
        if not path:
            return
        sheets = list_sheets(path)
        sheet = sheets[0]
        if len(sheets) > 1:
            dlg = SheetChoiceDialog(self, sheets)
            if not dlg.result:
                return
            sheet = dlg.result
        result = import_excel(path, sheet)
        ImportPreviewDialog(self, result, self._on_import_confirm)
        self.set_status(f"已选择文件：{os.path.basename(path)}（{sheet}），请在预览窗口核对")

    def _on_import_confirm(self, students: List[Student],
                           warnings: List[WarningItem]) -> None:
        self.students = students
        self.import_warnings = warnings
        self.plans = []
        self.notices = []
        self.current_index = 0
        self.combo.configure(values=[])
        self.var_plan.set("")
        self.var_students.set(f"已导入 {len(students)} 名学生｜"
                              f"重点关注 {sum(1 for s in students if s.is_focus())} 人｜"
                              f"无总分 {sum(1 for s in students if s.score is None)} 人")
        self.board.set_plan(None, self.layout, self.students)
        self.eval_panel.show_empty("学生名单已确认。点左侧「生成座位方案（F5）」，"
                                   "这里会显示柱状图和分布统计。")
        self.set_status(f"导入完成，共 {len(students)} 人，请点「生成座位方案」")
        if warnings:
            show_info_list(self, "导入提醒",
                           [w.display() for w in warnings[:30]],
                           subtitle="以下记录供老师核对（预览表里已经可以修改）：")

    @safe_call
    def _download_template(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self, title="保存标准 Excel 模板", defaultextension=".xlsx",
            initialfile="学生成绩标准模板.xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        download_template(path)
        messagebox.showinfo("模板已保存", f"标准模板已保存到：\n{path}\n\n"
                            "里面有示例数据和「填写说明」Sheet，照着填即可。", parent=self)

    # ------------------------------------------------------------------
    # ② 布局
    # ------------------------------------------------------------------
    @safe_call
    def _open_layout(self) -> None:
        LayoutEditorDialog(self, self.layout, self._on_layout_confirm)

    def _on_layout_confirm(self, layout: ClassroomLayout) -> None:
        self.layout = layout
        self.var_layout.set(layout.summary())
        had_plans = bool(self.plans)
        self.plans = []
        self.combo.configure(values=[])
        self.var_plan.set("")
        self.board.set_plan(None, self.layout, self.students)
        self.eval_panel.show_empty("教室布局已更新，请重新点「生成座位方案」。")
        self.set_status(f"布局已更新：{layout.summary()}")
        if had_plans:
            messagebox.showinfo("布局已更新",
                                "教室行列或过道变了，原有座位方案已清空。\n"
                                "请重新点「生成座位方案」（只需一秒）。", parent=self)

    # ------------------------------------------------------------------
    # ④ 生成 / 切换
    # ------------------------------------------------------------------
    @safe_call
    def _generate(self) -> None:
        if not self.students:
            messagebox.showwarning("还没有学生", "请先点「导入学生 Excel」并确认导入名单", parent=self)
            return
        if not self.layout.seats():
            messagebox.showwarning("没有可用座位", "当前教室布局里没有学生座位，请先打开布局编辑器标记座位",
                                   parent=self)
            return
        options = SeatingOptions(
            gender_balance=bool(self.var_gender.get()),
            use_blacklist=bool(self.var_blacklist.get()),
            spread_weight=max(0.0, min(1.0, float(self.var_spread.get()) / 100.0)),
            plan_count=int(self.var_plan_count.get()),
        )
        self.set_status("正在计算方案，请稍候…")
        self.update_idletasks()
        plans, notices = generate_plans(self.students, self.layout, options)
        attach_stats(plans, self.layout, self.students)
        self.plans, self.notices, self.current_index = plans, notices, 0
        self.combo.configure(values=[p.name for p in plans])
        self._show_current_plan()
        self.set_status(f"已生成 {len(plans)} 套差异化方案，可在下拉框一键切换；拖动方格即可换座")
        if notices:
            show_info_list(self, "生成提示与降级方案", notices,
                           subtitle="系统已经自动降级处理，方案可以正常使用，建议看一眼：")
        else:
            messagebox.showinfo("生成完成", f"已生成 {len(plans)} 套方案。\n"
                                "用左侧下拉框切换，右下角图表会实时对比。", parent=self)

    def _show_current_plan(self) -> None:
        if not self.plans:
            return
        self.current_index = max(0, min(self.current_index, len(self.plans) - 1))
        plan = self.plans[self.current_index]
        self.combo.current(self.current_index)
        self.var_plan.set(plan.name)
        self.var_desc.set(plan.description)
        self.board.set_plan(plan, self.layout, self.students)
        self.eval_panel.show(plan, self.layout, self.students)

    @safe_call
    def _on_combo(self) -> None:
        idx = self.combo.current()
        if idx >= 0 and self.plans:
            self.current_index = idx
            self._show_current_plan()
            self.set_status(f"已切换到 {self.plans[idx].name}")

    @safe_call
    def _prev_plan(self) -> None:
        if not self.plans:
            messagebox.showinfo("提示", "请先点「生成座位方案」", parent=self)
            return
        self.current_index = (self.current_index - 1) % len(self.plans)
        self._show_current_plan()

    @safe_call
    def _next_plan(self) -> None:
        if not self.plans:
            messagebox.showinfo("提示", "请先点「生成座位方案」", parent=self)
            return
        self.current_index = (self.current_index + 1) % len(self.plans)
        self._show_current_plan()

    def _on_board_change(self) -> None:
        """拖拽 / 撤销后重算统计，保证图表永远和画布一致。"""
        if not self.plans:
            return
        attach_stats(self.plans, self.layout, self.students)
        self.eval_panel.show(self.plans[self.current_index], self.layout, self.students)

    # ------------------------------------------------------------------
    # ⑤ 导出
    # ------------------------------------------------------------------
    def _current_plan(self) -> Optional[Plan]:
        if not self.plans:
            messagebox.showwarning("还没有方案", "请先点「生成座位方案」，然后再导出", parent=self)
            return None
        return self.plans[self.current_index]

    @safe_call
    def _export_excel(self) -> None:
        plan = self._current_plan()
        if plan is None:
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="导出 Excel 座位表", defaultextension=".xlsx",
            initialfile=f"{plan.name}_座位表.xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        saved = export_seating_excel(path, plan, self.layout, self.students)
        messagebox.showinfo("导出成功", f"已导出：\n{saved}\n\n"
                            "含「学生名单」和「座位示意图」两个工作表，可直接打印。", parent=self)
        self.set_status(f"已导出 Excel 座位表：{os.path.basename(saved)}")

    @safe_call
    def _export_png(self) -> None:
        plan = self._current_plan()
        if plan is None:
            return
        path = filedialog.asksaveasfilename(
            parent=self, title="导出座位图为 PNG", defaultextension=".png",
            initialfile=f"{plan.name}_座位图.png", filetypes=[("PNG 图片", "*.png")])
        if not path:
            return
        saved = export_canvas_png(path, plan, self.layout)
        messagebox.showinfo("导出成功", f"座位图已保存为图片：\n{saved}\n\n"
                            "图片只含讲台、排号和姓名方格：过道位置已留出空白，"
                            "不画过道方框，也不含重点关注标记、标题、图例、导出时间，"
                            "可直接发家长群或打印张贴。",
                            parent=self)
        self.set_status(f"已导出 PNG 座位图：{os.path.basename(saved)}")

    @safe_call
    def _export_report(self) -> None:
        plan = self._current_plan()
        if plan is None:
            return
        attach_stats(self.plans, self.layout, self.students)
        path = filedialog.asksaveasfilename(
            parent=self, title="导出方案统计报表", defaultextension=".xlsx",
            initialfile="座位方案统计报表.xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        saved = export_report_xlsx(path, self.plans, self.layout,
                                   self.students, self.notices)
        messagebox.showinfo("导出成功", f"统计报表已保存：\n{saved}\n\n"
                            "含「方案对比」「分区分布」「提示与冲突」三个工作表。", parent=self)
        self.set_status(f"已导出统计报表：{os.path.basename(saved)}")

    # ------------------------------------------------------------------
    # ⑥ 帮助
    # ------------------------------------------------------------------
    @safe_call
    def _show_notices(self) -> None:
        lines = list(self.notices)
        for p in self.plans:
            lines += [f"{p.name}：{c}" for c in p.conflicts]
        if not lines:
            lines = ["本次生成没有冲突或降级提示，所有约束都满足了。"]
        show_info_list(self, "生成提示与降级建议", lines,
                       subtitle="这里记录座位不够、黑名单无解等情况以及建议做法：")

    @safe_call
    def _show_help(self) -> None:
        tips = [
            "滚动：左侧控制面板内容较长，鼠标放在面板任意位置（按钮、文字、右侧滚动条上都行）"
            "滚动滚轮即可上下翻，也可以直接拖动右侧滚动条；右侧座位画布的滚轮只移动座位图，"
            "Ctrl+滚轮才是缩放。",
            "导入：点「导入学生 Excel」→ 如果多个工作表会让你选 → 弹出预览窗口，"
            "先看告警（空姓名、异常分数、重复学生、缺考名单），双击单元格就能改，删掉多余行后点「确认导入」。",
            "模板：点「下载标准 Excel 模板」，字段为 学号、姓名、总分、性别、特殊标记、"
            "禁止同桌名单、各科成绩。表头不在第一行、合并单元格、空行、底部合计行都能自动处理；"
            "缺考 / — / N/A 都算缺失，绝不按 0 分。",
            "布局：点「打开可视化布局编辑器」，先设行列生成空白网格；左键点 = 学生座位，"
            "右键点 = 过道/禁用，按住拖动 = 框选批量修改；可保存 / 加载 JSON 教室模板。",
            "排座：勾选性别均衡、黑名单，拖「优生打散权重」滑块，选生成 3~5 套，点「生成座位方案」。"
            "五套策略分别是 强弱帮扶 / 同水平就近 / 均衡随机 / 性别优先 / 特殊学生优先。",
            "看方案：左侧下拉框一键切换，右侧座位图实时刷新；下方柱状图对比前中后排平均分，"
            "文字区给出优生、学困生、重点关注学生分布和一句话结论。",
            "调座位：直接用鼠标左键把学生方格拖到另一个方格（或空位）即可交换；Ctrl+Z 或"
            "「撤销换座」按钮可悔棋；鼠标悬浮可看学号、总分、标签、禁止同桌对象；Ctrl+滚轮缩放。",
            "导出：Excel 座位表（名单 + 彩色座位示意图）、PNG 座位图、方案统计报表。"
            "PNG 是给家长看的：只画讲台、排号和姓名方格，过道按原位置留白"
            "（不画方框、不写「过道」），不显示重点关注标记（浅红底色也不显示），"
            "空位统一留在最后几排；"
            "老师自己在主界面画布上仍能看到浅红标记和灰色过道。",
            "依赖：Python 官方安装包自带 tkinter；另需 pandas、openpyxl、xlrd、matplotlib，"
            "安装命令见根目录 requirements.txt；完整教程在 docs/使用说明.md。",
            "作者：北离Beilies",
        ]
        show_info_list(self, "使用说明 / 常见问题", tips,
                       subtitle="面向老师的六步流程（不需要联网、不需要登录）：")
