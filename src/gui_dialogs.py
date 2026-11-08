# -*- coding: utf-8 -*-
"""
图形界面 · 弹窗模块（gui_dialogs.py）
------------------------------------------------------------------
包含：
    setup_style / show_error / safe_call / show_info_list   通用界面工具
    SheetChoiceDialog       选择 Excel 工作表
    EditableTable           可双击编辑的表格控件（导入预览用）
    ImportPreviewDialog     导入预览 + 告警 + 直接改数据 + 下载模板
    LayoutEditorDialog      可视化教室布局编辑器（左键/右键/框选/模板）
    EvaluationPanel         嵌入主窗口的 matplotlib 柱状图 + 统计文本

所有错误统一转成中文弹窗，完整异常只写入「运行日志.txt」，
老师不会看到任何原始代码报错。
"""
from __future__ import annotations

import functools
import os
import traceback
from datetime import datetime
from typing import Callable, List, Optional

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure

from evaluate import evaluate_plan
from excel_io import (PREVIEW_COLUMNS, ImportResult, download_template,
                      students_from_table, students_to_table)
from exporter import setup_chinese_font
from layout import (BLOCK, CELL_LABEL, PODIUM, SEAT, ClassroomLayout,
                    LayoutError)
from models import COLOR, WarningItem, Student
from seating import Plan

from app_paths import BASE_DIR, LOG_FILENAME
from ui_scale import px
LOG_FILE = os.path.join(BASE_DIR, LOG_FILENAME)


# ----------------------------------------------------------------------
# 通用工具
# ----------------------------------------------------------------------
def setup_style(root: tk.Misc) -> None:
    """统一窗口外观（浅色主题、微软雅黑、稍大的控件间距与字号）。

    字号用「点」：Tk 已按真实 DPI 设置 scaling，11pt 在 150% 屏上会自动
    变成约 22 物理像素，笔画重新光栅化，不会是拉伸后的模糊字。
    行高 / 内边距用「像素」，必须用 px() 同步放大，否则字大框小。
    基础字号整体提了 1pt（10→11、9→10），100% DPI 的屏幕也更清楚。
    """
    try:
        style = ttk.Style(root)
        if "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure(".", font=("微软雅黑", 11))
        style.configure("TButton", padding=(px(9), px(6)), font=("微软雅黑", 11))
        style.configure("Title.TLabel", font=("微软雅黑", 13, "bold"), foreground="#1F2933")
        style.configure("Hint.TLabel", font=("微软雅黑", 11), foreground=COLOR["subtext"])
        style.configure("Card.TLabelframe", padding=px(8))
        style.configure("Card.TLabelframe.Label", font=("微软雅黑", 11, "bold"),
                        foreground="#2C5282")
        style.configure("Treeview", rowheight=px(28), font=("微软雅黑", 11))
        style.configure("Treeview.Heading", font=("微软雅黑", 11, "bold"))
        style.configure("Accent.TButton", foreground="#FFFFFF", background="#4472C4")
        style.map("Accent.TButton", background=[("active", "#3559A0")])
    except Exception:
        pass  # 样式设置失败不影响功能


def fit_dialog(win, pref_w: int, pref_h: int,
               min_w: int = 520, min_h: int = 360) -> None:
    """
    让弹窗适应老师电脑的屏幕并居中，避免按钮跑到屏幕外。

    为什么要单独处理：老式 1366×768 笔记本、或系统显示缩放 125%/150% 时，
    写死 1080×720 的窗口会高过屏幕可用区域，窗口最下面的「确认导入」按钮
    就看不到了，老师会卡在预览这一步。这里按屏幕尺寸留出任务栏/标题栏的
    余量再决定窗口大小，同时把 minsize 也限制在屏幕内。

    注意：Tk 的 pack 按「添加顺序」分配空间，最后才添加的底部按钮在空间
    不足时会被裁掉。所以调用方必须先 pack(side=tk.BOTTOM) 按钮区，
    再 pack 可伸缩的表格/画布，让空间不够时优先压缩内容区而不是按钮。
    """
    # 设计稿按 96 DPI 写死，高 DPI 屏上要同步放大；否则按钮文字变大了、
    # 窗口却还是原来那么小，底部确认按钮又会被挤出去。
    pref_w, pref_h = px(pref_w), px(pref_h)
    min_w, min_h = px(min_w), px(min_h)
    try:
        win.update_idletasks()
        sw, sh = win.winfo_screenwidth(), win.winfo_screenheight()
        avail_w = max(px(560), sw - px(80))       # 左右留边
        avail_h = max(px(340), sh - px(100))      # 上下留任务栏 / 标题栏 / 边框
        w = min(pref_w, avail_w)
        h = min(pref_h, avail_h)
        win.minsize(min(min_w, avail_w), min(min_h, avail_h))
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2 - 24)
        win.geometry(f"{w}x{h}+{x}+{y}")
    except Exception:
        try:
            win.geometry(f"{pref_w}x{pref_h}")
            win.minsize(min_w, min_h)
        except Exception:
            pass


def log_exception(exc: Exception, where: str = "", detail: str = "") -> None:
    """把完整堆栈写入日志文件；界面只显示友好中文。"""
    try:
        text = detail or traceback.format_exc()
        if not text or text.startswith("NoneType: None"):
            text = f"{type(exc).__name__}: {exc}\n"
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} {where} =====\n")
            f.write(text)
    except Exception:
        pass


def show_error(parent, exc: Exception, where: str = "") -> None:
    """统一的友好报错弹窗：自定义业务异常显示原因，其它异常只给通用建议。"""
    log_exception(exc, where)
    known = ("ExcelParseError", "SeatingError", "LayoutError", "ExportError")
    if type(exc).__name__ in known and str(exc).strip():
        text = str(exc)
    else:
        text = ("刚才这一步没能完成，常见原因：文件正被 Excel 打开、文件损坏、"
                "数据为空或模板格式不对。\n\n可以换一个文件重试；详细记录已写入「运行日志.txt」，"
                "发给技术人员即可定位问题。")
    try:
        messagebox.showerror("操作未完成", text, parent=parent)
    except Exception:
        pass


def safe_call(func):
    """给所有按钮回调套一层保护：任何异常都变成中文弹窗，绝不崩溃退出。"""
    @functools.wraps(func)
    def wrapper(self, *args, **kwargs):
        try:
            return func(self, *args, **kwargs)
        except Exception as exc:
            try:
                parent = self.winfo_toplevel()
            except Exception:
                parent = None
            show_error(parent, exc, where=getattr(func, "__qualname__", "界面操作"))
            return None
    return wrapper


def show_info_list(parent, title: str, lines: List[str], subtitle: str = "") -> None:
    """通用「提示 / 降级方案」窗口，纯图形展示，老师点关闭即可。"""
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    fit_dialog(win, 720, 420, min_w=480, min_h=260)
    # 关闭按钮先占住底部，窗口再小也不会被挤掉
    ttk.Button(win, text="知道了", command=win.destroy).pack(
        side=tk.BOTTOM, pady=(4, 12))
    ttk.Label(win, text=subtitle or title, style="Title.TLabel").pack(
        side=tk.TOP, padx=14, pady=(12, 6), anchor="w")
    frame = ttk.Frame(win)
    frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=14, pady=6)
    text = tk.Text(frame, wrap=tk.WORD, font=("微软雅黑", 11), bg="#FBFCFD",
                   relief=tk.FLAT, padx=8, pady=8)
    sb = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text.yview)
    text.configure(yscrollcommand=sb.set)
    text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    sb.pack(side=tk.RIGHT, fill=tk.Y)
    for i, line in enumerate(lines or ["（没有额外提示）"], start=1):
        text.insert(tk.END, f"{i}. {line}\n\n")
    text.configure(state=tk.DISABLED)
    win.grab_set()
    win.focus_set()


# ----------------------------------------------------------------------
# 选择 Sheet
# ----------------------------------------------------------------------
class SheetChoiceDialog(tk.Toplevel):
    """导入有多个 Sheet 时，让老师选一个。"""

    def __init__(self, master, sheets: List[str]):
        super().__init__(master)
        self.title("选择工作表")
        self.transient(master)
        self.resizable(False, False)
        fit_dialog(self, 720, 420, min_w=480, min_h=260)
        self.result: Optional[str] = None

        # 按钮区先占底部，再放内容，保证「确定」永远可见
        bar = ttk.Frame(self)
        bar.pack(side=tk.BOTTOM, fill=tk.X, padx=16, pady=(6, 14))
        ttk.Button(bar, text="确定", style="Accent.TButton",
                   command=self._ok).pack(side=tk.RIGHT, padx=6)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side=tk.RIGHT)

        ttk.Label(self, text="这个 Excel 里有多个工作表，请选择要导入的那一个：",
                  style="Title.TLabel").pack(side=tk.TOP, padx=16, pady=(14, 8), anchor="w")
        self.listbox = tk.Listbox(self, height=min(10, max(4, len(sheets))),
                                  font=("微软雅黑", 12), activestyle="dotbox")
        for s in sheets:
            self.listbox.insert(tk.END, s)
        self.listbox.selection_set(0)
        self.listbox.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=16, pady=4)
        self.listbox.bind("<Double-Button-1>", lambda e: self._ok())
        self.bind("<Return>", lambda e: self._ok())
        self.grab_set()
        self.wait_window(self)

    def _ok(self) -> None:
        sel = self.listbox.curselection()
        if sel:
            self.result = self.listbox.get(sel[0])
        self.destroy()


# ----------------------------------------------------------------------
# 可双击编辑的表格控件
# ----------------------------------------------------------------------
class EditableTable(ttk.Frame):
    """带滚动条、可双击单元格修改、可删除/新增行的 Treeview 封装。"""

    def __init__(self, master, columns: List[str], rows: List[List[str]],
                 widths: Optional[List[int]] = None, height: int = 14):
        super().__init__(master)
        self.columns = list(columns)
        self.tree = ttk.Treeview(self, columns=self.columns, show="headings",
                                 height=height, selectmode="extended")
        default_widths = [110, 100, 90, 70, 130, 180, 230]
        for i, col in enumerate(self.columns):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=(widths or default_widths)[i] if i < len(widths or default_widths) else 120,
                             anchor=tk.CENTER, stretch=True)
        vsb = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        self.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.tree.tag_configure("odd", background="#F6F8FB")
        self.tree.tag_configure("bad", background="#FDE9E7")
        self.tree.bind("<Double-Button-1>", self._begin_edit)
        self.editor: Optional[ttk.Entry] = None
        self.set_rows(rows)

    def set_rows(self, rows: List[List[str]]) -> None:
        self.tree.delete(*self.tree.get_children())
        for i, row in enumerate(rows):
            vals = [("" if v is None else str(v)) for v in row]
            vals += [""] * (len(self.columns) - len(vals))
            tags = ["odd"] if i % 2 else []
            if not vals[1].strip():
                tags.append("bad")
            self.tree.insert("", tk.END, values=vals[:len(self.columns)], tags=tags)

    def get_rows(self) -> List[List[str]]:
        return [list(self.tree.item(i, "values")) for i in self.tree.get_children()]

    def _begin_edit(self, event) -> None:
        row_id = self.tree.identify_row(event.y)
        col_id = self.tree.identify_column(event.x)
        if not row_id or not col_id or self.editor is not None:
            return
        idx = int(col_id.lstrip("#")) - 1
        if not (0 <= idx < len(self.columns)):
            return
        x, y, w, h = self.tree.bbox(row_id, col_id)
        column_key = self.columns[idx]
        value = self.tree.set(row_id, column_key)
        editor = ttk.Entry(self.tree, font=("微软雅黑", 11))
        editor.place(x=x, y=y, width=w, height=h)
        editor.insert(0, value)
        editor.select_range(0, tk.END)
        editor.focus_set()
        self.editor = editor

        def commit(_=None):
            if self.editor is not editor:
                return
            self.tree.set(row_id, column_key, editor.get())
            # 改完姓名后刷新行底色
            if column_key == "姓名":
                tags = [t for t in self.tree.item(row_id, "tags") if t != "bad"]
                if not editor.get().strip():
                    tags.append("bad")
                self.tree.item(row_id, tags=tags)
            editor.destroy()
            self.editor = None

        editor.bind("<Return>", commit)
        editor.bind("<FocusOut>", commit)
        editor.bind("<Escape>", lambda e: (editor.destroy(), setattr(self, "editor", None)))

    def delete_selected(self) -> int:
        ids = self.tree.selection()
        for i in ids:
            self.tree.delete(i)
        return len(ids)

    def add_row(self) -> None:
        self.tree.insert("", tk.END, values=[""] * len(self.columns), tags=["bad"])
        children = self.tree.get_children()
        if children:
            self.tree.see(children[-1])

    def count(self) -> int:
        return len(self.tree.get_children())


# ----------------------------------------------------------------------
# 导入预览窗口
# ----------------------------------------------------------------------
class ImportPreviewDialog(tk.Toplevel):
    """导入后的预览：告警列表 + 可直接改的学生表格 + 下载模板。"""

    def __init__(self, master, result: ImportResult,
                 on_confirm: Callable[[List[Student], List[WarningItem]], None]):
        super().__init__(master)
        self.title("导入预览 · 请核对后再确认")
        self.transient(master)
        # 按屏幕自适应居中：写死 1080×720 时，125%/150% 缩放的笔记本会把
        # 窗口底部顶出屏幕，老师就找不到「确认导入」按钮了。
        fit_dialog(self, 1080, 720, min_w=760, min_h=460)
        self.result = result
        self.on_confirm = on_confirm

        # ① 底部操作条必须先 pack。Tk 的 pack 按添加顺序分配空间，
        #    空间不足时最后添加的元素会被直接裁掉；先占住底部，按钮永远可见。
        bar = ttk.Frame(self)
        bar.pack(side=tk.BOTTOM, fill=tk.X, padx=14, pady=(4, 10))
        # 主按钮必须最先 pack：窄窗口下 pack 会裁掉最后添加的元素，
        # 先让「确认导入」占住右侧，其余次要按钮放不下时优先被压缩。
        self.confirm_button_bottom = ttk.Button(
            bar, text="确认导入 →", style="Accent.TButton", width=16,
            command=self._confirm)
        self.confirm_button_bottom.pack(side=tk.RIGHT, padx=6)
        ttk.Button(bar, text="取消", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Button(bar, text="下载标准 Excel 模板",
                   command=self._download_template).pack(side=tk.LEFT)
        ttk.Button(bar, text="新增一行", command=self._add_row).pack(side=tk.LEFT, padx=6)
        ttk.Button(bar, text="删除选中行", command=self._delete_rows).pack(side=tk.LEFT)

        # ② 顶部：标题 + 右侧醒目的主按钮（双保险，一抬眼就能看到下一步）
        head = ttk.Frame(self)
        head.pack(side=tk.TOP, fill=tk.X, padx=14, pady=(10, 4))
        # 右上角主按钮先占位，左侧说明文字用 expand 自适应剩余宽度
        right = ttk.Frame(head)
        right.pack(side=tk.RIGHT, padx=(10, 0))
        self.confirm_button_top = ttk.Button(
            right, text="下一步：确认导入 →", style="Accent.TButton",
            width=20, command=self._confirm)
        self.confirm_button_top.pack(side=tk.TOP)
        ttk.Label(right, text="核对无误就点上面", style="Hint.TLabel").pack(
            side=tk.TOP, pady=(2, 0))
        left = ttk.Frame(head)
        left.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Label(left, text="第 1 步：核对导入结果", style="Title.TLabel").pack(anchor="w")
        info = (f"文件：{os.path.basename(result.path)}　工作表：{result.sheet_name}　"
                f"表头在第 {result.header_row} 行　识别到 {len(result.students)} 名学生　"
                f"忽略合计/空行 {result.ignored_rows} 行")
        ttk.Label(left, text=info, style="Hint.TLabel").pack(anchor="w", pady=(2, 0))

        # ③ 告警区（固定高度，不参与伸缩）
        warn_box = ttk.LabelFrame(self, text="导入告警（红色需要留意；蓝色只是说明）",
                                  style="Card.TLabelframe")
        warn_box.pack(side=tk.TOP, fill=tk.X, padx=14, pady=6)
        self.warn_tree = ttk.Treeview(warn_box, columns=("kind", "detail"),
                                      show="headings", height=5)
        self.warn_tree.heading("kind", text="类别")
        self.warn_tree.heading("detail", text="说明")
        self.warn_tree.column("kind", width=px(110), anchor=tk.CENTER, stretch=False)
        self.warn_tree.column("detail", width=px(700), anchor=tk.W)
        warn_scroll = ttk.Scrollbar(warn_box, orient=tk.VERTICAL, command=self.warn_tree.yview)
        self.warn_tree.configure(yscrollcommand=warn_scroll.set)
        self.warn_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0), pady=4)
        warn_scroll.pack(side=tk.RIGHT, fill=tk.Y, pady=4)
        self.warn_tree.tag_configure("danger", background="#FDE9E7", foreground="#9B2C2C")
        self.warn_tree.tag_configure("warn", background="#FFF3DC", foreground="#8A5A00")
        self.warn_tree.tag_configure("info", background="#EAF2FC", foreground="#24527A")
        self._fill_warnings(result.warnings)

        # ④ 表格区最后 pack，空间不够时优先压缩它，而不是挤掉按钮
        table_box = ttk.LabelFrame(
            self,
            text="学生名单（双击任意单元格可修改；可删除多余行）　"
                 "※ 核对完点右上角或窗口最下方的「确认导入 →」",
            style="Card.TLabelframe")
        table_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=14, pady=6)
        self.table = EditableTable(table_box, PREVIEW_COLUMNS,
                                   students_to_table(result.students), height=10)
        self.table.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # Ctrl+回车也能确认；不占用普通回车，避免和单元格编辑冲突
        self.bind("<Control-Return>", lambda e: self._confirm())
        self.grab_set()
        self.focus_set()

    def _fill_warnings(self, warnings: List[WarningItem]) -> None:
        for w in warnings:
            tag = "danger" if w.kind in ("empty_name", "bad_score", "duplicate", "conflict") \
                else ("warn" if w.kind == "absent" else "info")
            from models import WARN_KIND
            self.warn_tree.insert("", tk.END,
                                  values=(WARN_KIND.get(w.kind, "提示"), w.message),
                                  tags=[tag])
        if not warnings:
            self.warn_tree.insert("", tk.END, values=("说明", "导入顺利，没有发现问题"), tags=["info"])

    @safe_call
    def _download_template(self) -> None:
        path = filedialog.asksaveasfilename(
            parent=self, title="保存标准 Excel 模板", defaultextension=".xlsx",
            initialfile="学生成绩标准模板.xlsx",
            filetypes=[("Excel 文件", "*.xlsx")])
        if not path:
            return
        download_template(path)
        messagebox.showinfo("模板已保存", f"模板已保存到：\n{path}\n\n照着「填写说明」Sheet 填写即可。",
                            parent=self)

    @safe_call
    def _add_row(self) -> None:
        self.table.add_row()

    @safe_call
    def _delete_rows(self) -> None:
        n = self.table.delete_selected()
        if n == 0:
            messagebox.showinfo("提示", "请先在表格里点选要删除的行（可按住 Ctrl 多选）", parent=self)

    @safe_call
    def _confirm(self) -> None:
        # 若正在编辑某个单元格，先把内容落表
        if self.table.editor is not None:
            self.table.editor.event_generate("<FocusOut>")
        students, warnings = students_from_table(self.table.get_rows())
        if not students:
            messagebox.showwarning("名单是空的", "预览表里没有有效学生，请重新导入或新增行", parent=self)
            return
        empty_names = [s for s in students if not s.name.strip()]
        if empty_names:
            ok = messagebox.askyesno(
                "还有空姓名",
                f"预览表里还有 {len(empty_names)} 行没有姓名。\n\n"
                "点「是」：忽略这些行继续排座；\n点「否」：回到预览表补写姓名。",
                parent=self)
            if not ok:
                return
            students = [s for s in students if s.name.strip()]
        self.on_confirm(students, warnings)
        self.destroy()


# ----------------------------------------------------------------------
# 教室布局编辑器
# ----------------------------------------------------------------------
class LayoutEditorDialog(tk.Toplevel):
    """
    可视化布局编辑器：
        左键单击 = 标记为学生座位（或套用当前画笔）
        右键单击 = 标记为过道 / 禁用空位
        按住左键/右键拖动 = 框选批量修改
        保存 / 加载 JSON 模板，实时显示座位统计
    """

    def __init__(self, master, layout: ClassroomLayout,
                 on_confirm: Callable[[ClassroomLayout], None]):
        super().__init__(master)
        self.title("教室布局编辑器 · 左键座位 / 右键过道 / 拖动框选")
        self.transient(master)
        fit_dialog(self, 1020, 700, min_w=760, min_h=460)
        self.on_confirm = on_confirm
        self.layout = ClassroomLayout.from_dict(layout.to_dict())
        self.brush = tk.StringVar(value=SEAT)
        self.zoom = tk.DoubleVar(value=1.0)
        self.cell_w, self.cell_h = px(72), px(48)
        self.drag_start: Optional[tuple] = None
        self.drag_button = 1
        self.drag_moved = False

        # ---- 底部信息与按钮：先 pack 到 BOTTOM，窗口再小也不会被裁掉 ----
        bottom = ttk.Frame(self)
        bottom.pack(side=tk.BOTTOM, fill=tk.X, padx=12, pady=(4, 10))
        self.info = ttk.Label(bottom, text="", style="Title.TLabel")
        self.info.pack(anchor="w", pady=(0, 6))
        # 主按钮先占位，窄窗口下也不会被裁掉
        self.confirm_button_bottom = ttk.Button(
            bottom, text="确认保存布局 →", style="Accent.TButton",
            width=18, command=self._confirm)
        self.confirm_button_bottom.pack(side=tk.RIGHT, padx=6)
        ttk.Button(bottom, text="取消", command=self.destroy).pack(side=tk.RIGHT)
        ttk.Label(bottom, text="缩放").pack(side=tk.LEFT)
        ttk.Scale(bottom, from_=0.6, to=1.6, variable=self.zoom, length=140,
                  command=lambda e: self.redraw()).pack(side=tk.LEFT, padx=(2, 12))
        ttk.Button(bottom, text="保存为教室模板", command=self._save_template).pack(side=tk.LEFT)
        ttk.Button(bottom, text="加载历史模板", command=self._load_template).pack(side=tk.LEFT, padx=6)

        # ---- 工具栏 ----
        top = ttk.LabelFrame(self, text="① 设置行列与画笔", style="Card.TLabelframe")
        top.pack(side=tk.TOP, fill=tk.X, padx=12, pady=(12, 6))
        ttk.Label(top, text="行数").grid(row=0, column=0, padx=(4, 2))
        self.row_var = tk.StringVar(value=str(self.layout.rows))
        ttk.Spinbox(top, from_=2, to=30, width=5, textvariable=self.row_var).grid(row=0, column=1)
        ttk.Label(top, text="列数").grid(row=0, column=2, padx=(12, 2))
        self.col_var = tk.StringVar(value=str(self.layout.cols))
        ttk.Spinbox(top, from_=2, to=30, width=5, textvariable=self.col_var).grid(row=0, column=3)
        ttk.Button(top, text="生成空白网格", command=self._regenerate).grid(row=0, column=4, padx=10)
        ttk.Button(top, text="全部设为座位", command=self._all_seats).grid(row=0, column=5, padx=2)
        # 顶部也放一个主按钮，避免底部被小屏幕挤掉时找不到「下一步」
        self.confirm_button_top = ttk.Button(
            top, text="完成并保存 →", style="Accent.TButton",
            width=16, command=self._confirm)
        self.confirm_button_top.grid(row=0, column=6, padx=(12, 2))

        brush_bar = ttk.Frame(top)
        brush_bar.grid(row=1, column=0, columnspan=8, sticky="w", pady=(8, 2))
        ttk.Label(brush_bar, text="画笔：").pack(side=tk.LEFT)
        for value, label in ((SEAT, "学生座位"), (BLOCK, "过道/禁用"), (PODIUM, "讲台")):
            ttk.Radiobutton(brush_bar, text=label, value=value,
                            variable=self.brush).pack(side=tk.LEFT, padx=6)
        ttk.Label(brush_bar, text="　（左键单击=画笔；右键单击=过道；按住拖动=框选批量修改）",
                  style="Hint.TLabel").pack(side=tk.LEFT)

        # ---- 画布 ----
        mid = ttk.Frame(self)
        mid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, padx=12, pady=4)
        self.canvas = tk.Canvas(mid, bg=COLOR["board"], highlightthickness=1,
                                highlightbackground="#CBD3DB")
        hsb = ttk.Scrollbar(mid, orient=tk.HORIZONTAL, command=self.canvas.xview)
        vsb = ttk.Scrollbar(mid, orient=tk.VERTICAL, command=self.canvas.yview)
        self.canvas.configure(xscrollcommand=hsb.set, yscrollcommand=vsb.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        self.canvas.bind("<ButtonPress-1>", lambda e: self._press(e, 1))
        self.canvas.bind("<ButtonPress-3>", lambda e: self._press(e, 3))
        self.canvas.bind("<B1-Motion>", self._motion)
        self.canvas.bind("<B3-Motion>", self._motion)
        self.canvas.bind("<ButtonRelease-1>", lambda e: self._release(e, 1))
        self.canvas.bind("<ButtonRelease-3>", lambda e: self._release(e, 3))
        self.canvas.bind("<MouseWheel>", self._wheel)   # Windows 滚轮

        self.redraw()
        self.grab_set()
        self.focus_set()

    # -- 坐标换算 ------------------------------------------------------
    def _geom(self):
        z = float(self.zoom.get())
        return self.cell_w * z, self.cell_h * z

    def _pos_at(self, x, y):
        w, h = self._geom()
        c = int(self.canvas.canvasx(x) // w) - 1
        r = int(self.canvas.canvasy(y) // h) - 1
        if 0 <= r < self.layout.rows and 0 <= c < self.layout.cols:
            return (r, c)
        return None

    # -- 绘制 ----------------------------------------------------------
    def redraw(self) -> None:
        self.canvas.delete("all")
        w, h = self._geom()
        pad = px(10)
        self.canvas.create_rectangle(
            pad + w * 0.15, pad, pad + w * (self.layout.cols - 0.15), pad + h * 0.72,
            fill=COLOR["podium"], outline=COLOR["podium"])
        self.canvas.create_text(pad + w * self.layout.cols / 2, pad + h * 0.36,
                                text="讲　台　（前方）", fill="white",
                                font=("微软雅黑", 11, "bold"))
        for r in range(self.layout.rows):
            for c in range(self.layout.cols):
                value = self.layout.get(r, c)
                x0 = pad + (c + 1) * w
                y0 = pad + (r + 1) * h
                if value == SEAT:
                    fill, outline, text, fg = COLOR["normal"], COLOR["normal_border"], "座位", COLOR["text"]
                elif value == BLOCK:
                    fill, outline, text, fg = COLOR["block"], COLOR["block_border"], "过道", "#5A6B7B"
                else:
                    fill, outline, text, fg = COLOR["podium"], COLOR["podium"], "讲台", "white"
                self.canvas.create_rectangle(x0 + px(3), y0 + px(3),
                                            x0 + w - px(3), y0 + h - px(3),
                                            fill=fill, outline=outline, width=1,
                                            tags=(f"cell{r}_{c}", "cell"))
                self.canvas.create_text(x0 + w / 2, y0 + h / 2, text=text, fill=fg,
                                        font=("微软雅黑", 10), tags=(f"txt{r}_{c}", "cell"))
        # 行号
        for r in range(self.layout.rows):
            self.canvas.create_text(pad + w * 0.5, pad + (r + 1.5) * h, text=f"{r + 1}排",
                                    fill=COLOR["subtext"], font=("微软雅黑", 10))
        self.canvas.configure(scrollregion=(
            0, 0, pad * 2 + (self.layout.cols + 1.5) * w,
            pad * 2 + (self.layout.rows + 1.5) * h))
        self.info.configure(text="实时预览：" + self.layout.summary())

    def _apply(self, pos, value: str) -> None:
        if pos is None:
            return
        self.layout.set_cell(pos[0], pos[1], value)
        self.redraw()

    # -- 鼠标事件 ------------------------------------------------------
    def _press(self, event, button: int) -> None:
        self.drag_start = (event.x, event.y, self.canvas.canvasx(event.x),
                           self.canvas.canvasy(event.y))
        self.drag_button = button
        self.drag_moved = False

    def _motion(self, event) -> None:
        if not self.drag_start:
            return
        x0, y0 = self.drag_start[0], self.drag_start[1]
        if abs(event.x - x0) + abs(event.y - y0) > 8:
            self.drag_moved = True
        self.canvas.delete("rubber")
        if self.drag_moved:
            self.canvas.create_rectangle(x0, y0, event.x, event.y,
                                         outline=COLOR["select"], width=2,
                                         dash=(4, 3), tags="rubber")

    def _release(self, event, button: int) -> None:
        if not self.drag_start:
            return
        value = BLOCK if button == 3 else self.brush.get()
        if self.drag_moved:
            # 框选：把矩形覆盖到的所有格子批量改成目标类型
            w, h = self._geom()
            pad = 10
            cx0, cy0 = self.drag_start[2], self.drag_start[3]
            cx1, cy1 = self.canvas.canvasx(event.x), self.canvas.canvasy(event.y)
            c0 = int(min(cx0, cx1) // w) - 1
            c1 = int(max(cx0, cx1) // w) - 1
            r0 = int(min(cy0, cy1) // h) - 1
            r1 = int(max(cy0, cy1) // h) - 1
            changed = 0
            for r in range(max(0, r0), min(self.layout.rows - 1, r1) + 1):
                for c in range(max(0, c0), min(self.layout.cols - 1, c1) + 1):
                    self.layout.set_cell(r, c, value)
                    changed += 1
            self.redraw()
            if changed:
                self.info.configure(text=f"已批量修改 {changed} 格为「{CELL_LABEL.get(value)}」｜" +
                                    self.layout.summary())
        else:
            self._apply(self._pos_at(event.x, event.y), value)
        self.canvas.delete("rubber")
        self.drag_start = None
        self.drag_moved = False

    def _wheel(self, event) -> None:
        if event.state & 0x0004:      # 按住 Ctrl 才缩放，否则普通滚动
            z = min(1.6, max(0.6, float(self.zoom.get()) + (0.08 if event.delta > 0 else -0.08)))
            self.zoom.set(z)
            self.redraw()

    # -- 按钮 ----------------------------------------------------------
    @safe_call
    def _regenerate(self) -> None:
        rows, cols = int(self.row_var.get()), int(self.col_var.get())
        self.layout.resize(rows, cols, preserve=False)
        self.layout.name = "自定义教室"
        self.redraw()
        messagebox.showinfo("已生成", f"已生成 {rows} 行 × {cols} 列的空白网格，"
                            "请用左键点座位、右键点过道。", parent=self)

    @safe_call
    def _all_seats(self) -> None:
        self.layout.fill_all(SEAT)
        self.redraw()

    @safe_call
    def _save_template(self) -> None:
        name = simpledialog.askstring("保存教室模板", "给这个布局起个名字：",
                                      initialvalue=self.layout.name, parent=self)
        if not name:
            return
        self.layout.name = name.strip()
        folder = ClassroomLayout.template_dir(BASE_DIR)
        path = os.path.join(folder, ClassroomLayout.safe_filename(name) + ".json")
        self.layout.save(path)
        messagebox.showinfo("已保存", f"教室模板已保存：\n{path}\n下次点「加载历史模板」即可复用。",
                            parent=self)

    @safe_call
    def _load_template(self) -> None:
        paths = ClassroomLayout.list_templates(BASE_DIR)
        if not paths:
            path = filedialog.askopenfilename(
                parent=self, title="选择教室模板 JSON",
                filetypes=[("教室模板", "*.json"), ("所有文件", "*.*")])
            if not path:
                return
        else:
            picker = tk.Toplevel(self)
            picker.title("历史教室模板")
            picker.transient(self)
            picker.geometry(f"{px(460)}x{px(300)}")
            ttk.Label(picker, text="双击模板名即可加载（也可以选择其它 JSON 文件）：",
                      style="Title.TLabel").pack(padx=12, pady=(12, 6), anchor="w")
            lb = tk.Listbox(picker, font=("微软雅黑", 11))
            for p in paths:
                lb.insert(tk.END, os.path.basename(p))
            lb.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
            chosen = {"path": None}

            def pick(_=None):
                sel = lb.curselection()
                if sel:
                    chosen["path"] = paths[sel[0]]
                picker.destroy()

            def browse():
                p = filedialog.askopenfilename(parent=picker, title="选择其它模板",
                                               filetypes=[("教室模板", "*.json")])
                if p:
                    chosen["path"] = p
                picker.destroy()

            lb.bind("<Double-Button-1>", pick)
            bar = ttk.Frame(picker)
            bar.pack(fill=tk.X, padx=12, pady=(2, 12))
            ttk.Button(bar, text="浏览其它文件…", command=browse).pack(side=tk.LEFT)
            ttk.Button(bar, text="取消", command=picker.destroy).pack(side=tk.RIGHT)
            picker.grab_set()
            self.wait_window(picker)
            path = chosen["path"]
            if not path:
                return
        self.layout = ClassroomLayout.load(path)
        self.row_var.set(str(self.layout.rows))
        self.col_var.set(str(self.layout.cols))
        self.redraw()
        messagebox.showinfo("加载成功", f"已加载：{self.layout.name}\n{self.layout.summary()}", parent=self)

    @safe_call
    def _confirm(self) -> None:
        if not self.layout.seats():
            messagebox.showwarning("没有座位", "至少要留一个学生座位才能排座，请用左键点几个格子。",
                                   parent=self)
            return
        self.on_confirm(self.layout)
        self.destroy()


# ----------------------------------------------------------------------
# 方案评估面板（matplotlib 嵌入 tkinter）
# ----------------------------------------------------------------------
class EvaluationPanel(ttk.Frame):
    """主窗口右下角：柱状图 + 中文统计文本，切换方案时实时刷新。"""

    def __init__(self, master):
        super().__init__(master)
        setup_chinese_font()
        bar = ttk.Frame(self)
        bar.pack(fill=tk.X, padx=6, pady=(6, 2))
        ttk.Label(bar, text="方案评估 · 各区平均分对比与分布统计",
                  style="Title.TLabel").pack(side=tk.LEFT)
        self.caption = ttk.Label(bar, text="尚未生成方案", style="Hint.TLabel")
        self.caption.pack(side=tk.RIGHT)

        body = ttk.Frame(self)
        body.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        # 柱状图原来是 100 DPI 的位图，高 DPI 屏上放大后字和柱子都发虚。
        # 这里只给基础值 130：matplotlib 的 Tk 后端会自己按系统缩放把
        # figure.dpi 再乘一次（150% 屏 → 195 DPI），画布像素 4.6×195≈897，
        # 正好匹配 px(600) 的容器；若在这里再乘 UI_SCALE 会变成 292 DPI、
        # 画布 1350 像素被 900 像素的容器裁掉一大半。
        chart_dpi = 130
        left = ttk.Frame(body, width=px(600))
        left.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)
        left.pack_propagate(False)
        self.fig = Figure(figsize=(4.6, 2.2), dpi=chart_dpi, facecolor="white")
        self.ax = self.fig.add_subplot(111)
        self.canvas = FigureCanvasTkAgg(self.fig, master=left)
        self.canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

        right = ttk.Frame(body)
        right.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(8, 0))
        self.text = tk.Text(right, wrap=tk.WORD, font=("微软雅黑", 11),
                            bg="#FBFCFD", relief=tk.FLAT, padx=px(10), pady=px(8),
                            foreground=COLOR["text"])
        sb = ttk.Scrollbar(right, orient=tk.VERTICAL, command=self.text.yview)
        self.text.configure(yscrollcommand=sb.set)
        self.text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.show_empty()

    def show_empty(self, message: str = "请先「导入学生」并点击「生成方案」，这里会显示柱状图和分布统计。") -> None:
        self.ax.clear()
        self.ax.set_facecolor("white")
        self.ax.axis("off")
        self.ax.text(0.5, 0.5, "暂无方案\n\n导入成绩表 → 生成方案", ha="center",
                     va="center", fontsize=12, color=COLOR["subtext"])
        self.canvas.draw()
        self.text.configure(state=tk.NORMAL)
        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, message)
        self.text.configure(state=tk.DISABLED)
        self.caption.configure(text="尚未生成方案")

    def show(self, plan: Plan, layout: ClassroomLayout, students: List[Student]) -> None:
        st = plan.stats or evaluate_plan(plan, layout, students)
        from layout import ZONE_NAMES
        values = [st["zone_avgs"].get(z, 0) for z in ZONE_NAMES]
        overall = st["overall_avg"]

        self.ax.clear()
        colors = ["#7FB3E8", "#4472C4", "#A9C6EA"]
        bars = self.ax.bar(ZONE_NAMES, values, color=colors,
                           edgecolor="#3559A0", linewidth=0.8, width=0.58)
        self.ax.axhline(overall, color="#D0603C", linestyle="--", linewidth=1.3,
                        label=f"班级平均 {overall:.1f}")
        low = min(values + [overall]) * 0.96 if min(values + [overall]) > 0 else 0
        high = max(values + [overall]) * 1.06 + 1
        self.ax.set_ylim(low, high)
        for bar, v in zip(bars, values):
            self.ax.text(bar.get_x() + bar.get_width() / 2, v, f"{v:.1f}",
                         ha="center", va="bottom", fontsize=9, color=COLOR["text"])
        self.ax.set_ylabel("平均分", fontsize=9)
        self.ax.set_title(f"{plan.name}：前/中/后排平均分", fontsize=11, color=COLOR["text"])
        self.ax.tick_params(labelsize=9)
        self.ax.grid(axis="y", linestyle=":", alpha=0.45)
        self.ax.set_axisbelow(True)
        for side in ("top", "right"):
            self.ax.spines[side].set_visible(False)
        self.ax.legend(fontsize=8, loc="lower right", frameon=False)
        try:
            self.fig.tight_layout()
        except Exception:
            pass
        self.canvas.draw()

        self.text.configure(state=tk.NORMAL)
        self.text.delete("1.0", tk.END)
        self.text.insert(tk.END, st["text"])
        self.text.configure(state=tk.DISABLED)
        self.caption.configure(text=f"当前：{plan.name}｜区最大差 {st['max_diff']:.1f} 分")
