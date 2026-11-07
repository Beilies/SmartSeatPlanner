# -*- coding: utf-8 -*-
"""
智能学生排座位系统 · 程序入口（main.py）
==================================================================
两种运行方式：
  1) 免安装 exe（推荐给老师）：双击「智能排座位系统.exe」，什么都不用装，
     不需要 Python、不需要 pip、不需要联网，一个文件拷到哪都能用。
  2) 源码方式：双击项目根目录的 启动.bat（内部执行 python src/main.py），
     需先 pip install -r requirements.txt。

启动前会检查依赖；缺什么都用中文弹窗说明，不会抛一堆英文报错然后闪退。

隐藏参数 --diagnose：
    不打开界面，自动跑一遍「依赖检查 → 脏 Excel 导入 → 5 套排座 → 三种导出
    → 真实创建一次主窗口」，结果写入 exe 同目录「打包自检结果.txt」。
    用于发布前验证打包是否完整，也方便老师排查「双击没反应」的问题。
"""
from __future__ import annotations

import io
import os
import sys
import traceback
from datetime import datetime

# 是否运行在 PyInstaller 打包出来的独立 exe 里
IS_FROZEN = bool(getattr(sys, "frozen", False))

# 源码方式双击运行时，工作目录可能不是项目目录，先把源码所在目录加入模块搜索路径。
# 打包成 exe 后模块搜索路径由 PyInstaller 自己管理，不需要也不应该去改。
if not IS_FROZEN:
    _SRC_DIR = os.path.dirname(os.path.abspath(__file__))
    if _SRC_DIR not in sys.path:
        sys.path.insert(0, _SRC_DIR)

# --windowed 打包出来的 exe 没有控制台：sys.stdout 可能是 None，也可能是
# 一个 GBK 编码的空流。后者遇到「✅」这类字符会抛 UnicodeEncodeError，
# 让程序莫名退出。统一换成 UTF-8 的空输出，界面和日志完全不受影响。
if IS_FROZEN:
    _stdout_enc = (getattr(sys.stdout, "encoding", None) or "").lower().replace("-", "")
    if sys.stdout is None or _stdout_enc != "utf8":
        sys.stdout = open(os.devnull, "w", encoding="utf-8")
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w", encoding="utf-8")

from app_paths import BASE_DIR, LOG_FILENAME

INSTALL_CMD = "pip install -r requirements.txt"
LOG_FILE = os.path.join(BASE_DIR, LOG_FILENAME)
DIAG_FILE = os.path.join(BASE_DIR, "打包自检结果.txt")

# (导入名, 说明, 是否必须)
DEPENDENCIES = [
    ("tkinter", "图形界面（官方 Python 自带）", True),
    ("pandas", "数据表格处理", True),
    ("openpyxl", "读写 .xlsx / 导出 Excel", True),
    ("xlrd", "读取老式 .xls 成绩表", True),
    ("matplotlib", "统计柱状图 / 导出 PNG 座位图", True),
]


def _write_log(text: str) -> None:
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"\n===== {datetime.now():%Y-%m-%d %H:%M:%S} =====\n{text}\n")
    except Exception:
        pass


def _missing_tip(missing: list[str]) -> str:
    """缺组件时的中文说明：免安装版和源码版给老师的操作完全不同。"""
    if IS_FROZEN:
        return (
            "这个程序文件不完整，缺少内部组件：\n\n" + "\n".join(missing) +
            "\n\n【怎么办】\n"
            "请重新拷贝一份完整的「智能排座位系统.exe」（只要这一个文件，"
            "不要只复制桌面的快捷方式），或请发放程序的老师重新发一份。\n\n"
            "免安装版本不需要安装 Python、不需要联网、也不需要敲任何命令。"
        )
    return (
        "这个程序还缺少以下运行组件：\n\n" + "\n".join(missing) +
        "\n\n【安装办法】按 Win+R 输入 cmd 回车，在黑窗口里执行：\n" +
        f"    {INSTALL_CMD}\n\n" +
        "如果提示 pip 不是内部命令，说明 Python 没装好：\n"
        "请到 python.org 下载官方 Python 3.10 及以上版本，安装时勾选 "
        "「Add Python to PATH」和「tcl/tk and IDLE」（tkinter 就会一起装上）。\n\n"
        "装完重新双击本文件即可；也可以直接改用免安装版 exe，什么都不用装。"
    )


def check_dependencies() -> bool:
    """返回 True 表示依赖齐全；缺依赖时弹中文提示并返回 False。"""
    missing = []
    for module, desc, required in DEPENDENCIES:
        try:
            __import__(module)
        except Exception:
            if required:
                missing.append(f"  · {module}：{desc}")
    if not missing:
        return True

    tip = _missing_tip(missing)
    _write_log("缺少依赖：\n" + "\n".join(missing))
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("缺少运行组件", tip)
        root.destroy()
    except Exception:
        print(tip)
        try:
            input("\n按回车键退出…")
        except Exception:
            pass
    return False


def diagnose() -> int:
    """
    隐藏参数 --diagnose 的实现：不开窗口，验证「打包是否完整」。
    检查 ① 所有依赖能导入 ② 导入/排座/导出全链路能跑 ③ 真实创建主窗口。
    结果同时打印和写入「打包自检结果.txt」，返回 0 表示全部通过。
    """
    lines: list[str] = [
        "智能学生排座位系统 · 完整性自检",
        f"时间：{datetime.now():%Y-%m-%d %H:%M:%S}",
        "运行模式：" + ("免安装 exe（绿色版）" if IS_FROZEN else "源码方式"),
        "程序目录：" + BASE_DIR,
        "",
        "【1】运行组件",
    ]
    ok = True
    for mod in ("tkinter", "pandas", "numpy", "openpyxl", "xlrd",
                "matplotlib", "PIL"):
        try:
            m = __import__(mod)
            lines.append(f"    [正常] {mod} {getattr(m, '__version__', '')}".rstrip())
        except Exception as exc:
            lines.append(f"    [缺失] {mod}：{exc}")
            ok = False

    lines.append("")
    lines.append("【2】全链路自检（脏 Excel → 5 套方案 → 三种导出）")
    buf = io.StringIO()
    try:
        import selfcheck
        old_stdout, sys.stdout = sys.stdout, buf
        try:
            selfcheck.main()
        finally:
            sys.stdout = old_stdout
        lines.append(buf.getvalue().strip())
    except Exception:
        lines.append(traceback.format_exc())
        ok = False

    lines.append("")
    lines.append("【3】图形界面")
    try:
        from gui_main import MainWindow
        app = MainWindow()
        app.update()
        lines.append(f"    [正常] 主窗口创建成功（{app.winfo_width()} × {app.winfo_height()}），"
                     "座位画布、控制面板、统计图表均已加载")
        lines.append(f"    [正常] 窗口标题：{app.title()}")
        lines.append(f"    [正常] 默认布局：{app.layout.summary()}")

        # 回归检查：左侧控制面板的滚轮在按钮、右侧滚动条上都必须能滚动，
        # 滚轮不能误改「方案数量」，也不能把右侧座位画布的滚轮抢过来。
        def first_of(widget, cls):
            if widget.winfo_class() == cls:
                return widget
            for child in widget.winfo_children():
                found = first_of(child, cls)
                if found is not None:
                    return found
            return None

        def wheel(widget, delta=-120, rootx=None, rooty=None):
            widget.event_generate(
                "<MouseWheel>", delta=delta, x=5, y=5,
                rootx=rootx if rootx is not None else widget.winfo_rootx() + 5,
                rooty=rooty if rooty is not None else widget.winfo_rooty() + 5)
            app.update()

        app.geometry("1100x680")      # 保证面板内容一定超出可视高度，测得出滚动
        app.update()
        left = app.left_canvas
        left.yview_moveto(0)
        wheel_checks = []
        for label, target in (("按钮", first_of(app.left_panel, "TButton")),
                              ("右侧滚动条", first_of(app.left_outer, "TScrollbar"))):
            before = left.yview()[0]
            if target is not None:
                wheel(target)
            wheel_checks.append((label, target is not None and left.yview()[0] > before))
        spin = first_of(app.left_panel, "TSpinbox")
        count_before = app.var_plan_count.get()
        if spin is not None:
            wheel(spin)
        wheel_checks.append(("不误改数值", spin is not None
                             and app.var_plan_count.get() == count_before))
        before = left.yview()[0]
        wheel(app.board.canvas)
        wheel_checks.append(("右侧画布不误滚", abs(left.yview()[0] - before) < 1e-9))
        wheel_ok = all(v for _, v in wheel_checks)
        lines.append("    [%s] 左侧面板滚轮：%s"
                     % ("正常" if wheel_ok else "异常",
                        "，".join(f"{k}{'✓' if v else '✗'}" for k, v in wheel_checks)))
        if not wheel_ok:
            ok = False

        # 回归检查：故意把导入预览压成很小的窗口（模拟小屏幕 / 高 DPI 缩放），
        # 顶部和底部的「确认导入」按钮都必须仍然可见，不能跑到窗口外。
        from excel_io import ImportResult
        from gui_dialogs import ImportPreviewDialog
        from models import Student
        fake = ImportResult(
            students=[Student(name=f"学生{i:02d}", score=300 + i)
                      for i in range(1, 61)],
            warnings=[], sheet_name="测试", path="内存测试数据.xlsx",
            header_row=1, subject_names=[], ignored_rows=0)
        dlg = ImportPreviewDialog(app, fake, lambda s, w: None)
        dlg.minsize(400, 300)
        dlg.geometry("640x420")
        dlg.update()
        win_w, win_h = dlg.winfo_width(), dlg.winfo_height()
        checks = []
        for label, btn in (("顶部", dlg.confirm_button_top),
                           ("底部", dlg.confirm_button_bottom)):
            rel_x = btn.winfo_rootx() - dlg.winfo_rootx()
            rel_y = btn.winfo_rooty() - dlg.winfo_rooty()
            # 水平和垂直都要检查：窄窗口会裁右侧，矮窗口会裁底部
            visible = (bool(btn.winfo_ismapped())
                       and rel_x >= 0 and rel_y >= 0
                       and rel_x + btn.winfo_width() <= win_w + 2
                       and rel_y + btn.winfo_height() <= win_h + 2)
            checks.append((label, visible, rel_x, rel_y))
        all_visible = all(v for _, v, _, _ in checks)
        lines.append(
            f"    [{'正常' if all_visible else '异常'}] 导入预览按钮在 640×420 小窗口下仍可见"
            f"（{'，'.join(f'{k} ({x},{y})' for k, v, x, y in checks)}）")
        if not all_visible:
            ok = False
        dlg.destroy()
        app.destroy()
    except Exception:
        lines.append(traceback.format_exc())
        ok = False

    lines.append("")
    lines.append("【4】程序图标资源")
    try:
        from app_paths import asset_path
        for fname in ("app.ico", "app_logo.png"):
            p = asset_path(fname)
            exists = os.path.isfile(p)
            size = os.path.getsize(p) if exists else 0
            lines.append(f"    [{'正常' if exists else '缺失'}] {fname}"
                         f"（{size:,} 字节）")
            if not exists:
                ok = False
    except Exception:
        lines.append(traceback.format_exc())
        ok = False

    lines += ["", "结论：" + ("全部通过，这个程序可以正常发放给老师使用。"
                            if ok else "发现问题，请把本文件发给技术人员。")]
    text = "\n".join(lines)
    try:
        with open(DIAG_FILE, "w", encoding="utf-8") as f:
            f.write(text + "\n")
        lines.append(f"结果已保存到：{DIAG_FILE}")
    except Exception:
        pass
    try:
        print("\n".join(lines))
    except Exception:
        pass  # 免安装版没有控制台，结果已经写进文件，打印失败不影响结论
    return 0 if ok else 1


def main() -> None:
    if not check_dependencies():
        return
    try:
        from gui_main import MainWindow
        app = MainWindow()
        app.mainloop()
    except Exception as exc:
        # 启动阶段的意外也只写日志 + 中文提示，不把堆栈甩给老师
        _write_log(traceback.format_exc())
        if IS_FROZEN:
            message = (
                "程序启动时遇到问题，常见原因：系统缺少图形界面组件、"
                "程序文件被安全软件破坏，或系统字体异常。\n\n"
                "建议：重新拷贝一份完整的 exe 再双击；如果一直不行，"
                "把同目录的「运行日志.txt」发给技术人员。\n"
                "（免安装版本不需要安装任何东西。）"
            )
        else:
            message = (
                "程序启动时遇到问题，常见原因：桌面环境不支持图形窗口、"
                "依赖版本过旧，或配置文件损坏。\n\n"
                "建议：重新执行 " + INSTALL_CMD + " 后重试；"
                "详细记录已写入「运行日志.txt」，可发给技术人员。"
            )
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("启动未完成", message)
            root.destroy()
        except Exception:
            print(message)
            try:
                input("\n按回车键退出…")
            except Exception:
                pass
        print(f"（内部信息：{type(exc).__name__}: {exc}）")


if __name__ == "__main__":
    if "--diagnose" in sys.argv:
        sys.exit(diagnose())
    main()
