# -*- coding: utf-8 -*-
"""
界面清晰度 / 高 DPI 适配（ui_scale.py）
==================================================================
Windows 在「显示缩放 125% / 150%」的笔记本上，会把没有声明 DPI 感知的
程序先按 96 DPI 画成一张位图，再整体放大到屏幕上——于是按钮、文字、
座位格全都发虚，看起来像「分辨率不够」。这是 tkinter 程序最常见的模糊原因，
和屏幕本身的分辨率无关。

本模块做三件事，且必须在「创建任何窗口之前」完成（模块导入时即执行）：
  1. 声明进程级 DPI 感知（优先 Per-Monitor V2）：让系统按真实物理像素绘制，
     文字交给 ClearType 直接渲染，不再经过位图拉伸。
  2. 读出真实系统 DPI，算出统一缩放系数 UI_SCALE，并提供 px() 把设计稿里的
     逻辑像素换成物理像素（面板宽度、格子大小、行高、弹窗尺寸…）。
  3. apply_tk_scaling()：把 Tk 的「每点对应多少像素」从固定 96/72 提到
     (96/72)×UI_SCALE，所有以「点」为单位的微软雅黑字体自动等比放大，
     视觉大小和以前一致，但笔画是重新光栅化的，因而清晰锐利。

老师如果想手动微调，可以设置环境变量 SEAT_UI_SCALE（例如 1.25 / 1.5），
不设置就完全跟随 Windows 的显示缩放。非 Windows 平台静默按 1.0 处理。
"""
from __future__ import annotations

import ctypes
import os
import sys

IS_WINDOWS = sys.platform == "win32"
MIN_SCALE, MAX_SCALE = 1.0, 2.0          # 太小看不清，太大会撑爆窗口

DPI_AWARE = False                        # 是否不再被 Windows 整窗位图拉伸
DPI_LEVEL = -1                           # -1 未感知 / 0 系统级 / 1 每显示器 / 2 V2
DPI_MODE = ""                            # 感知模式，写进 --diagnose 报告
SYSTEM_DPI = 96                          # 真实系统 DPI（感知声明后读取）
UI_SCALE = 1.0                           # 界面统一缩放系数


def _env_scale(fallback: float) -> float:
    """允许用环境变量覆盖缩放（方便测试 / 老师手动微调），非法值自动忽略。"""
    raw = os.environ.get("SEAT_UI_SCALE", "").strip()
    try:
        value = float(raw) if raw else fallback
    except ValueError:
        value = fallback
    return min(MAX_SCALE, max(MIN_SCALE, value))


def _awareness_level() -> int:
    """读取当前线程 DPI 感知级别：-1 未感知 / 0 系统级 / 1 每显示器 / 2 V2。

    ctypes 默认把返回值当 32 位 int，而 DPI_AWARENESS_CONTEXT 是 64 位句柄，
    不显式声明 restype / argtypes 会被截断，导致级别永远读不出来。
    """
    try:
        u = ctypes.windll.user32
        u.GetThreadDpiAwarenessContext.restype = ctypes.c_void_p
        u.GetAwarenessFromDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        u.GetAwarenessFromDpiAwarenessContext.restype = ctypes.c_int
        ctx = u.GetThreadDpiAwarenessContext()
        return int(u.GetAwarenessFromDpiAwarenessContext(ctx))
    except Exception:
        return -1


def _init() -> None:
    global DPI_AWARE, DPI_LEVEL, DPI_MODE, SYSTEM_DPI, UI_SCALE
    if not IS_WINDOWS:
        DPI_MODE = "非 Windows 平台，按 100% 缩放"
        UI_SCALE = _env_scale(1.0)
        return

    u = ctypes.windll.user32
    # ① Win10 1703+ 最佳：Per-Monitor V2，跨不同缩放的显示器也清晰
    try:
        # -4 = DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2（必须按指针传，不能按 int）
        u.SetProcessDpiAwarenessContext.argtypes = [ctypes.c_void_p]
        u.SetProcessDpiAwarenessContext.restype = ctypes.c_int
        u.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
    except Exception:
        pass
    level = _awareness_level()
    # ② 旧版 Win8.1 / Win10：shcore 的 Per-Monitor Aware
    if level < 1:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(2)
        except Exception:
            pass
        level = _awareness_level()
    # ③ Vista 之后都有的系统级 DPI 感知（单显示器够用）
    if level < 0:
        try:
            u.SetProcessDPIAware()
        except Exception:
            pass
        level = _awareness_level()

    # 0「系统级」及以上都不会被 Windows 整窗拉伸；只有 -1 会糊。
    # PyInstaller 生成的 exe 可能自带 dpiAware manifest，此时 API 调用会
    # 返回「已设置」，但 level 已经是 0/2，同样算生效，不能误报失败。
    DPI_LEVEL = level
    DPI_AWARE = level >= 0
    DPI_MODE = {2: "Per-Monitor V2（最佳）", 1: "Per-Monitor Aware",
                0: "System DPI Aware（exe 自带清单）",
                -1: "未启用（系统仍会拉伸窗口）"}.get(level, f"级别 {level}")

    # 声明感知之后读 DPI 才是真实值；未声明时 Windows 会骗我们说永远是 96
    try:
        SYSTEM_DPI = int(u.GetDpiForSystem())
    except Exception:
        try:
            hdc = u.GetDC(0)
            SYSTEM_DPI = int(ctypes.windll.gdi32.GetDeviceCaps(hdc, 88))  # LOGPIXELSX
            u.ReleaseDC(0, hdc)
        except Exception:
            SYSTEM_DPI = 96
    SYSTEM_DPI = max(72, min(288, SYSTEM_DPI))
    UI_SCALE = _env_scale(SYSTEM_DPI / 96.0)


_init()                                   # import 即生效，早于任何 tk.Tk()


def px(value: float) -> int:
    """设计稿里的逻辑像素 → 当前屏幕的物理像素（用于宽度 / 行高 / 间距）。"""
    return max(1, int(round(value * UI_SCALE)))


def apply_tk_scaling(root) -> float:
    """在创建任何控件之前调用：让 Tk 按真实 DPI 重新光栅化所有点数字体。"""
    target = (96.0 / 72.0) * UI_SCALE
    try:
        root.tk.call("tk", "scaling", target)
        return float(root.tk.call("tk", "scaling"))
    except Exception:
        return 96.0 / 72.0


def window_size(root, base_w: int, base_h: int):
    """主窗口按缩放放大，但不超出屏幕（笔记本 / 投影仪也不会被裁掉）。"""
    try:
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    except Exception:
        sw, sh = 1920, 1080
    w = min(px(base_w), max(720, int(sw * 0.94)))
    h = min(px(base_h), max(520, int(sh * 0.92)))
    return w, h


def minimum_size(root, base_w: int, base_h: int):
    """最小窗口同样按缩放放大，但保证不超过屏幕的 80%。"""
    try:
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
    except Exception:
        sw, sh = 1920, 1080
    return (min(px(base_w), max(640, int(sw * 0.80))),
            min(px(base_h), max(480, int(sh * 0.80))))


def describe() -> list[str]:
    """给 --diagnose 用的中文说明，方便老师/开发者确认清晰度设置生效。"""
    return [
        f"DPI 感知：{DPI_MODE}（已生效：{'是' if DPI_AWARE else '否'}）",
        f"系统 DPI：{SYSTEM_DPI}（Windows 显示缩放约 {SYSTEM_DPI / 96 * 100:.0f}%）",
        f"界面缩放系数：{UI_SCALE:.2f}×；Tk 字体按 (96/72)×缩放重新渲染",
        "字体：全界面统一「微软雅黑」，正文基础字号由 10pt 提到 11pt",
        "座位图 PNG：300 DPI 输出；评估柱状图：按屏幕缩放后的高 DPI 渲染",
    ]
