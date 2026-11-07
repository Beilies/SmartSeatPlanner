# -*- coding: utf-8 -*-
"""
程序路径统一处理（app_paths.py）
==================================================================
为什么要单独写这个模块：
    老师用的免安装 exe 是用 PyInstaller 打包的，运行时会把 Python 和所有
    依赖解压到一个临时目录（_MEIxxxxx），程序退出后这个目录会被删掉。
    如果还用 __file__ 定位目录，「教室模板」「运行日志」「自检输出」就会
    写进临时目录然后丢失，老师下次打开会发现模板不见了。

    所以统一约定：
      · 源码运行时  BASE_DIR = 项目根目录（src/ 的上一级）
      · 免安装 exe 时 BASE_DIR = exe 文件所在目录（绿色软件习惯）
      · 如果 exe 放在只读位置（例如 C:\\Program Files），自动退回
        「我的文档\\SmartSeatPlanner」，保证程序永远有地方存数据。

    工程采用 src/ 布局：源码在 src/，图标在 assets/，模板与日志统一写 BASE_DIR，
    这样 `git clone` 下来的仓库根目录不会被运行产物弄脏（相关目录已在 .gitignore 里）。
"""
from __future__ import annotations

import os
import sys

# 是否被打包成了独立 exe（PyInstaller / cx_Freeze 都会设置 sys.frozen）
IS_FROZEN = bool(getattr(sys, "frozen", False))


def _project_root() -> str:
    """源码方式下定位工程根目录：src/ 的上一级，且应有 pyproject.toml。"""
    here = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(here)
    if os.path.isfile(os.path.join(parent, "pyproject.toml")):
        return parent
    # 模块被单独拷走 / pip 安装到 site-packages 时，退回当前工作目录
    return os.getcwd()


def _can_write(path: str) -> bool:
    """用一个临时探测文件判断目录是否真的可写（比只看权限位可靠）。"""
    try:
        os.makedirs(path, exist_ok=True)
        probe = os.path.join(path, ".可写测试.tmp")
        with open(probe, "w", encoding="utf-8") as f:
            f.write("ok")
        os.remove(probe)
        return True
    except Exception:
        return False


def _resolve_base_dir() -> str:
    if IS_FROZEN:
        # sys.executable 在打包后就是 exe 本身的完整路径
        candidate = os.path.dirname(os.path.abspath(sys.executable))
    else:
        candidate = _project_root()
    if _can_write(candidate):
        return candidate

    # 只读目录的兜底：改存到用户文档下
    docs = os.path.join(os.path.expanduser("~"), "Documents", "SmartSeatPlanner")
    if _can_write(docs):
        return docs
    home = os.path.join(os.path.expanduser("~"), "SmartSeatPlanner")
    try:
        os.makedirs(home, exist_ok=True)
    except Exception:
        pass
    return home


BASE_DIR = _resolve_base_dir()

# 数据子目录（模板、自检输出都放在 BASE_DIR 下，跟着 exe 走）
TEMPLATE_DIRNAME = "教室模板"
OUTPUT_DIRNAME = "自检输出"
LOG_FILENAME = "运行日志.txt"


def asset_path(filename: str) -> str:
    """
    取程序内置资源（图标等）的路径。

    打包成单文件 exe 后，图标不在 exe 旁边，而是被 PyInstaller 解压到
    临时目录 sys._MEIPASS，所以不能用 BASE_DIR 去找；退出后临时目录会被
    删除，但程序运行期间读取是安全的。源码方式运行时资源在 assets/ 目录。
    找不到时返回第一个候选路径，由调用方判断文件是否存在并安静降级。
    """
    candidates = []
    if IS_FROZEN:
        meipass = getattr(sys, "_MEIPASS", "")
        if meipass:
            candidates.append(os.path.join(meipass, filename))
        candidates.append(os.path.join(
            os.path.dirname(os.path.abspath(sys.executable)), filename))
    else:
        root = _project_root()
        candidates.append(os.path.join(root, "assets", filename))
        # 兼容「图标直接放在源码目录 / 项目根目录」的旧布局
        candidates.append(os.path.join(root, "src", filename))
        candidates.append(os.path.join(root, filename))
    for path in candidates:
        if os.path.isfile(path):
            return path
    return candidates[0]
