# -*- coding: utf-8 -*-
"""把 src/ 下的源码汇总成一个便于阅读 / 备份的 txt（docs/完整代码汇总.txt）。

用法（项目根目录）：
    python scripts/生成代码汇总.py
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"
OUT = ROOT / "docs" / "完整代码汇总.txt"
SEP = "=" * 78

# (显示名, 相对路径) —— 顺序即阅读顺序：入口 → 基础 → 业务 → 界面 → 自检
FILES = [
    ("src/main.py", "main.py"),
    ("src/app_paths.py", "app_paths.py"),
    ("src/models.py", "models.py"),
    ("src/excel_io.py", "excel_io.py"),
    ("src/layout.py", "layout.py"),
    ("src/seating.py", "seating.py"),
    ("src/evaluate.py", "evaluate.py"),
    ("src/exporter.py", "exporter.py"),
    ("src/gui_main.py", "gui_main.py"),
    ("src/gui_dialogs.py", "gui_dialogs.py"),
    ("src/selfcheck.py", "selfcheck.py"),
    ("requirements.txt", "requirements.txt"),
    ("启动.bat", "启动.bat"),
]


def main() -> None:
    parts = [
        "智能学生排座位系统 · 完整代码汇总",
        SEP,
        "作者：北离Beilies",
        "生成时间：" + datetime.now().strftime("%Y-%m-%d %H:%M"),
        "说明：本文件仅供阅读 / 备份使用，由 scripts/生成代码汇总.py 自动生成。",
        "      老师使用请到 GitHub Releases 下载免安装 exe；源码方式请看 README.md。",
        "      工程采用 src/ 布局，各模块职责见 docs/开发文档.md。",
        "",
        "",
    ]
    total = 0
    for i, (display, rel) in enumerate(FILES, 1):
        # Windows 批处理用 GBK（中文控制台 936），其余文件用 UTF-8
        enc = "gbk" if rel.endswith(".bat") else "utf-8-sig"
        path = (SRC / rel) if rel.endswith(".py") else (ROOT / rel)
        lines = path.read_text(encoding=enc).splitlines()
        total += len(lines)
        parts += [SEP, f"【文件 {i}/{len(FILES)}】{display}　（{len(lines)} 行）", SEP]
        parts += lines
        parts.append("")

    parts += [SEP, f"合计 {len(FILES)} 个文件、约 {total} 行（含注释）。",
              "依赖安装：pip install -r requirements.txt", SEP]
    OUT.write_text("\n".join(parts) + "\n", encoding="utf-8-sig")
    print(f"已生成 {OUT.relative_to(ROOT)}，源码合计 {total} 行")


if __name__ == "__main__":
    main()
