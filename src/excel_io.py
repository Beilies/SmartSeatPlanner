# -*- coding: utf-8 -*-
"""
Excel 导入模块（excel_io.py）
------------------------------------------------------------------
职责：
    1. 列出 Sheet；用 pandas + openpyxl(.xlsx) / xlrd(.xls) 读取成二维表
    2. 自动模糊匹配列名：学号 / 姓名 / 总分 / 性别 / 特殊标记 /
       禁止同桌名单 / 各科成绩
    3. 容错：表头不在第一行、合并单元格、空行、底部合计行、
       缺考与 N/A 等缺失值（绝不按 0 分）、异常分数、重复学生
    4. 生成标准 Excel 模板（老师可下载后填表）
    5. 预览窗口编辑后的表格文本 -> 学生对象（双向转换）

为什么不用 pandas 直接 header=0：
    成绩表经常前面有几行标题、说明、合并单元格，直接读会把说明当表头。
    这里先读成无表头矩阵，自己扫描前 20 行找真正的表头行，
    再用 openpyxl / xlrd 把合并单元格的值补进矩阵，容错能力更强。
"""
from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from models import MISSING_TOKENS, Student, WarningItem

# 预览窗口里老师可以直接编辑的列（顺序固定，不要随意调整）
PREVIEW_COLUMNS = ["学号", "姓名", "总分", "性别", "特殊标记", "禁止同桌名单", "各科成绩"]

# 模板里的标准字段
TEMPLATE_COLUMNS = ["学号", "姓名", "总分", "性别", "特殊标记", "禁止同桌名单",
                    "语文", "数学", "英语"]


class ExcelParseError(Exception):
    """Excel 解析失败。上层 GUI 会捕获并转成中文弹窗，不向老师抛原始异常。"""


# ----------------------------------------------------------------------
# 通用小工具
# ----------------------------------------------------------------------
def is_blank(value) -> bool:
    """判断是否为空值 / 缺考标记。注意：缺考标记不能按 0 分处理。"""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        pass
    return str(value).strip().lower() in MISSING_TOKENS


def as_text(value) -> str:
    """把单元格转成干净文本；NaN / None / 缺考标记统一返回空串。"""
    if is_blank(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def clean_score(value) -> Tuple[Optional[float], str]:
    """
    分数清洗：返回 (分数, 状态)，状态 = ok / missing / invalid
    - missing：空、—、N/A、缺考 等，返回 None（后续绝不当 0 分）
    - invalid：负数、超过 1000、无法识别的文本，返回 None 并告警
    """
    if is_blank(value):
        return None, "missing"
    if isinstance(value, bool):
        return None, "invalid"
    if isinstance(value, (int, float)):
        num = float(value)
    else:
        s = str(value).strip()
        # 全角数字 / 全角小数点 / 中文「分」/ 千分位逗号 的容错
        s = s.translate(str.maketrans("０１２３４５６７８９．－＋", "0123456789.-+"))
        s = s.replace("，", "").replace(",", "").replace("分", "").strip()
        if not re.fullmatch(r"[-+]?\d+(\.\d+)?", s):
            return None, "invalid"
        num = float(s)
    if num < 0 or num > 1000:
        return None, "invalid"
    return num, "ok"


def normalize_gender(value) -> str:
    s = as_text(value).lower()
    if s in ("男", "m", "male", "1", "男生", "男孩", "男同学"):
        return "男"
    if s in ("女", "f", "female", "2", "女生", "女孩", "女同学"):
        return "女"
    return ""


def split_forbid(value) -> List[str]:
    """禁止同桌名单：支持 顿号 / 逗号 / 分号 / 斜杠 / 空格 分隔。"""
    s = as_text(value)
    if not s or s.lower() in ("无", "没有", "none"):
        return []
    parts = re.split(r"[、,，;；/|\s]+", s)
    return [p for p in (x.strip() for x in parts) if p]


def parse_subjects_text(text: str) -> Dict[str, Optional[float]]:
    """预览表格里「语文=90;数学=85」这样的文本 -> 科目分数字典。"""
    out: Dict[str, Optional[float]] = {}
    if not text or not text.strip():
        return out
    for chunk in re.split(r"[;；,，\n]+", text):
        chunk = chunk.strip()
        if not chunk:
            continue
        m = re.match(r"^([^=:]+?)\s*[=:]\s*(.+)$", chunk)
        if not m:
            continue
        subject = m.group(1).strip()
        val, _ = clean_score(m.group(2).strip())
        if subject:
            out[subject] = val
    return out


def subjects_to_text(subjects: Dict[str, Optional[float]]) -> str:
    parts = []
    for k, v in subjects.items():
        parts.append(f"{k}={'缺考' if v is None else f'{v:g}'}")
    return ";".join(parts)


# ----------------------------------------------------------------------
# 一、读取 Excel
# ----------------------------------------------------------------------
@dataclass
class ColumnMapping:
    """表头行中各字段所在的列号（0 开始）。None 表示该列不存在。"""
    sid: Optional[int] = None
    name: Optional[int] = None
    score: Optional[int] = None
    gender: Optional[int] = None
    tag: Optional[int] = None
    forbid: Optional[int] = None
    subjects: Dict[int, str] = field(default_factory=dict)  # 列号 -> 科目名


@dataclass
class ImportResult:
    """一次导入的完整结果，交给预览窗口使用。"""
    students: List[Student]
    warnings: List[WarningItem]
    sheet_name: str
    path: str
    header_row: int                 # 表头在原表的第几行（1 开始），给老师看
    subject_names: List[str]
    ignored_rows: int


# 字段模糊匹配规则（按优先级从上到下匹配，命中即停）
FIELD_PATTERNS = [
    ("forbid", re.compile(r"禁止|黑名单|不同桌|不可同桌|避免同桌|不相邻|回避|forbid|ban", re.I)),
    ("sid", re.compile(r"学号|编号|考号|座号|序号|学籍|id\b|no\.?$|number", re.I)),
    ("name", re.compile(r"姓名|名字|name|学生名", re.I)),
    ("score", re.compile(r"总分|总成绩|总分数|合计|总评|成绩总和|total|sum", re.I)),
    ("gender", re.compile(r"性别|男女|gender|sex", re.I)),
    ("tag", re.compile(r"特殊|标记|备注|关注|标签|情况|类型|tag|note|remark|说明", re.I)),
    ("subject", re.compile(r"语文|数学|英语|外语|物理|化学|生物|政治|历史|地理|科学|"
                           r"道法|品德|信息|体育|音乐|美术|综合|科目|学科", re.I)),
]

# 这些列即便全是数字也不是科目
NOT_SUBJECT = re.compile(r"名次|排名|序号|班级|年龄|电话|学号|编号|考勤|次数", re.I)
SUMMARY_ROW = re.compile(r"合计|总计|平均|小计|统计|总人数|最高分|最低分|年级总", re.I)


def is_supported_file(path: str) -> bool:
    return os.path.splitext(path)[1].lower() in (".xlsx", ".xls")


def list_sheets(path: str) -> List[str]:
    """列出所有 Sheet 名。文件损坏 / 缺依赖时抛中文错误。"""
    if not os.path.isfile(path):
        raise ExcelParseError("找不到这个 Excel 文件，请重新选择")
    if not is_supported_file(path):
        raise ExcelParseError("目前只支持 .xlsx 和 .xls 格式，请先另存为这两种格式")
    try:
        with pd.ExcelFile(path) as book:
            return list(book.sheet_names)
    except ImportError:
        raise ExcelParseError("读取 .xls 需要 xlrd 组件，请按使用说明安装依赖后重试")
    except Exception as exc:
        raise ExcelParseError(f"打开 Excel 失败：{_brief(exc)}")


def _brief(exc: Exception) -> str:
    """把底层异常压成一句中文可读的短说明，不暴露堆栈。"""
    msg = str(exc).strip().splitlines()[0] if str(exc).strip() else type(exc).__name__
    return msg[:110]


def read_matrix(path: str, sheet_name: str) -> pd.DataFrame:
    """读成「无表头」二维 DataFrame，并补齐合并单元格的值。"""
    ext = os.path.splitext(path)[1].lower()
    engine = "xlrd" if ext == ".xls" else "openpyxl"
    try:
        df = pd.read_excel(path, sheet_name=sheet_name, header=None,
                           engine=engine, dtype=object)
    except ImportError:
        raise ExcelParseError("缺少读取组件，请安装 xlrd（.xls）或 openpyxl（.xlsx）")
    except Exception as exc:
        raise ExcelParseError(f"读取工作表「{sheet_name}」失败：{_brief(exc)}")
    if df.empty:
        raise ExcelParseError(f"工作表「{sheet_name}」是空的，请检查文件内容")

    # --- 展开合并单元格：只有左上角格子有值，其余格子需补成同样的值 ---
    try:
        if ext == ".xlsx":
            wb = load_workbook(path, data_only=True, read_only=False)
            ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]
            for rng in ws.merged_cells.ranges:
                value = ws.cell(rng.min_row, rng.min_col).value
                if value is None:
                    continue
                for r in range(rng.min_row, rng.max_row + 1):
                    for c in range(rng.min_col, rng.max_col + 1):
                        if r - 1 < df.shape[0] and c - 1 < df.shape[1]:
                            df.iat[r - 1, c - 1] = value
            wb.close()
        else:
            import xlrd  # .xls 专用
            book = xlrd.open_workbook(path, formatting_info=True)
            sh = book.sheet_by_name(sheet_name)
            for rlo, rhi, clo, chi in sh.merged_cells:
                value = sh.cell_value(rlo, clo)
                for r in range(rlo, rhi):
                    for c in range(clo, chi):
                        if r < df.shape[0] and c < df.shape[1]:
                            df.iat[r, c] = value
    except Exception:
        # 合并单元格信息读不到不影响主流程，静默降级（数据仍可用）
        pass
    return df


def detect_header(df: pd.DataFrame, max_scan: int = 20) -> Tuple[int, ColumnMapping]:
    """扫描前若干行，找出最像表头的一行，并做列名模糊匹配。"""
    best_row, best_score, best_mapping = -1, -1, None
    for r in range(min(max_scan, len(df))):
        texts = [as_text(df.iat[r, c]) for c in range(df.shape[1])]
        non_empty = [t for t in texts if t]
        if len(non_empty) < 2:
            continue
        mapping = _match_columns(texts)
        score = _header_score(texts, mapping)
        # 没有姓名列的行（例如大标题行）不可能是表头，直接降权
        if mapping.name is None:
            score -= 20
        if score > best_score:
            best_row, best_score, best_mapping = r, score, mapping

    if best_row < 0 or best_mapping is None:
        raise ExcelParseError("表格是空的或前 20 行里找不到表头")
    if best_mapping.name is None or best_score < 10:
        raise ExcelParseError(
            "没能识别出「姓名」列。请确认表头里有一列写着姓名（或学生姓名/Name），"
            "也可以先下载标准模板对照填写")
    return best_row, best_mapping


def _match_columns(texts: List[str]) -> ColumnMapping:
    """对一行文本做字段匹配，生成 ColumnMapping。"""
    mapping = ColumnMapping()
    occupied = set()
    for c, text in enumerate(texts):
        if not text:
            continue
        kind = None
        for key, pat in FIELD_PATTERNS:
            if pat.search(text):
                kind = key
                break
        if kind == "name" and "name" not in occupied:
            mapping.name = c
            occupied.add("name")
        elif kind == "sid" and mapping.sid is None:
            mapping.sid = c
        elif kind == "score" and mapping.score is None:
            mapping.score = c
        elif kind == "gender" and mapping.gender is None:
            mapping.gender = c
        elif kind == "tag" and mapping.tag is None:
            mapping.tag = c
        elif kind == "forbid" and mapping.forbid is None:
            mapping.forbid = c
        elif kind == "subject":
            mapping.subjects[c] = text
    # 未识别但全是数字、且列名不像「名次/班级」的列，也当作科目（兼容自定义科目名）
    return mapping


def _header_score(texts: List[str], mapping: ColumnMapping) -> int:
    weight = {"name": 10, "sid": 2, "score": 2, "gender": 1,
              "tag": 1, "forbid": 1, "subject": 1}
    score = 0
    for c, text in enumerate(texts):
        if not text:
            continue
        for key, pat in FIELD_PATTERNS:
            if pat.search(text):
                score += weight[key]
                break
    return score


def _guess_numeric_subjects(df: pd.DataFrame, header_row: int,
                            mapping: ColumnMapping) -> None:
    """把「表头没写科目名、但整列都是分数」的列补成科目列。"""
    n_rows = len(df)
    for c in range(df.shape[1]):
        if c in (mapping.sid, mapping.name, mapping.score, mapping.gender,
                 mapping.tag, mapping.forbid) or c in mapping.subjects:
            continue
        header_text = as_text(df.iat[header_row, c]) if header_row < n_rows else ""
        if NOT_SUBJECT.search(header_text):
            continue
        vals = [df.iat[r, c] for r in range(header_row + 1, n_rows)]
        vals = [v for v in vals if not is_blank(v)]
        if len(vals) < 2:
            continue
        numeric = sum(1 for v in vals if clean_score(v)[1] in ("ok", "missing"))
        if numeric / len(vals) >= 0.6:
            mapping.subjects[c] = header_text or f"科目{get_column_letter(c + 1)}"


# ----------------------------------------------------------------------
# 二、矩阵 -> 学生对象
# ----------------------------------------------------------------------
def build_students(df: pd.DataFrame, header_row: int,
                   mapping: ColumnMapping) -> Tuple[List[Student], List[WarningItem], int, List[str]]:
    """把数据区逐行转成 Student，同时收集所有告警。"""
    if mapping.name is None:
        raise ExcelParseError("没有找到姓名列，无法生成学生名单")
    _guess_numeric_subjects(df, header_row, mapping)
    students: List[Student] = []
    warnings: List[WarningItem] = []
    ignored = 0
    auto_filled = 0                      # 总分由各科自动求和的人数
    subject_names = [mapping.subjects[c] for c in sorted(mapping.subjects)]

    def cell(r: int, c: Optional[int]):
        if c is None or c >= df.shape[1]:
            return ""
        return df.iat[r, c]

    for r in range(header_row + 1, len(df)):
        row_values = [df.iat[r, c] for c in range(df.shape[1])]
        if all(is_blank(v) for v in row_values):
            continue  # 空行

        line_no = r + 1  # Excel 里显示的行号
        sid = as_text(cell(r, mapping.sid))
        name = as_text(cell(r, mapping.name))

        # 底部合计 / 平均分 之类的统计行：直接忽略，不进入预览表。
        # 「合计」常写在学号列或姓名列，所以只检查学号/姓名/总分这三格，
        # 避免学生备注里出现「平均分」之类字样时被误删。
        stat_cells = (sid, name, as_text(cell(r, mapping.score)))
        if any(SUMMARY_ROW.search(x) for x in stat_cells) and \
                (not name or SUMMARY_ROW.search(name)):
            ignored += 1
            continue
        # 只有一个单元格有内容、又没有姓名/学号的杂项行，也忽略
        if not name and not sid and sum(1 for v in row_values if not is_blank(v)) <= 1:
            ignored += 1
            continue

        if not name:
            warnings.append(WarningItem("empty_name",
                                        "这一行姓名为空，已保留在预览表里，请补姓名或直接删除该行",
                                        line_no))

        # --- 各科成绩 ---
        subjects: Dict[str, Optional[float]] = {}
        absent_subjects: List[str] = []
        for c, sub_name in mapping.subjects.items():
            val, status = clean_score(cell(r, c))
            subjects[sub_name] = val
            if status == "missing":
                absent_subjects.append(sub_name)
            elif status == "invalid":
                warnings.append(WarningItem(
                    "bad_score", f"「{sub_name}」的值「{as_text(cell(r, c))}」不是有效分数，已按缺失处理（不按 0 分）",
                    line_no))

        # --- 总分：缺考/为空时，用有效科目分求和补齐，缺考科目不按 0 分 ---
        score, status = clean_score(cell(r, mapping.score)) if mapping.score is not None else (None, "missing")
        if status == "invalid":
            warnings.append(WarningItem(
                "bad_score", f"总分「{as_text(cell(r, mapping.score))}」不是有效分数，已按缺失处理", line_no))
        valid_scores = [v for v in subjects.values() if v is not None]
        if score is None and valid_scores:
            score = float(sum(valid_scores))
            auto_filled += 1
        if score is None:
            warnings.append(WarningItem(
                "absent", f"{name or '未填姓名'}：没有可用总分，排座时按班级平均水平处理（不按 0 分）", line_no))
        if absent_subjects:
            warnings.append(WarningItem(
                "absent", f"{name or '未填姓名'} 缺考/未填：{'、'.join(absent_subjects)}", line_no))

        students.append(Student(
            sid=sid, name=name, score=score,
            gender=normalize_gender(cell(r, mapping.gender)),
            tag=as_text(cell(r, mapping.tag)),
            forbid=split_forbid(cell(r, mapping.forbid)),
            subjects=subjects, source_row=line_no,
        ))

    # --- 整体说明类告警 ---
    if mapping.score is None:
        warnings.insert(0, WarningItem("info", "表格里没有「总分」列，已自动用各科成绩求和得到总分"))
    if auto_filled:
        warnings.insert(0, WarningItem(
            "info", f"共 {auto_filled} 人总分留空或写了缺考，已用有效科目成绩自动求和（缺考科目不按 0 分计）"))
    if subject_names:
        warnings.insert(0, WarningItem("info", f"识别到科目列：{'、'.join(subject_names)}"))
    warnings.insert(0, WarningItem(
        "info", f"表头位于第 {header_row + 1} 行，共读入 {len(students)} 名学生，忽略 {ignored} 行合计/无效行"))

    # --- 重复学生 ---
    for label, getter in (("学号", lambda s: s.sid), ("姓名", lambda s: s.name)):
        seen: Dict[str, int] = {}
        for s in students:
            key = getter(s).strip()
            if not key:
                continue
            seen[key] = seen.get(key, 0) + 1
        dup = [f"{k}({v}次)" for k, v in seen.items() if v > 1]
        if dup:
            warnings.append(WarningItem("duplicate",
                                        f"{label}重复：{'、'.join(dup[:8])}{'…' if len(dup) > 8 else ''}，"
                                        "请确认是否同一人，可在预览表里修改", 0))
    return students, warnings, ignored, subject_names


def import_excel(path: str, sheet_name: Optional[str] = None) -> ImportResult:
    """导入主入口：GUI 只需调用这一个函数。"""
    sheets = list_sheets(path)
    sheet = sheet_name if sheet_name in sheets else sheets[0]
    df = read_matrix(path, sheet)
    header_row, mapping = detect_header(df)
    students, warnings, ignored, subject_names = build_students(df, header_row, mapping)
    if not students:
        raise ExcelParseError("没有读到任何学生数据，请检查表头下方是否填写了学生名单")
    return ImportResult(students=students, warnings=warnings, sheet_name=sheet,
                        path=path, header_row=header_row + 1,
                        subject_names=subject_names, ignored_rows=ignored)


# ----------------------------------------------------------------------
# 三、预览表格 <-> 学生对象（供老师直接改数据）
# ----------------------------------------------------------------------
def students_to_table(students: List[Student]) -> List[List[str]]:
    rows = []
    for s in students:
        rows.append([
            s.sid, s.name, "" if s.score is None else f"{s.score:g}",
            s.gender, s.tag, s.forbid_text(), subjects_to_text(s.subjects),
        ])
    return rows


def students_from_table(rows: List[List[str]]) -> Tuple[List[Student], List[WarningItem]]:
    """预览窗口点「确认导入」时调用：把可能被老师改过的表格转回 Student。"""
    students: List[Student] = []
    warnings: List[WarningItem] = []
    for i, row in enumerate(rows, start=1):
        cells = list(row) + [""] * (len(PREVIEW_COLUMNS) - len(row))
        sid, name = as_text(cells[0]), as_text(cells[1])
        if not any(as_text(c).strip() for c in cells[:6]):
            continue  # 整行被清空 = 删除
        if not name:
            warnings.append(WarningItem("empty_name", f"预览表第 {i} 行姓名为空", i))
        subjects = parse_subjects_text(as_text(cells[6]))
        score, status = clean_score(cells[2])
        if status == "invalid":
            warnings.append(WarningItem("bad_score", f"预览表第 {i} 行总分不是有效数字", i))
        if score is None:
            vals = [v for v in subjects.values() if v is not None]
            score = float(sum(vals)) if vals else None
        students.append(Student(
            sid=sid, name=name, score=score, gender=normalize_gender(cells[3]),
            tag=as_text(cells[4]), forbid=split_forbid(cells[5]),
            subjects=subjects, source_row=i,
        ))
    return students, warnings


# ----------------------------------------------------------------------
# 四、下载标准 Excel 模板
# ----------------------------------------------------------------------
def download_template(path: str) -> str:
    """生成带示例与填写说明的标准模板，返回保存路径。"""
    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "学生名单"
        head_fill = PatternFill("solid", fgColor="4472C4")
        head_font = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
        thin = Side(style="thin", color="B6C2CE")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for c, title in enumerate(TEMPLATE_COLUMNS, start=1):
            cell = ws.cell(row=1, column=c, value=title)
            cell.fill = head_fill
            cell.font = head_font
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
        ws.row_dimensions[1].height = 26

        tips = [
            "建议填学号，留空也能用", "必填，不能为空", "可留空：留空时自动用各科成绩求和",
            "填 男 / 女，可留空", "写了「重点关注/特殊/学困」等词，座位会显示浅红色并优先靠前",
            "填不能坐在一起的学生姓名，多人用顿号分隔，如：张三、李四",
            "可自行增删科目列；缺考请写「缺考」或「—」，不要写 0 分",
            "同上", "同上",
        ]
        for c, tip in enumerate(tips, start=1):
            ws.cell(row=1, column=c).comment = Comment(tip, "排座系统")

        samples = [
            ["2026001", "张小明", 268, "男", "", "李小雨", 88, 92, 88],
            ["2026002", "李小雨", 255, "女", "重点关注", "张小明", 85, 86, 84],
            ["2026003", "王思远", "缺考", "男", "视力需照顾", "", "缺考", 79, 81],
            ["2026004", "赵一诺", 241, "女", "", "", 78, 82, 81],
            ["2026005", "陈子涵", "", "男", "", "", 90, 88, 92],  # 演示总分自动求和
        ]
        for r, row in enumerate(samples, start=2):
            for c, v in enumerate(row, start=1):
                cell = ws.cell(row=r, column=c, value=v)
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="微软雅黑", size=10)

        for c in range(1, len(TEMPLATE_COLUMNS) + 1):
            ws.column_dimensions[get_column_letter(c)].width = 20 if c == 6 else 14
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(TEMPLATE_COLUMNS))}1"

        note = wb.create_sheet("填写说明")
        lines = [
            ["字段", "是否必填", "填写说明"],
            ["学号", "选填", "用于识别重复学生，重复时系统会在导入告警里提示"],
            ["姓名", "必填", "不能为空；空姓名行可在导入预览窗口里直接补写或删除"],
            ["总分", "选填", "留空 / 缺考 / — / N/A 都视为缺失，系统自动用各科成绩求和；绝不按 0 分计算"],
            ["性别", "选填", "男 / 女；开启「性别均衡」开关后参与排座"],
            ["特殊标记", "选填", "含「重点关注、特殊、学困、视力、听力」等词时，方格显示浅红色，并优先安排靠前、靠过道"],
            ["禁止同桌名单", "选填", "填写不能相邻/同桌的学生姓名，多人用 、 , ; 分隔；开启黑名单开关后生效"],
            ["各科成绩", "选填", "列名写科目名即可（语文、数学、英语……可自行增加列），缺考写「缺考」或「—」"],
            ["", "", ""],
            ["容错说明", "", "表头可以不在第一行；支持合并单元格、空行、底部合计行（自动忽略）"],
            ["运行依赖", "", "pandas、openpyxl、xlrd（读 .xls）、matplotlib；tkinter 随官方 Python 自带"],
            ["小提示", "", "导入后一定先看预览窗口的告警，老师可以直接双击单元格修改，再点「确认导入」"],
        ]
        for r, row in enumerate(lines, start=1):
            for c, v in enumerate(row, start=1):
                cell = note.cell(row=r, column=c, value=v)
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                if r == 1:
                    cell.fill = head_fill
                    cell.font = head_font
                else:
                    cell.font = Font(name="微软雅黑", size=10)
        note.column_dimensions["A"].width = 18
        note.column_dimensions["B"].width = 12
        note.column_dimensions["C"].width = 88

        folder = os.path.dirname(os.path.abspath(path))
        if folder:
            os.makedirs(folder, exist_ok=True)
        wb.save(path)
        return path
    except PermissionError:
        raise ExcelParseError("模板文件正被 Excel 打开，请关闭后换个文件名再保存")
    except Exception as exc:
        raise ExcelParseError(f"生成模板失败：{_brief(exc)}")
