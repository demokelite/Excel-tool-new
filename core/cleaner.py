"""数据清洗：全角归一、去空格、空值统一、可选数值化 / 日期解析 / 去重。

只处理 DataFrame，不依赖 streamlit；所有异常都转成中文提示。
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

import pandas as pd

MAX_CELL_LENGTH = 32767
NUMERIC_MIN_RATIO = 0.8
DATE_MIN_RATIO = 0.8
EXCEL_EPOCH = "1899-12-30"
EXCEL_SERIAL_MIN = 20_000
EXCEL_SERIAL_MAX = 60_000
DATE_NAME_HINTS = ("日期", "时间", "date", "time")
BLANK_TEXTS = {"", "-", "—", "无", "N/A", "NA", "NULL", "NONE"}
CURRENCY_CHARS = "¥$€￥,， "
LEADING_ZERO = re.compile(r"0\d+")


@dataclass(frozen=True)
class CleanResult:
    """一次清洗的结果；summary 直接展示给用户。"""

    df: pd.DataFrame
    summary: str
    numeric_columns: tuple[str, ...] = ()
    date_columns: tuple[str, ...] = ()
    blank_cells: int = 0
    date_failed: int = 0
    dropped_duplicates: int = 0
    truncated_cells: int = 0


def clean_table(
    df: pd.DataFrame,
    to_numeric: bool = False,
    to_date: bool = False,
    drop_duplicates: bool = False,
) -> CleanResult:
    """清洗一张表：文本层常开，数值化 / 日期 / 去重按开关执行。"""
    working = df.copy()
    blanks, truncated = _clean_text_columns(working)
    numeric_cols = _convert_numbers(working) if to_numeric else ()
    date_cols, date_failed = _convert_dates(working) if to_date else ((), 0)
    dropped = 0
    if drop_duplicates and len(working):
        before = len(working)
        working = working.drop_duplicates(ignore_index=True)
        dropped = before - len(working)
    summary = _summary(numeric_cols, date_cols, blanks, date_failed, dropped, truncated)
    return CleanResult(
        df=working,
        summary=summary,
        numeric_columns=numeric_cols,
        date_columns=date_cols,
        blank_cells=blanks,
        date_failed=date_failed,
        dropped_duplicates=dropped,
        truncated_cells=truncated,
    )


def _clean_text_columns(df: pd.DataFrame) -> tuple[int, int]:
    """就地清洗文本列，返回（识别到的空值数, 被截断的超长文本数）。"""
    blanks = 0
    truncated = 0
    for index in range(len(df.columns)):
        series = df.iloc[:, index]
        if not pd.api.types.is_object_dtype(series):
            continue
        values: list[object] = []
        for value in series:
            text, was_blank, was_cut = _clean_cell(value)
            blanks += int(was_blank)
            truncated += int(was_cut)
            values.append(text)
        df.isetitem(index, values)
    return blanks, truncated


def _clean_cell(value: object) -> tuple[object, bool, bool]:
    """返回（清洗后的值, 是否空值, 是否被截断）。"""
    if _is_blank(value):
        return pd.NA, True, False
    text = unicodedata.normalize("NFKC", str(value)).replace("\u3000", " ").strip()
    if not text:
        return pd.NA, True, False
    if len(text) > MAX_CELL_LENGTH:
        return text[:MAX_CELL_LENGTH], False, True
    return text, False, False


def _convert_numbers(df: pd.DataFrame) -> tuple[str, ...]:
    """把「¥1,234」「12%」这类文本列转成数值，返回被转换的列名。"""
    converted: list[str] = []
    for index in range(len(df.columns)):
        series = df.iloc[:, index]
        if not _looks_like_numbers(series):
            continue
        parsed = series.map(_parse_number)
        df.isetitem(index, pd.to_numeric(parsed, errors="coerce"))
        converted.append(str(df.columns[index]))
    return tuple(converted)


def _looks_like_numbers(series: pd.Series) -> bool:
    """非空单元格 ≥80% 能解析成数字，且不是带前导零的编号列。"""
    if pd.api.types.is_numeric_dtype(series):
        return False
    values = [value for value in series if not _is_blank(value)]
    if not values:
        return False
    if any(LEADING_ZERO.fullmatch(str(value).strip()) for value in values):
        return False
    parsed = sum(1 for value in values if _parse_number(value) is not None)
    return parsed / len(values) >= NUMERIC_MIN_RATIO


def _parse_number(value: object) -> float | None:
    """去掉货币符号、千分位和空格后转数字；百分号按小数折算，转不了返回 None。"""
    if _is_blank(value):
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    text = unicodedata.normalize("NFKC", str(value)).strip()
    percent = text.endswith("%")
    if percent:
        text = text[:-1]
    for char in CURRENCY_CHARS:
        text = text.replace(char, "")
    text = text.strip()
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return number / 100 if percent else number


def _convert_dates(df: pd.DataFrame) -> tuple[tuple[str, ...], int]:
    """只处理列名带「日期 / 时间 / date / time」的列，成功率 ≥80% 才转。"""
    converted: list[str] = []
    failed = 0
    for index in range(len(df.columns)):
        if not _looks_like_date_column(str(df.columns[index])):
            continue
        series = df.iloc[:, index]
        if pd.api.types.is_datetime64_any_dtype(series):
            continue
        values = [value for value in series if not _is_blank(value)]
        if not values:
            continue
        parsed = _parse_dates(series)
        ok = int(parsed.notna().sum())
        failed += len(values) - ok
        if ok / len(values) < DATE_MIN_RATIO:
            continue
        df.isetitem(index, parsed)
        converted.append(str(df.columns[index]))
    return tuple(converted), failed


def _looks_like_date_column(name: str) -> bool:
    lower = name.lower()
    return any(hint in lower for hint in DATE_NAME_HINTS)


def _parse_dates(series: pd.Series) -> pd.Series:
    """Excel 序列号按 1899-12-30 起算，其余文本交给 pandas 批量解析。"""
    serial = series.map(_parse_serial)
    texts: list[object] = []
    for value in series:
        is_serial = _parse_serial(value) is not None
        texts.append(None if is_serial or _parse_number(value) is not None else _date_text(value))
    parsed = pd.to_datetime(
        pd.Series(texts, index=series.index, dtype=object),
        errors="coerce",
        format="mixed",
    )
    return pd.to_datetime(serial.where(serial.notna(), parsed))


def _date_text(value: object) -> str:
    """把「2024年1月7日」这类中文日期改写成 pandas 认得的格式。"""
    return str(value).strip().replace("年", "-").replace("月", "-").replace("日", "")


def _parse_serial(value: object) -> pd.Timestamp | None:
    """Excel 日期序列号（如 45292）→ Timestamp，不像序列号返回 None。"""
    number = _parse_number(value)
    if number is None or not float(number).is_integer():
        return None
    if not EXCEL_SERIAL_MIN <= number <= EXCEL_SERIAL_MAX:
        return None
    return pd.Timestamp(EXCEL_EPOCH) + pd.Timedelta(int(number), "D")


def _is_blank(value: object) -> bool:
    """判断空值：None / NaN / 空串 / 「-」「无」「N/A」等。"""
    if value is None:
        return True
    try:
        if pd.isna(value):
            return True
    except (TypeError, ValueError):
        return False
    text = unicodedata.normalize("NFKC", str(value)).strip()
    return text.upper() in BLANK_TEXTS


def _summary(
    numeric: tuple[str, ...],
    dates: tuple[str, ...],
    blanks: int,
    date_failed: int,
    dropped: int,
    truncated: int,
) -> str:
    """拼一行「本次清洗：…」，让用户看得见洗了什么。"""
    if not any((numeric, dates, blanks, date_failed, dropped, truncated)):
        return "本次清洗：没有发现需要处理的问题。"
    parts: list[str] = []
    if numeric:
        parts.append(f"数值化 {len(numeric)} 列")
    if dates:
        parts.append(f"日期化 {len(dates)} 列")
    parts.append(f"识别空值 {blanks} 个")
    if date_failed:
        parts.append(f"日期解析失败 {date_failed} 个")
    if dropped:
        parts.append(f"去重 {dropped} 行")
    if truncated:
        parts.append(f"截断超长文本 {truncated} 个")
    return "本次清洗：" + "、".join(parts) + "。"
