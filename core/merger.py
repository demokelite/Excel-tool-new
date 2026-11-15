"""多表合并：按列名对齐后纵向拼接，支持并集 / 交集对齐与来源列。

只处理 DataFrame，不依赖 streamlit；所有异常都转成中文提示。
"""

from __future__ import annotations

from collections.abc import Sequence

import pandas as pd

from core.loader import ExcelToolError

ALIGN_UNION = "并集"
ALIGN_INTERSECT = "交集"
ALIGN_MODES = (ALIGN_UNION, ALIGN_INTERSECT)
SOURCE_BY_FILE = "文件名"
SOURCE_BY_SHEET = "sheet 名"
SOURCE_OPTIONS = (SOURCE_BY_FILE, SOURCE_BY_SHEET)
DEFAULT_SOURCE_COLUMN = "来源"


def align_columns(
    frames: Sequence[pd.DataFrame], mode: str = ALIGN_UNION
) -> list[str]:
    """计算合并后的列顺序：以第一张表为准，新列按出现顺序追加到末尾。"""
    if not frames:
        raise ExcelToolError(
            "没有可以合并的表",
            "一个文件或 sheet 都没选上。",
            "先上传至少 2 个文件再合并。",
        )
    columns = list(dict.fromkeys(str(name) for name in frames[0].columns))
    if mode == ALIGN_INTERSECT:
        common = set(columns)
        for frame in frames[1:]:
            common &= {str(name) for name in frame.columns}
        return [name for name in columns if name in common]
    seen = set(columns)
    for frame in frames[1:]:
        for name in frame.columns:
            text = str(name)
            if text not in seen:
                seen.add(text)
                columns.append(text)
    return columns


def merge_tables(
    frames: Sequence[pd.DataFrame],
    labels: Sequence[str] | None = None,
    mode: str = ALIGN_UNION,
    source_column: str | None = None,
) -> pd.DataFrame:
    """按列名对齐合并多张表；source_column 为 None 时不加来源列。"""
    columns = align_columns(frames, mode)
    if not columns:
        raise ExcelToolError(
            "这些表的表头没有交集，可能表头行选错了",
            "选「交集」对齐时，每张表必须至少有一个同名列，现在一个都没有。",
            "改用「并集」对齐；或者回到上传区逐张检查「表头行号」是不是指到了标题行。",
        )
    parts: list[pd.DataFrame] = []
    source_name = _unique_source_column(columns, source_column) if source_column else None
    for index, frame in enumerate(frames):
        part = _align_frame(frame, index).reindex(columns=columns)
        if source_name:
            part.insert(0, source_name, _label_at(labels, index))
        parts.append(part)
    return pd.concat(parts, ignore_index=True, sort=False)


def _unique_source_column(columns: Sequence[str], source_column: str) -> str:
    """来源列和原表列名撞车时，依次尝试「来源_2」「来源_3」……避免合并后重名。"""
    if source_column not in columns:
        return source_column
    index = 2
    while f"{source_column}_{index}" in columns:
        index += 1
    return f"{source_column}_{index}"


def _align_frame(frame: pd.DataFrame, index: int) -> pd.DataFrame:
    """统一列名为字符串，并拦住同一张表里的重复列名。"""
    names = [str(name) for name in frame.columns]
    duplicates = sorted({name for name in names if names.count(name) > 1})
    if duplicates:
        raise ExcelToolError(
            f"第 {index + 1} 张表里有重复的列名：{'、'.join(duplicates)}",
            "两列同名时合并后分不清谁是谁，数据容易对错列。",
            "回到原表把重复的列名改掉，或者在「表头行号」里换一行再试。",
        )
    aligned = frame.copy()
    aligned.columns = names
    return aligned


def _label_at(labels: Sequence[str] | None, index: int) -> str:
    """取第 index 张表的来源标签，取不到就留空。"""
    if not labels or index >= len(labels):
        return ""
    value = labels[index]
    return "" if value is None else str(value)
