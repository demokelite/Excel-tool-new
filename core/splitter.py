"""按列拆分：把一张表按 1–2 列的唯一值切成多份，并给出安全的文件名。

本模块只处理 DataFrame，不依赖 streamlit；文件名规则集中在 _safe_filename。
"""

from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime

import pandas as pd

from core.loader import ExcelToolError

MAX_GROUPS = 300
MAX_FILENAME_LENGTH = 80
EMPTY_LABEL = "空值"
NOTES_FILENAME = "_拆分说明.txt"
INVALID_FILENAME_CHARS = '\\/:*?"<>|'
EXTRA_INVALID_CHARS = "\r\n\t\u3000"


def validate_columns(df: pd.DataFrame, columns: Sequence[str]) -> tuple[str, ...]:
    """检查分组列：必须有 1–2 列且都在表里，返回去重后的列名元组。"""
    selected = tuple(dict.fromkeys(name for name in columns if name))
    if not selected:
        raise ExcelToolError(
            "还没选分组列",
            "拆分至少需要 1 列作为依据。",
            "在上方「分组列」里选一列。",
        )
    if len(selected) > 2:
        raise ExcelToolError(
            "最多只能选 2 列分组",
            "列太多会让分组数量爆炸。",
            "先去掉一列再拆分。",
        )
    missing = [name for name in selected if name not in df.columns]
    if missing:
        raise ExcelToolError(
            f"找不到分组列：{'、'.join(missing)}",
            "表头可能在读取时被改过，或者换了 sheet。",
            "重新选一次分组列。",
        )
    _check_empty_columns(df, selected)
    return selected


def build_notes(
    source_name: str,
    columns: Sequence[str],
    file_count: int,
    generated_at: datetime | None = None,
) -> str:
    """生成 zip 里的拆分说明：源文件、分组列、文件数、生成时间。"""
    moment = (generated_at or datetime.now()).strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "拆分说明",
        "========",
        f"源文件：{source_name or '未命名文件'}",
        f"分组列：{'、'.join(columns)}",
        f"文件数：{file_count}",
        f"生成时间：{moment}",
        "",
        "由 Excel 工具箱生成；原始文件只在内存中处理，不保存、不外传。",
        "",
    ]
    return "\n".join(lines)


def _check_empty_columns(df: pd.DataFrame, selected: tuple[str, ...]) -> None:
    """整列都是空值时直接提示，避免拆出一个没有意义的「空值.xlsx」。"""
    if df.empty:
        return
    for name in selected:
        column = df[name]
        if column.isna().all() or all(_is_blank(value) for value in column):
            raise ExcelToolError(
                f"「{name}」整列都是空值，没法按它拆分",
                "所有行都会落进同一个「空值」分组，拆出来没有意义。",
                "换一列分组，或者先把这一列补全再拆。",
            )


def _is_blank(value: object) -> bool:
    """判断单元格是否为空（NaN 或只有空格/全角空格）。"""
    try:
        if value is None or pd.isna(value):
            return True
    except (TypeError, ValueError):
        return False
    return not str(value).strip().replace("\u3000", "")


def estimate_group_count(df: pd.DataFrame, columns: Sequence[str]) -> int:
    """预估会拆成多少个文件，口径与 build_groups 完全一致。"""
    selected = validate_columns(df, columns)
    if df.empty:
        return 0
    return int(df.groupby(_group_keys(df, selected)).ngroups)


def build_groups(df: pd.DataFrame, columns: Sequence[str]) -> dict[str, pd.DataFrame]:
    """按列分组，返回 {文件名: 子表}；分组数超过 300 时抛 ExcelToolError。"""
    selected = validate_columns(df, columns)
    if df.empty:
        raise ExcelToolError(
            "表格里没有数据行",
            "只有表头的表拆不出内容。",
            "确认上传的文件有数据行。",
        )
    # 用循环收集分组结果：pandas 在 dropna=False + 空值时直接 list() 会触发 .groups 报错
    grouped: list[tuple[object, pd.DataFrame]] = []
    for key, frame in df.groupby(_group_keys(df, selected), sort=False):
        grouped.append((key, frame))
    if len(grouped) > MAX_GROUPS:
        raise ExcelToolError(
            f"分组过多：这一拆会得到 {len(grouped)} 个文件",
            f"超过 {MAX_GROUPS} 个文件后，下载和逐个打开都会很麻烦。",
            "换一列分组，或者先筛选掉一部分数据再拆。",
        )
    groups: dict[str, pd.DataFrame] = {}
    used: set[str] = set()
    for key, frame in sorted(grouped, key=lambda item: _join_labels(item[0])):
        label = _join_labels(key)
        filename = _unique_filename(_safe_filename(label), used) + ".xlsx"
        groups[filename] = frame.reset_index(drop=True)
    return groups


def _group_keys(df: pd.DataFrame, selected: tuple[str, ...]) -> list[pd.Series]:
    """把空值（含纯空格）统一成「空值」，保证它们只落进一个分组。"""
    return [
        df[name].where(~df[name].map(_is_blank), EMPTY_LABEL)
        for name in selected
    ]


def _join_labels(key: object) -> str:
    """把分组键拼成一个展示用的名字，两列用下划线连接。"""
    keys = key if isinstance(key, tuple) else (key,)
    return "_".join(_label(value) for value in keys)


def _label(value: object) -> str:
    """单元格值 → 文件名用的文本；空值统一叫「空值」。"""
    try:
        if value is None or pd.isna(value):
            return EMPTY_LABEL
    except (TypeError, ValueError):
        return EMPTY_LABEL
    if isinstance(value, pd.Timestamp):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    text = str(value).strip().replace("\u3000", "")
    return text or EMPTY_LABEL


def _safe_filename(label: str) -> str:
    """去掉 Windows 文件名非法字符，截断到 80 字符。"""
    cleaned = "".join(
        "_" if char in INVALID_FILENAME_CHARS or char in EXTRA_INVALID_CHARS else char
        for char in label
    )
    cleaned = cleaned.strip().strip(".")[:MAX_FILENAME_LENGTH]
    return cleaned or EMPTY_LABEL


def _unique_filename(base: str, used: set[str]) -> str:
    """同名时追加 _2、_3；比较时忽略大小写，截断后仍保证不超长。"""
    candidate = base
    index = 2
    while candidate.lower() in used:
        suffix = f"_{index}"
        candidate = base[: MAX_FILENAME_LENGTH - len(suffix)] + suffix
        index += 1
    used.add(candidate.lower())
    return candidate
