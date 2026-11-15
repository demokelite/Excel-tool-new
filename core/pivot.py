"""数据透视表：把明细压成「行维度 × 列维度」的交叉表，带合计行 / 合计列。

只处理 DataFrame，不依赖 streamlit；异常统一转成中文提示。
"""

from __future__ import annotations

from collections.abc import Sequence

import pandas as pd

from core.loader import ExcelToolError, unique_columns

AGG_FUNCS = {
    "求和": "sum",
    "计数": "count",
    "平均": "mean",
    "最大": "max",
    "最小": "min",
}
AGG_OPTIONS = tuple(AGG_FUNCS)
MARGIN_LABEL = "合计"
EMPTY_LABEL = "空值"
MAX_PIVOT_ROWS = 5000
DISPLAY_ROWS = 500
MAX_CHART_ROWS = 100
CHART_BAR = "柱状图"
CHART_LINE = "折线图"
CHART_PIE = "饼图"
CHART_OPTIONS = (CHART_BAR, CHART_LINE, CHART_PIE)


def numeric_columns(df: pd.DataFrame) -> list[str]:
    """返回可以当指标的数值列（布尔列不算）。"""
    names: list[str] = []
    for name, series in df.items():
        text = str(name)
        if text in names or pd.api.types.is_bool_dtype(series):
            continue
        if pd.api.types.is_numeric_dtype(series):
            names.append(text)
    return names


def build_pivot(
    df: pd.DataFrame,
    row_cols: Sequence[str],
    column_col: str | None = None,
    value_col: str = "",
    agg_label: str = "求和",
) -> pd.DataFrame:
    """生成带合计行 / 合计列的交叉表，返回压平表头后的结果，可直接导出。"""
    df = _ensure_unique_columns(df)
    rows = _validate(df, row_cols, column_col, value_col)
    working = _fill_blank_dimensions(df, rows, column_col)
    pivot = pd.pivot_table(
        working,
        index=list(rows),
        columns=[column_col] if column_col else None,
        values=value_col,
        aggfunc=AGG_FUNCS[agg_label],
        fill_value=0,
        margins=True,
        margins_name=MARGIN_LABEL,
        sort=False,
        dropna=False,
    )
    ordered = pivot.reindex(
        index=_sorted_labels(pivot.index),
        columns=_sorted_labels(pivot.columns),
    )
    flat = ordered.reset_index()
    flat.columns = _flatten_names(flat.columns)
    return flat


def chart_data(result: pd.DataFrame, row_cols: Sequence[str]) -> pd.DataFrame:
    """从透视结果取画图数据：去掉合计行 / 合计列，行维度拼成索引，最多 100 行。"""
    rows = [name for name in row_cols if name in result.columns]
    if not rows or result.empty:
        return pd.DataFrame()
    body = result[result[rows[0]] != MARGIN_LABEL]
    labels = body[rows[0]].astype(str)
    for name in rows[1:]:
        labels = labels.str.cat(body[name].astype(str), sep=" / ")
    values = body.drop(columns=[*rows, MARGIN_LABEL], errors="ignore")
    values = values.apply(pd.to_numeric, errors="coerce").fillna(0)
    values.index = labels
    return values.head(MAX_CHART_ROWS)


def pie_data(result: pd.DataFrame, row_cols: Sequence[str]) -> pd.DataFrame:
    """饼图数据：去掉合计行，只保留单一指标，返回「维度 / 数值」两列。"""
    rows = [name for name in row_cols if name in result.columns]
    if not rows or result.empty:
        return pd.DataFrame(columns=["维度", "数值"])
    body = result[result[rows[0]] != MARGIN_LABEL]
    values = body.drop(columns=[*rows, MARGIN_LABEL], errors="ignore")
    if len(values.columns) != 1:
        raise ExcelToolError(
            "饼图只支持单一指标",
            f"当前结果有 {len(values.columns)} 个数值列（列维度把指标展开了），饼图只能画一圈。",
            "把「列维度」设为「（不选）」，或者改用柱状图 / 折线图。",
        )
    labels = body[rows[0]].astype(str)
    for name in rows[1:]:
        labels = labels.str.cat(body[name].astype(str), sep=" / ")
    series = pd.to_numeric(values.iloc[:, 0], errors="coerce").fillna(0)
    data = pd.DataFrame(
        {"维度": labels.to_numpy(), "数值": series.to_numpy()}
    )
    data = data[data["数值"] > 0]
    return data.head(MAX_CHART_ROWS).reset_index(drop=True)


def _ensure_unique_columns(df: pd.DataFrame) -> pd.DataFrame:
    """防御重复列名：上游若传入同名两列，先改名再透视，否则 pandas 会报 not 1-dimensional。"""
    names = [str(name) for name in df.columns]
    if len(set(names)) == len(names):
        return df
    fixed = df.copy()
    fixed.columns = unique_columns(names)
    return fixed


def _validate(
    df: pd.DataFrame,
    row_cols: Sequence[str],
    column_col: str | None,
    value_col: str,
) -> tuple[str, ...]:
    """检查维度列和指标列，返回去重后的行维度。"""
    if df.empty:
        raise ExcelToolError(
            "表格里没有数据行",
            "只有表头的表做不出透视。",
            "确认上传的文件有数据行。",
        )
    rows = tuple(dict.fromkeys(str(name) for name in row_cols if name))
    if not rows:
        raise ExcelToolError(
            "还没选行维度",
            "透视表至少要有 1 个行维度。",
            "在上方「行维度」里选一列。",
        )
    if len(rows) > 2:
        raise ExcelToolError(
            "行维度最多 2 列",
            "维度太多，交叉表会长得没法看。",
            "去掉一列再试。",
        )
    if not value_col:
        raise ExcelToolError(
            "还没选指标列",
            "透视表需要一个数值列来汇总。",
            "在「指标列」里选一列数值。",
        )
    names = [str(name) for name in df.columns]
    missing = [name for name in (*rows, value_col) if name not in names]
    if missing:
        raise ExcelToolError(
            f"找不到这些列：{'、'.join(missing)}",
            "表头可能在读取时被改过，或者换了 sheet。",
            "重新选一次行维度和指标列。",
        )
    if column_col and column_col not in names:
        raise ExcelToolError(
            f"找不到列维度：{column_col}",
            "表头可能在读取时被改过，或者换了 sheet。",
            "重新选一次列维度。",
        )
    if column_col and column_col in rows:
        raise ExcelToolError(
            "列维度和行维度不能是同一列",
            "同一列同时放在行和列上，交叉表没法展开。",
            "换一列当列维度，或者把列维度设为「（不选）」。",
        )
    if value_col in rows or value_col == column_col:
        raise ExcelToolError(
            "指标列不能同时又当维度",
            "同一列既用来分组、又用来汇总，pandas 会直接报「Grouper for 'xx' not 1-dimensional」。",
            "把「指标列」换成另一列数值，比如数量 / 销售额；或者把这一列从行 / 列维度里去掉。",
        )
    series = df[value_col]
    if not pd.api.types.is_numeric_dtype(series):
        raise ExcelToolError(
            f"「{value_col}」不是数值列，不能当指标",
            "这一列里有文字，求和 / 平均都没法算。",
            "换一个数值列当指标，或者在 Excel 里先把这一列转成数字。",
        )
    if series.notna().sum() == 0:
        raise ExcelToolError(
            f"「{value_col}」整列都是空值",
            "没有可以汇总的数据。",
            "换一个指标列，或者先补全这一列再试。",
        )
    return rows


def _fill_blank_dimensions(
    df: pd.DataFrame, rows: Sequence[str], column_col: str | None
) -> pd.DataFrame:
    """把维度列里的空值统一成「空值」，否则合计行 / 合计列会出现 NaN。"""
    working = df.copy()
    targets = [*rows, column_col] if column_col else list(rows)
    for name in targets:
        column = working[name]
        working[name] = column.where(~column.map(_is_blank), EMPTY_LABEL)
    return working


def _is_blank(value: object) -> bool:
    """判断单元格是否为空（NaN 或只有空格 / 全角空格）。"""
    try:
        if value is None or pd.isna(value):
            return True
    except (TypeError, ValueError):
        return False
    return not str(value).strip().replace("\u3000", "")


def _sorted_labels(index: pd.Index) -> pd.Index:
    """按「转成字符串」的顺序排列，合计放最后，空值排最前。"""
    items = list(index)
    head = [item for item in items if MARGIN_LABEL not in _parts(item)]
    tail = [item for item in items if MARGIN_LABEL in _parts(item)]
    ordered = sorted(head, key=lambda item: tuple(_text(part) for part in _parts(item)))
    ordered = ordered + tail
    if isinstance(index, pd.MultiIndex):
        return pd.MultiIndex.from_tuples(ordered, names=index.names)
    return pd.Index(ordered, name=index.name)


def _parts(item: object) -> tuple[object, ...]:
    """把行 / 列标签统一成元组，方便多级索引和单级索引一起处理。"""
    return item if isinstance(item, tuple) else (item,)


def _text(value: object) -> str:
    """排序用的文本：空值统一排在前面。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def _flatten_names(columns: pd.Index) -> list[str]:
    """把可能的多级表头压成一层，重名时追加 _2、_3。"""
    names: list[str] = []
    for item in columns:
        parts = [str(part).strip() for part in _parts(item)]
        text = "_".join(part for part in parts if part not in ("", "nan", "None"))
        names.append(text or EMPTY_LABEL)
    return _unique_names(names)


def _unique_names(names: list[str]) -> list[str]:
    """同名表头追加 _2、_3，避免导出时出现重复列名。"""
    used: set[str] = set()
    result: list[str] = []
    for name in names:
        candidate = name
        index = 2
        while candidate in used:
            candidate = f"{name}_{index}"
            index += 1
        used.add(candidate)
        result.append(candidate)
    return result
