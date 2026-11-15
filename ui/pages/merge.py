"""Tab2：多表合并。"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core import merger
from core.exporter import XLSX_MIME, to_xlsx_bytes
from core.loader import ExcelToolError, LoadedTable
from ui.components import (
    render_download,
    render_error,
    render_multi_table_panel,
    render_preview,
)

OUTPUT_NAME = "合并结果.xlsx"
SHEET_NAME = "合并结果"
RESULT_KEY = "merge_result"
SOURCE_COLUMN = merger.DEFAULT_SOURCE_COLUMN


def render() -> None:
    """渲染「多表合并」Tab：上传 → 对齐方式 → 来源列 → 预览 → 合并并下载。"""
    st.subheader("多表合并")
    st.caption("把多个文件（或同一个文件的多个 sheet）按列名纵向拼成一张表。")
    tables = render_multi_table_panel("merge")
    if len(tables) < 2:
        st.info("至少要有 2 张表才能合并：多传一个文件，或者在同一个文件里多选一个 sheet。")
        st.button("合并并下载", disabled=True, key="merge_run")
        return

    mode = st.radio(
        "表头对齐方式",
        merger.ALIGN_MODES,
        horizontal=True,
        key="merge_mode",
        help="并集：保留所有列，某张表没有的列留空；交集：只保留每张表都有的列。",
    )
    add_source = st.toggle(
        "添加来源列",
        value=False,
        key="merge_source_on",
        help="在结果第一列标出每一行来自哪个文件或 sheet。",
    )
    source_by = merger.SOURCE_BY_FILE
    if add_source:
        source_by = st.radio(
            "来源列的值",
            merger.SOURCE_OPTIONS,
            horizontal=True,
            key="merge_source_by",
        )
    labels = [_source_label(table, source_by) for table in tables]

    try:
        merged = merger.merge_tables(
            [table.df for table in tables],
            labels=labels,
            mode=mode,
            source_column=SOURCE_COLUMN if add_source else None,
        )
    except ExcelToolError as error:
        render_error(error)
        return

    render_preview(merged)
    if add_source and merged.columns[0] != SOURCE_COLUMN:
        st.caption(
            f"原表里已经有「{SOURCE_COLUMN}」列，新增的来源列改名为「{merged.columns[0]}」。"
        )
    signature = (_tables_signature(tables), mode, source_by if add_source else None)
    if st.button("合并并下载", type="primary", key="merge_run"):
        data = _build(merged)
        if data is None:
            return
        st.session_state[RESULT_KEY] = {"signature": signature, "data": data}
        st.rerun()

    result = st.session_state.get(RESULT_KEY)
    if result and result["signature"] == signature:
        st.success(f"已合并 {len(merged):,} 行 × {len(merged.columns)} 列，可以下载了。")
        render_download(f"下载 {OUTPUT_NAME}", result["data"], OUTPUT_NAME, XLSX_MIME)


def _build(merged: pd.DataFrame) -> bytes | None:
    """生成 xlsx 字节流，失败时渲染中文错误并返回 None。"""
    try:
        with st.spinner("正在合并…"):
            return to_xlsx_bytes(merged, SHEET_NAME)
    except ExcelToolError as error:
        render_error(error)
        return None


def _source_label(table: LoadedTable, source_by: str) -> str:
    """来源列的取值：文件名，或 sheet 名（CSV 没有 sheet 时退回文件名）。"""
    if source_by == merger.SOURCE_BY_SHEET and table.sheet_name:
        return table.sheet_name
    return table.source_name


def _tables_signature(tables: list[LoadedTable]) -> tuple:
    """把每张表的身份压成元组，用于判断上一次的下载结果是否过期。"""
    return tuple(
        (table.source_name, table.sheet_name, len(table.df), tuple(map(str, table.df.columns)))
        for table in tables
    )
