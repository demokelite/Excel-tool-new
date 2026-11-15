"""Tab3：数据透视表。"""

from __future__ import annotations

import pandas as pd
import streamlit as st

from core import pivot
from core.exporter import XLSX_MIME, to_xlsx_bytes
from core.loader import ExcelToolError, LoadedTable
from ui.components import render_download, render_error, render_single_table_panel

NONE_OPTION = "（不选）"
OUTPUT_NAME = "透视结果.xlsx"
SHEET_NAME = "透视表"
RESULT_KEY = "pivot_result"


def render() -> None:
    """渲染「数据透视表」Tab：上传 → 行列维度 + 指标 → 预览 → 生成并下载。"""
    st.subheader("数据透视表")
    st.caption("按行、列维度把明细压成交叉表，自动带合计行和合计列。")
    table = render_single_table_panel("pivot")
    if table is None:
        return
    df = table.df
    if df.empty:
        st.info("这张表没有数据行，没法做透视。")
        return
    numeric = pivot.numeric_columns(df)
    if not numeric:
        st.info("找不到数值列：可能整表都按文本读取了，回上方关掉「全部按文本读取」再试。")
        return

    st.markdown("**透视配置**")
    columns = list(df.columns)
    row1 = st.selectbox("行维度", columns, key="pivot_row1")
    row2 = st.selectbox("第二行维度（可选）", [NONE_OPTION, *columns], key="pivot_row2")
    row_cols = (row1,) if row2 == NONE_OPTION else (row1, row2)
    column_col = st.selectbox("列维度（可选）", [NONE_OPTION, *columns], key="pivot_col")
    value_col = st.selectbox("指标列（只列数值列）", numeric, key="pivot_value")
    agg_label = st.radio("聚合方式", pivot.AGG_OPTIONS, horizontal=True, key="pivot_agg")

    if len(set(row_cols)) < len(row_cols):
        st.warning("两个行维度选了同一列，换一列再生成。")
        return
    column = None if column_col == NONE_OPTION else column_col
    try:
        result = pivot.build_pivot(df, row_cols, column, value_col, agg_label)
    except ExcelToolError as error:
        render_error(error)
        return

    _render_table(result)
    signature = _signature(table, row_cols, column, value_col, agg_label)
    if st.button("生成透视表", type="primary", key="pivot_run"):
        data = _build(result)
        if data is None:
            return
        st.session_state[RESULT_KEY] = {"signature": signature, "data": data}
        st.rerun()

    saved = st.session_state.get(RESULT_KEY)
    if saved and saved["signature"] == signature:
        st.success(f"已生成 {len(result):,} 行 × {len(result.columns)} 列，可以下载了。")
        render_download(f"下载 {OUTPUT_NAME}", saved["data"], OUTPUT_NAME, XLSX_MIME)


def _render_table(result: pd.DataFrame) -> None:
    """预览交叉表；超过 5000 行只展示前 500 行，导出仍是完整的。"""
    st.caption(f"透视结果：{len(result):,} 行 × {len(result.columns)} 列")
    if len(result) > pivot.MAX_PIVOT_ROWS:
        st.warning(
            f"结果有 {len(result):,} 行，超过 {pivot.MAX_PIVOT_ROWS:,} 行，"
            f"页面只展示前 {pivot.DISPLAY_ROWS} 行，下载文件是完整的。"
        )
        st.dataframe(result.head(pivot.DISPLAY_ROWS))
        return
    st.dataframe(result)


def _build(result: pd.DataFrame) -> bytes | None:
    """生成 xlsx 字节流，失败时渲染中文错误并返回 None。"""
    try:
        with st.spinner("正在生成…"):
            return to_xlsx_bytes(result, SHEET_NAME)
    except ExcelToolError as error:
        render_error(error)
        return None


def _signature(
    table: LoadedTable,
    row_cols: tuple[str, ...],
    column_col: str | None,
    value_col: str,
    agg_label: str,
) -> tuple:
    """配置签名：任何一项变了，上一次的下载按钮就作废。"""
    return (
        table.source_name,
        table.sheet_name,
        len(table.df),
        tuple(map(str, table.df.columns)),
        tuple(row_cols),
        column_col,
        value_col,
        agg_label,
    )
