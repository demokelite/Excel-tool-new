"""Tab1：按列拆分。"""

from __future__ import annotations

import streamlit as st

from core import splitter
from core.exporter import ZIP_MIME, to_xlsx_bytes, to_zip_bytes
from core.loader import ExcelToolError
from ui.components import render_download, render_error, render_single_table_panel

NONE_OPTION = "（不选）"
SHEET_NAME = "数据"
RESULT_KEY = "split_result"


def render() -> None:
    """渲染「按列拆分」Tab：上传 → 选分组列 → 预估 → 拆分 → 下载 zip。"""
    st.subheader("按列拆分")
    st.caption("按某一列（或两列）把一张表拆成多个 xlsx，打包成一个 zip 下载。")
    table = render_single_table_panel("split")
    if table is None:
        return
    df = table.df
    if df.empty:
        st.info("这张表没有数据行，没法拆分。")
        return

    st.markdown("**拆分配置**")
    columns = list(df.columns)
    first = st.selectbox("分组列", columns, key="split_col1")
    second = st.selectbox("第二分组列（可选）", [NONE_OPTION, *columns], key="split_col2")
    selected = (first,) if second == NONE_OPTION else (first, second)

    if len(selected) == 2 and selected[0] == selected[1]:
        st.warning("两列选了同一个，换一列再拆。")
        return

    try:
        count = splitter.estimate_group_count(df, selected)
    except ExcelToolError as error:
        render_error(error)
        return
    st.caption(f"将拆成 {count} 个文件，合计 {len(df):,} 行")
    if count > splitter.MAX_GROUPS:
        st.warning(f"分组有 {count} 个，超过 {splitter.MAX_GROUPS} 个上限，请换一列。")
        return

    signature = (len(df), tuple(df.columns), selected)
    result = st.session_state.get(RESULT_KEY)
    if result and result["signature"] == signature:
        _render_result(result)

    if st.button("开始拆分", type="primary", key="split_run"):
        payload = _split(df, selected, table.source_name)
        if payload is None:
            return
        st.session_state[RESULT_KEY] = {"signature": signature, **payload}
        st.rerun()


def _split(df, selected: tuple[str, ...], source_name: str) -> dict | None:
    """执行拆分并打包，失败时渲染中文错误并返回 None。"""
    try:
        with st.spinner("正在拆分…"):
            groups = splitter.build_groups(df, selected)
            files = {name: to_xlsx_bytes(frame, SHEET_NAME) for name, frame in groups.items()}
            notes = splitter.build_notes(source_name, selected, len(groups))
            files[splitter.NOTES_FILENAME] = notes.encode("utf-8-sig")
            archive = to_zip_bytes(files)
    except ExcelToolError as error:
        render_error(error)
        return None
    return {
        "zip": archive,
        "count": len(groups),
        "rows": sum(len(frame) for frame in groups.values()),
    }


def _render_result(result: dict) -> None:
    """渲染上一次的拆分结果与下载按钮。"""
    st.success(f"已拆成 {result['count']} 个文件，合计 {result['rows']:,} 行。")
    render_download(
        f"下载 zip（{result['count']} 个文件）",
        result["zip"],
        "拆分结果.zip",
        ZIP_MIME,
    )
