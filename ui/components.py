"""共用 UI 组件：上传区、sheet / 表头选择、预览、下载、中文错误提示。

本模块只负责展示与收集参数，不做 DataFrame 处理。
"""

from __future__ import annotations

from dataclasses import replace

import pandas as pd
import streamlit as st

from core import cleaner, loader
from core.loader import ExcelToolError, LoadedTable

PREVIEW_ROWS = 20
UPLOAD_TYPES = ["xlsx", "xls", "csv"]
CACHE_ENTRIES = 4


@st.cache_data(show_spinner=False, max_entries=CACHE_ENTRIES)
def _cached_sheets(file_bytes: bytes, filename: str) -> tuple[str, ...]:
    return loader.list_sheets(file_bytes, filename)


@st.cache_data(show_spinner=False, max_entries=CACHE_ENTRIES)
def _cached_table(
    file_bytes: bytes,
    filename: str,
    sheet_name: str | None,
    header_row: int | None,
    all_text: bool,
) -> LoadedTable:
    return loader.load_table(file_bytes, filename, sheet_name, header_row, all_text)


def render_upload_panel(key: str, multiple: bool = False) -> list:
    """渲染上传区，返回上传的文件对象列表。"""
    label = "上传多个 Excel / CSV 文件" if multiple else "上传 Excel / CSV 文件"
    files = st.file_uploader(
        label,
        type=UPLOAD_TYPES,
        accept_multiple_files=multiple,
        key=f"{key}_upload",
        help="单个文件最大 50MB；处理完即释放，不保存、不外传。",
    )
    if not files:
        return []
    return list(files) if multiple else [files]


def render_single_table_panel(key: str) -> LoadedTable | None:
    """上传单个文件 → 选 sheet / 表头行 → 返回解析结果；失败时展示错误并返回 None。"""
    files = render_upload_panel(key, multiple=False)
    if not files:
        return None
    uploaded = files[0]
    file_bytes = uploaded.getvalue()
    filename = uploaded.name

    try:
        sheets = _cached_sheets(file_bytes, filename)
    except ExcelToolError as error:
        render_error(error)
        return None

    sheet: str | None = None
    if len(sheets) > 1:
        sheet = st.selectbox("选择 sheet", sheets, key=f"{key}_sheet")
    elif sheets:
        sheet = sheets[0]

    all_text = st.toggle(
        "全部按文本读取",
        value=False,
        key=f"{key}_text",
        help="用来保留身份证号、工号的前导零；开启后透视表暂时没有数值指标。",
    )

    try:
        probe = _cached_table(file_bytes, filename, sheet, None, all_text)
    except ExcelToolError as error:
        render_error(error)
        return None

    header_no = st.number_input(
        "表头行号（从 1 开始）",
        min_value=1,
        max_value=10_000,
        value=probe.header_row + 1,
        step=1,
        key=f"{key}_header",
        help="默认已自动探测，如果预览的列名不对，改这里。",
    )

    table = probe
    if header_no - 1 != probe.header_row:
        try:
            table = _cached_table(file_bytes, filename, sheet, header_no - 1, all_text)
        except ExcelToolError as error:
            render_error(error)
            return None

    table, summary = _clean_table(table, _render_clean_panel(key))
    for warning in table.warnings:
        st.warning(warning)
    render_preview(table.df)
    st.caption(summary)
    return table


def render_multi_table_panel(key: str) -> list[LoadedTable]:
    """多文件上传 → 逐个选 sheet / 表头行 → 返回所有解析成功的表。"""
    files = render_upload_panel(key, multiple=True)
    if not files:
        return []
    options = _render_clean_panel(key)
    tables: list[LoadedTable] = []
    for index, uploaded in enumerate(files):
        tables.extend(_parse_uploaded(key, index, uploaded, options))
    return tables


def _render_clean_panel(key: str) -> tuple[bool, bool, bool]:
    """清洗开关：默认全关；文本层清洗在 core 里始终执行。"""
    with st.expander("数据清洗（可选）"):
        st.caption(
            "默认已经做了：去前后空格、全角转半角、「-」「无」「N/A」统一成空值。"
            "下面三项按需打开。"
        )
        to_numeric = st.toggle(
            "数值化",
            value=False,
            key=f"{key}_clean_num",
            help="把「¥1,234」「12%」这类文本转成数字，转不了的留空。",
        )
        to_date = st.toggle(
            "日期化",
            value=False,
            key=f"{key}_clean_date",
            help="只处理列名带「日期 / 时间」的列，支持 Excel 序列号。",
        )
        drop_duplicates = st.toggle(
            "去重",
            value=False,
            key=f"{key}_clean_dup",
            help="删除完全相同的行，默认关闭。",
        )
    return to_numeric, to_date, drop_duplicates


def _clean_table(
    table: LoadedTable, options: tuple[bool, bool, bool]
) -> tuple[LoadedTable, str]:
    """按开关清洗一张表，返回（清洗后的表, 一行统计文案）。"""
    result = cleaner.clean_table(table.df, *options)
    return replace(table, df=result.df), result.summary


def _parse_uploaded(
    key: str, index: int, uploaded, options: tuple[bool, bool, bool]
) -> list[LoadedTable]:
    """解析一个上传文件：表头设置整份文件共用，选中的 sheet 各出一张表。"""
    file_bytes = uploaded.getvalue()
    filename = uploaded.name
    with st.expander(filename, expanded=index == 0):
        targets = _sheet_targets(key, index, file_bytes, filename)
        if not targets:
            return []
        all_text = st.toggle(
            "全部按文本读取",
            value=False,
            key=f"{key}_multi_text_{index}",
            help="用来保留工号、订单号的前导零。",
        )
        try:
            probe = _cached_table(file_bytes, filename, targets[0], None, all_text)
        except ExcelToolError as error:
            render_error(error)
            return []
        header_no = st.number_input(
            "表头行号（从 1 开始）",
            min_value=1,
            max_value=10_000,
            value=probe.header_row + 1,
            step=1,
            key=f"{key}_multi_header_{index}",
            help="整份文件共用这一行；某张表的列名不对时，单独上传它再调。",
        )
        tables: list[LoadedTable] = []
        for target in targets:
            try:
                table = _cached_table(
                    file_bytes, filename, target, int(header_no) - 1, all_text
                )
            except ExcelToolError as error:
                render_error(error)
                continue
            table, summary = _clean_table(table, options)
            label = table.sheet_name or "CSV"
            st.caption(f"「{label}」共 {len(table.df):,} 行 × {len(table.df.columns)} 列")
            st.caption(summary)
            if table.warnings:
                st.caption("提示：" + "；".join(table.warnings))
            tables.append(table)
        return tables


def _sheet_targets(
    key: str, index: int, file_bytes: bytes, filename: str
) -> list[str | None]:
    """返回这个文件要合并的 sheet 列表；CSV 没有 sheet，返回 [None]。"""
    try:
        sheets = _cached_sheets(file_bytes, filename)
    except ExcelToolError as error:
        render_error(error)
        return []
    if not sheets:
        return [None]
    chosen = st.multiselect(
        "选择 sheet（可多选）",
        list(sheets),
        default=[sheets[0]],
        key=f"{key}_sheets_{index}",
        help="同一个文件里的多个 sheet 也可以一起合并。",
    )
    if not chosen:
        st.caption("至少选一个 sheet。")
    return list(chosen)


def render_preview(df: pd.DataFrame) -> None:
    """预览前 20 行，并显示总行数 × 列数。"""
    st.caption(f"共 {len(df):,} 行 × {len(df.columns)} 列，下面只预览前 {PREVIEW_ROWS} 行")
    st.dataframe(df.head(PREVIEW_ROWS))


def render_error(error: ExcelToolError) -> None:
    """按「发生了什么 / 可能原因 / 建议」三段式展示错误。"""
    st.error(f"❌ {error.message}")
    if error.reason:
        st.markdown(f"**可能原因**：{error.reason}")
    if error.suggestion:
        st.markdown(f"**建议**：{error.suggestion}")


def render_download(label: str, data: bytes, filename: str, mime: str) -> None:
    """统一的下载按钮，数据始终来自内存。"""
    st.download_button(label, data=data, file_name=filename, mime=mime)
