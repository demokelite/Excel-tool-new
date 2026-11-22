"""Tab4：批量套名。模板里的 XXX 逐个替换成名单里的姓名，一次生成多份。"""

from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from core import generator
from core.exporter import ZIP_MIME, to_zip_bytes
from core.loader import ExcelToolError
from ui.components import (
    render_download,
    render_error,
    render_single_table_panel,
)

RESULT_KEY = "generate_result"
TEMPLATE_TYPES = ["xlsx"]
NAME_HINTS = ("姓名", "名字", "name")


def render() -> None:
    """渲染「批量套名」Tab：模板 + 名单 → 选姓名列 → 生成 → 下载 zip。"""
    st.subheader("批量套名")
    st.caption(
        f"上传一份模板和一张名单，把模板里的 {generator.PLACEHOLDER} 换成每个人的姓名，"
        "一次生成多份文件。模板的图片、样式、合并单元格都保持原样。"
    )

    with st.expander("怎么准备模板和名单？"):
        st.markdown(
            f"""
1. 模板用 Excel 打开，在**要显示名字的单元格**里输入 `{generator.PLACEHOLDER}`（三个大写字母）；
2. 模板里所有 `{generator.PLACEHOLDER}` 都会被替换，别在其它地方留这三个字母；
3. 名单可以是 .xlsx / .csv，只要有一列是姓名，程序会按这一列逐个生成；
4. 模板请保存为 `.xlsx`，旧版 `.xls` 需先另存。
            """
        )

    template_col, names_col = st.columns(2)
    with template_col:
        st.markdown("**① 模板文件**")
        template = st.file_uploader(
            "上传模板（.xlsx）",
            type=TEMPLATE_TYPES,
            key="generate_template",
            help=f"名字位置写成 {generator.PLACEHOLDER}；单个文件最大 50MB。",
        )
        hits = _render_template_hint(template)

    with names_col:
        st.markdown("**② 姓名名单**")
        table = render_single_table_panel("generate_names")

    if template is None or hits is None or table is None or table.df.empty:
        return

    template_bytes = template.getvalue()
    columns = list(table.df.columns)
    if not columns:
        st.info("名单里没有可选的列。")
        return

    st.markdown("**③ 生成设置**")
    name_col = st.selectbox(
        "姓名列",
        columns,
        index=_default_column(columns),
        format_func=str,
        key="generate_name_col",
    )
    try:
        prepared = generator.prepare_names(table.df[name_col].tolist())
    except ExcelToolError as error:
        render_error(error)
        return

    st.caption(
        f"有效姓名 {len(prepared.names)} 人"
        f"（跳过空值 {prepared.skipped_blank} 个、重复 {prepared.skipped_duplicate} 个），"
        f"将生成 {len(prepared.names)} 个文件，文件名用姓名。"
    )
    if prepared.skipped_duplicate:
        st.warning("名单里有重名，已按第一次出现保留；如需要区分，请先在名单里改成不重名。")
    if hits == 0:
        st.warning(
            f"模板里没找到 {generator.PLACEHOLDER}，生成按钮已禁用。"
            f"请在要显示名字的单元格里输入 {generator.PLACEHOLDER} 后重新上传。"
        )

    signature = _signature(template.name, len(template_bytes), table, name_col, prepared.names)
    if st.button(
        f"开始生成（{len(prepared.names)} 个文件）",
        type="primary",
        key="generate_run",
        disabled=hits == 0,
    ):
        data = _build(template_bytes, template.name, prepared.names, table)
        if data is None:
            return
        st.session_state[RESULT_KEY] = {"signature": signature, "data": data}
        st.rerun()

    saved = st.session_state.get(RESULT_KEY)
    if saved and saved["signature"] == signature:
        st.success(f"已生成 {len(prepared.names)} 个文件，可以下载了。")
        render_download(
            "下载生成结果（zip）",
            saved["data"],
            _zip_name(template.name, len(prepared.names)),
            ZIP_MIME,
        )


def _render_template_hint(template) -> int | None:
    """上传模板后立即扫描占位符：返回出现次数；解析失败返回 None。"""
    if template is None:
        return None
    try:
        scan = generator.scan_template(template.getvalue())
    except ExcelToolError as error:
        render_error(error)
        return None
    if scan.hits:
        st.success(f"检测到 {scan.hits} 处 {generator.PLACEHOLDER} 占位符，可以生成。")
        if scan.suspicious:
            st.warning(
                f"模板里还有和字母 / 数字连着的写法（{'、'.join(scan.suspicious)}），"
                "它们可能不是占位符，但也会被一起替换成姓名，请确认。"
            )
    else:
        st.error(
            f"模板里没有 {generator.PLACEHOLDER}。"
            f"请在要显示名字的单元格里输入 {generator.PLACEHOLDER} 后重新保存、上传。"
        )
    return scan.hits


def _default_column(columns: list) -> int:
    """默认选中列名里带「姓名 / 名字 / name」的列，找不到就用第一列。"""
    for index, name in enumerate(columns):
        lower = str(name).lower()
        if any(hint in lower for hint in NAME_HINTS):
            return index
    return 0


def _build(
    template_bytes: bytes,
    template_name: str,
    names: tuple[str, ...],
    table,
) -> bytes | None:
    """生成各份 xlsx 并打包成 zip；出错时渲染中文提示并返回 None。"""
    try:
        with st.spinner("正在生成…"):
            files = generator.fill_template(
                template_bytes,
                names,
                template_name=template_name,
                source_name=table.source_name,
            )
            notes = generator.build_notes(template_name, table.source_name, len(names))
            files[generator.NOTES_FILENAME] = notes.encode("utf-8-sig")
            return to_zip_bytes(files)
    except ExcelToolError as error:
        render_error(error)
        return None


def _zip_name(template_name: str, count: int) -> str:
    """下载文件名：模板名 + 份数。"""
    stem = Path(template_name).stem or "批量套名"
    return f"{stem}_套名{count}份.zip"


def _signature(
    template_name: str,
    template_size: int,
    table,
    name_col: str,
    names: tuple[str, ...],
) -> tuple:
    """配置签名：模板、名单或姓名列一变，上一轮的下载按钮就作废。"""
    return (
        template_name,
        template_size,
        table.source_name,
        table.sheet_name,
        len(table.df),
        tuple(map(str, table.df.columns)),
        name_col,
        names,
    )
