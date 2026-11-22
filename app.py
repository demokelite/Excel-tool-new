"""Excel 工具箱入口：页面配置 + 四个 Tab 编排，不写业务逻辑。"""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from core.exporter import XLSX_MIME
from ui.pages import generate, merge, pivot, split

DEMO_FILE = Path(__file__).resolve().parent / "excel-docs" / "sample_sales.xlsx"

st.set_page_config(page_title="Excel 工具箱", layout="centered", page_icon="📊")

st.title("Excel 工具箱")
st.caption(
    "浏览器里处理 Excel：拆分、合并、透视、批量套名，不用装软件、不用写公式。"
    "文件处理完即释放，不保存、不外传、不做二次利用。"
)

if DEMO_FILE.is_file():
    st.download_button(
        "下载演示文件（sample_sales.xlsx）",
        data=DEMO_FILE.read_bytes(),
        file_name=DEMO_FILE.name,
        mime=XLSX_MIME,
        help="1000 行模拟销售数据，下载后可直接上传到下面任一 Tab 试用。",
    )

tab_split, tab_merge, tab_pivot, tab_generate = st.tabs(
    ["按列拆分", "多表合并", "数据透视表", "批量套名"]
)

with tab_split:
    split.render()

with tab_merge:
    merge.render()

with tab_pivot:
    pivot.render()

with tab_generate:
    generate.render()
