"""Excel 工具箱入口：页面配置 + 三个 Tab 编排，不写业务逻辑。"""

from __future__ import annotations

import streamlit as st

from ui.pages import merge, pivot, split

st.set_page_config(page_title="Excel 工具箱", layout="centered", page_icon="📊")

st.title("Excel 工具箱")
st.caption(
    "浏览器里处理 Excel：拆分、合并、透视，不用装软件、不用写公式。"
    "文件处理完即释放，不保存、不外传、不做二次利用。"
)

tab_split, tab_merge, tab_pivot = st.tabs(["按列拆分", "多表合并", "数据透视表"])

with tab_split:
    split.render()

with tab_merge:
    merge.render()

with tab_pivot:
    pivot.render()
