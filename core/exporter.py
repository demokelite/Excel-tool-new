"""导出层：DataFrame → xlsx 字节流、多文件 → zip 字节流，全部在内存完成。"""

from __future__ import annotations

import io
from collections.abc import Mapping
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd

from core.loader import ExcelToolError

XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
ZIP_MIME = "application/zip"
INVALID_SHEET_CHARS = '\\/:*?[]'


def to_xlsx_bytes(df: pd.DataFrame, sheet_name: str = "Sheet1") -> bytes:
    """把 DataFrame 写成 xlsx 字节流（XlsxWriter 引擎，不落盘）。"""
    buffer = io.BytesIO()
    try:
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, sheet_name=_safe_sheet_name(sheet_name), index=False)
    except Exception as exc:
        raise ExcelToolError(
            "导出失败：这个文件写不出来",
            f"单元格内容可能超过 Excel 的 32767 字符上限，或者数据量过大。（{type(exc).__name__}）",
            "先在预览里检查有没有超长文本，或者减少行数后重试。",
        ) from exc
    return buffer.getvalue()


def to_zip_bytes(files: Mapping[str, bytes]) -> bytes:
    """把 {文件名: 字节} 打包成 zip 字节流（不落盘）。"""
    buffer = io.BytesIO()
    try:
        with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
            for name, data in files.items():
                archive.writestr(name, data)
    except Exception as exc:
        raise ExcelToolError(
            "打包失败：zip 生成出错",
            f"文件数量或体积超出了内存限制。（{type(exc).__name__}）",
            "换一列分组，减少文件数量后重试。",
        ) from exc
    return buffer.getvalue()


def _safe_sheet_name(name: str) -> str:
    """Excel 的 sheet 名不能超过 31 字符，也不能含 [ ] : * ? / \\ 这些字符。"""
    cleaned = "".join("_" if char in INVALID_SHEET_CHARS else char for char in name)[:31]
    return cleaned or "Sheet1"
