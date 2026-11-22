"""模板套名：把模板里的 XXX 换成名单里的姓名，生成 N 份副本。

xlsx 本质是 zip + XML。这里直接替换 xl/sharedStrings.xml 与 xl/worksheets/*.xml
里的占位符，模板中的图片、图表、样式原样保留；全程在内存完成，不落盘。
"""

from __future__ import annotations

import io
import re
import unicodedata
import zipfile
from collections.abc import Iterable, Sequence
from dataclasses import dataclass
from datetime import datetime
from xml.sax.saxutils import escape

import pandas as pd

from core.loader import ExcelToolError
from core.splitter import INVALID_FILENAME_CHARS, MAX_FILENAME_LENGTH

PLACEHOLDER = "XXX"
MAX_NAMES = 300
NOTES_FILENAME = "_生成说明.txt"
EXTRA_INVALID_CHARS = "\r\n\t\u3000"
SHARED_STRINGS = "xl/sharedStrings.xml"
SHEET_PART = re.compile(r"^xl/worksheets/sheet\d+\.xml$")
EMBEDDED_PLACEHOLDER = re.compile(r"[A-Za-z0-9]XXX|XXX[A-Za-z0-9]")
CONTENT_TYPES = "[Content_Types].xml"


@dataclass(frozen=True)
class TemplateScan:
    """模板扫描结果：占位符出现次数，以及疑似不是占位符的内嵌用法。"""

    hits: int
    suspicious: tuple[str, ...] = ()


@dataclass(frozen=True)
class NameList:
    """名单整理结果：去空白、去重后的姓名，以及被跳过的数量。"""

    names: tuple[str, ...]
    skipped_blank: int = 0
    skipped_duplicate: int = 0


def scan_template(template_bytes: bytes) -> TemplateScan:
    """统计模板里 XXX 出现的次数，并找出「XXXL」这类可能不是占位符的写法。"""
    parts = _load_parts(template_bytes)
    hits = 0
    suspicious: list[str] = []
    for item, data in parts:
        if not _is_text_part(item.filename):
            continue
        text = data.decode("utf-8", errors="replace")
        hits += text.count(PLACEHOLDER)
        suspicious.extend(match.group(0) for match in EMBEDDED_PLACEHOLDER.finditer(text))
    return TemplateScan(hits=hits, suspicious=tuple(dict.fromkeys(suspicious))[:5])


def prepare_names(values: Iterable[object], limit: int = MAX_NAMES) -> NameList:
    """把名单列整理成姓名：归一全角、去首尾空格、去空值、按首次出现去重。"""
    names: list[str] = []
    seen: set[str] = set()
    blank = 0
    duplicate = 0
    for value in values:
        text = _clean_name(value)
        if not text:
            blank += 1
            continue
        if text in seen:
            duplicate += 1
            continue
        seen.add(text)
        names.append(text)
    if not names:
        raise ExcelToolError(
            "名单里没有可用的姓名",
            "选中列要么全是空值，要么只包含空格。",
            "确认「姓名列」选对了，或者先补全名单再上传。",
        )
    if len(names) > limit:
        raise ExcelToolError(
            f"名单太多：有效姓名 {len(names)} 个，超过 {limit} 个上限",
            "一次生成太多文件，压缩包会很大，下载也容易失败。",
            "先按部门 / 批次拆成几份名单，分几次生成。",
        )
    return NameList(names=tuple(names), skipped_blank=blank, skipped_duplicate=duplicate)


def fill_template(
    template_bytes: bytes,
    names: Sequence[str],
    template_name: str = "",
    source_name: str = "",
) -> dict[str, bytes]:
    """按姓名逐个生成 xlsx，返回 {文件名: 字节}；调用方再打包成 zip。"""
    if not names:
        raise ExcelToolError(
            "还没有可用的姓名",
            "名单为空时生成不出任何文件。",
            "先上传名单并选择「姓名列」。",
        )
    parts = _load_parts(template_bytes)
    if not any(
        _count_placeholder(data)
        for item, data in parts
        if _is_text_part(item.filename)
    ):
        raise ExcelToolError(
            f"模板里没找到占位符 {PLACEHOLDER}",
            f"生成时会查找单元格里的 {PLACEHOLDER}，模板里一个都没有。",
            f"在要显示名字的单元格里输入 {PLACEHOLDER}（三个大写字母），保存为 .xlsx 后重新上传。",
        )
    files: dict[str, bytes] = {}
    used: set[str] = set()
    for name in names:
        filename = _unique_filename(_safe_filename(name), used) + ".xlsx"
        files[filename] = _render(parts, name)
    return files


def build_notes(
    template_name: str,
    source_name: str,
    count: int,
    generated_at: datetime | None = None,
) -> str:
    """生成 zip 里的说明文件：模板、名单来源、人数、时间。"""
    moment = (generated_at or datetime.now()).strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "批量套名说明",
        "============",
        f"模板文件：{template_name or '未命名模板'}",
        f"名单来源：{source_name or '未命名名单'}",
        f"占位符：{PLACEHOLDER}",
        f"生成文件数：{count}",
        f"生成时间：{moment}",
        "",
        "由 Excel 工具箱生成；模板与名单只在内存中处理，不保存、不外传。",
        "",
    ]
    return "\n".join(lines)


def _load_parts(template_bytes: bytes) -> list[tuple[zipfile.ZipInfo, bytes]]:
    """把模板 zip 的所有成员读进内存；不是合法 xlsx 时给中文提示。"""
    buffer = io.BytesIO(template_bytes)
    if not zipfile.is_zipfile(buffer):
        raise ExcelToolError(
            "读取失败：模板不是 .xlsx 文件",
            "xlsx 本质是压缩包，这个文件不是；可能是旧版 .xls 或改了扩展名的其它文件。",
            "用 Excel / WPS 打开后「另存为 .xlsx」再上传。",
        )
    buffer.seek(0)
    try:
        with zipfile.ZipFile(buffer) as archive:
            return [(item, archive.read(item.filename)) for item in archive.infolist()]
    except (zipfile.BadZipFile, OSError) as exc:
        raise ExcelToolError(
            "读取失败：模板压缩结构损坏",
            "文件可能在传输过程中不完整，或者被加密。",
            "重新从原文件另存为 .xlsx 后再上传。",
        ) from exc


def _render(parts: Sequence[tuple[zipfile.ZipInfo, bytes]], name: str) -> bytes:
    """复制模板并替换占位符；非文本部件原样拷贝，保证图片和样式不丢。"""
    escaped = escape(name)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as archive:
        for item, data in parts:
            if _is_text_part(item.filename):
                text = data.decode("utf-8")
                data = text.replace(PLACEHOLDER, escaped).encode("utf-8")
            archive.writestr(item, data)
    return output.getvalue()


def _is_text_part(name: str) -> bool:
    """可能保存单元格文本的部件：共享字符串表和各个 worksheet。"""
    return name == SHARED_STRINGS or bool(SHEET_PART.match(name))


def _count_placeholder(data: bytes) -> int:
    return data.decode("utf-8", errors="replace").count(PLACEHOLDER)


def _clean_name(value: object) -> str:
    """姓名归一：全角转半角、去空格和换行；空值返回空串。"""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        return ""
    text = unicodedata.normalize("NFKC", str(value))
    for char in EXTRA_INVALID_CHARS:
        text = text.replace(char, "")
    return text.strip()


def _safe_filename(label: str) -> str:
    """去掉 Windows 文件名非法字符，截断到 80 字符（与拆分 Tab 同一套规则）。"""
    cleaned = "".join("_" if char in INVALID_FILENAME_CHARS else char for char in label)
    cleaned = cleaned.strip().strip(".")[:MAX_FILENAME_LENGTH]
    return cleaned or "未命名"


def _unique_filename(base: str, used: set[str]) -> str:
    """重名时追加 _2、_3；比较时忽略大小写，截断后仍不超长。"""
    candidate = base
    index = 2
    while candidate.lower() in used:
        suffix = f"_{index}"
        candidate = base[: MAX_FILENAME_LENGTH - len(suffix)] + suffix
        index += 1
    used.add(candidate.lower())
    return candidate
