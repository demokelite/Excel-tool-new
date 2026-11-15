"""文件解析：把上传的字节流解析成 DataFrame。

约束：本模块不依赖 streamlit；所有异常都转成中文提示，不向用户抛 traceback。
"""

from __future__ import annotations

import csv
import io
from dataclasses import dataclass

import pandas as pd

MAX_FILE_SIZE = 50 * 1024 * 1024
MAX_ROWS = 200_000
HEADER_SCAN_ROWS = 21
HEADER_SHAPE_ROWS = 3
HEADER_SHAPE_RATIO = 0.8
CSV_ENCODINGS = ("utf-8-sig", "gbk", "gb18030")
CSV_DELIMITERS = ",;\t|"
EXCEL_SUFFIXES = (".xlsx", ".xlsm")


class ExcelToolError(Exception):
    """可直接展示给用户的中文错误：发生了什么 / 可能原因 / 建议怎么做。"""

    def __init__(self, message: str, reason: str = "", suggestion: str = "") -> None:
        super().__init__(message)
        self.message = message
        self.reason = reason
        self.suggestion = suggestion


@dataclass(frozen=True)
class HeaderGuess:
    """表头探测结果：row 是 0 基行号；suspicious 时把 hint 展示给用户。"""

    row: int | None
    suspicious: bool = False
    hint: str = ""


@dataclass
class LoadedTable:
    """一次解析的结果，UI 层据此渲染预览。"""

    df: pd.DataFrame
    sheet_names: tuple[str, ...]
    sheet_name: str | None
    header_row: int
    warnings: tuple[str, ...] = ()
    source_name: str = ""


def file_kind(filename: str) -> str:
    """判断文件类型，返回 "excel" 或 "csv"；不支持的格式抛出中文错误。"""
    lower = filename.lower()
    if lower.endswith(".csv"):
        return "csv"
    if lower.endswith(EXCEL_SUFFIXES):
        return "excel"
    if lower.endswith(".xls"):
        raise ExcelToolError(
            "读取失败：暂时读不了 .xls 老格式",
            "当前环境只安装了 openpyxl，它只能读取 .xlsx。",
            "用 Excel 或 WPS 打开后「另存为 .xlsx」再上传。",
        )
    raise ExcelToolError(
        "读取失败：不支持这种文件类型",
        "工具箱只接受 .xlsx / .xls / .csv。",
        "把文件另存为 .xlsx 后重新上传。",
    )


def list_sheets(file_bytes: bytes, filename: str) -> tuple[str, ...]:
    """返回 Excel 的 sheet 名列表；CSV 没有 sheet，返回空元组。"""
    _check_file(file_bytes, filename)
    if file_kind(filename) == "csv":
        return ()
    try:
        with pd.ExcelFile(io.BytesIO(file_bytes), engine="openpyxl") as book:
            names = tuple(str(name) for name in book.sheet_names)
    except Exception as exc:
        raise _excel_read_error(exc) from exc
    if not names:
        raise ExcelToolError(
            "读取失败：文件里一个 sheet 都没有",
            "文件可能是空文件，或者内容被加密。",
            "用 Excel 打开确认内容后再上传。",
        )
    return names


def detect_header_row(raw: pd.DataFrame) -> int | None:
    """自上而下找第一行「非空单元格 ≥ 2 且无重复值」的行，返回 0 基行号。"""
    candidates = _header_candidates(raw)
    return candidates[0] if candidates else None


def guess_header(raw: pd.DataFrame) -> HeaderGuess:
    """在规则候选里挑最像表头的一行：非空列最多，且下方行的列数对得上。"""
    candidates = _header_candidates(raw)
    if not candidates:
        return HeaderGuess(
            row=None,
            suspicious=True,
            hint="没能自动识别表头，已默认使用第 1 行，可在上方手动调整「表头行号」。",
        )
    passing = [
        row
        for row in candidates
        if _shape_ok(raw, row) and _looks_like_labels(raw, row)
    ]
    if passing:
        best = max(passing, key=lambda row: (len(_non_empty_columns(raw, row)), -row))
        return HeaderGuess(row=best)
    first = candidates[0]
    return HeaderGuess(
        row=first,
        suspicious=True,
        hint=(
            f"疑似表头在第 {first + 1} 行，但它下方几行的列数对不上，"
            "请在「表头行号」里确认。"
        ),
    )


def _header_candidates(raw: pd.DataFrame) -> list[int]:
    """规则候选行：非空单元格 ≥ 2，且这一行没有重复值。"""
    candidates: list[int] = []
    for index in range(len(raw)):
        cells = [_cell_text(value) for value in raw.iloc[index].tolist()]
        non_empty = [text for text in cells if text]
        if len(non_empty) < 2 or len(set(non_empty)) != len(non_empty):
            continue
        candidates.append(index)
    return candidates


def _non_empty_columns(raw: pd.DataFrame, row: int) -> set[int]:
    """某一行里非空单元格所在的列号集合。"""
    return {
        position
        for position, value in enumerate(raw.iloc[row].tolist())
        if _cell_text(value)
    }


def _looks_like_labels(raw: pd.DataFrame, row: int) -> bool:
    """表头的非空单元格大多是文字；数字占多数的行多半是数据行。"""
    texts = [_cell_text(value) for value in raw.iloc[row].tolist()]
    non_empty = [text for text in texts if text]
    if not non_empty:
        return False
    labels = sum(1 for text in non_empty if not _is_number_text(text))
    return labels / len(non_empty) >= HEADER_SHAPE_RATIO


def _is_number_text(text: str) -> bool:
    """去掉货币符号、千分位、百分号和空格后能当数字看。"""
    cleaned = text
    for char in "¥$€￥,，% ":
        cleaned = cleaned.replace(char, "")
    try:
        float(cleaned)
    except ValueError:
        return False
    return True


def _shape_ok(raw: pd.DataFrame, row: int) -> bool:
    """表头下方最多 3 行里，数据落在表头列内的比例 ≥80% 的行占比 ≥80%。"""
    header_cols = _non_empty_columns(raw, row)
    if not header_cols:
        return False
    checked = 0
    passed = 0
    for offset in range(1, HEADER_SHAPE_ROWS + 1):
        below = row + offset
        if below >= len(raw):
            break
        cols = _non_empty_columns(raw, below)
        if not cols:
            continue
        checked += 1
        if len(cols & header_cols) / len(cols) >= HEADER_SHAPE_RATIO:
            passed += 1
    if not checked:
        return True
    return passed / checked >= HEADER_SHAPE_RATIO


def load_table(
    file_bytes: bytes,
    filename: str,
    sheet_name: str | None = None,
    header_row: int | None = None,
    all_text: bool = False,
) -> LoadedTable:
    """把上传的字节解析成 DataFrame。

    header_row 为 0 基行号，传 None 时自动探测；all_text=True 时所有列按文本读取。
    """
    _check_file(file_bytes, filename)
    kind = file_kind(filename)
    warnings: list[str] = []

    if kind == "csv":
        sheets: tuple[str, ...] = ()
        sheet: str | None = None
        encoding: str | None = _sniff_encoding(file_bytes)
    else:
        sheets = list_sheets(file_bytes, filename)
        sheet = sheet_name if sheet_name in sheets else sheets[0]
        encoding = None

    raw = _read_preview(file_bytes, kind, sheet, encoding)
    if raw.empty:
        raise ExcelToolError(
            "读取失败：文件里没有任何内容",
            "这可能是空文件，或者只有格式没有数据。",
            "确认文件内容后再上传。",
        )

    guess = guess_header(raw)
    if header_row is None:
        header_row = guess.row if guess.row is not None else 0
        if guess.hint:
            warnings.append(guess.hint)

    df = _read_full(file_bytes, kind, sheet, header_row, all_text, encoding)
    if len(df) > MAX_ROWS:
        raise ExcelToolError(
            f"读取失败：数据有 {len(df):,} 行，超过 20 万行上限",
            "行数过多容易耗尽服务器内存。",
            "先在 Excel 里删掉不需要的行，或按条件拆成多个文件再上传。",
        )
    warnings.extend(_column_warnings(df))
    return LoadedTable(
        df=df,
        sheet_names=sheets,
        sheet_name=sheet,
        header_row=header_row,
        warnings=tuple(warnings),
        source_name=filename,
    )


def _check_file(file_bytes: bytes, filename: str) -> None:
    if not file_bytes:
        raise ExcelToolError(
            "读取失败：文件是空的",
            "上传的文件没有任何内容。",
            "换一个文件重新上传。",
        )
    if len(file_bytes) > MAX_FILE_SIZE:
        size_mb = len(file_bytes) / 1024 / 1024
        raise ExcelToolError(
            f"读取失败：文件 {size_mb:.1f}MB，超过 50MB 上限",
            "大文件在浏览器端解析容易耗尽内存。",
            "先在 Excel 里删掉无关的 sheet 或列，或者把文件拆成两个再上传。",
        )


def _read_preview(
    file_bytes: bytes, kind: str, sheet_name: str | None, encoding: str | None
) -> pd.DataFrame:
    """只读前若干行，用于探测表头，避免整表读两遍。"""
    buffer = io.BytesIO(file_bytes)
    try:
        if kind == "csv":
            sep = _sniff_delimiter(file_bytes, encoding or "utf-8-sig")
            return pd.read_csv(
                buffer,
                header=None,
                nrows=HEADER_SCAN_ROWS,
                encoding=encoding,
                sep=sep,
                dtype=str,
                skip_blank_lines=False,
            )
        return pd.read_excel(
            buffer,
            sheet_name=sheet_name,
            header=None,
            nrows=HEADER_SCAN_ROWS,
            engine="openpyxl",
        )
    except Exception as exc:
        error = _csv_read_error(exc) if kind == "csv" else _excel_read_error(exc)
        raise error from exc


def _read_full(
    file_bytes: bytes,
    kind: str,
    sheet_name: str | None,
    header_row: int,
    all_text: bool,
    encoding: str | None,
) -> pd.DataFrame:
    """按确定的表头行读取整张表。"""
    buffer = io.BytesIO(file_bytes)
    dtype = str if all_text else None
    try:
        if kind == "csv":
            sep = _sniff_delimiter(file_bytes, encoding or "utf-8-sig")
            df = pd.read_csv(
                buffer,
                header=header_row,
                dtype=dtype,
                encoding=encoding,
                sep=sep,
            )
        else:
            df = pd.read_excel(
                buffer,
                sheet_name=sheet_name,
                header=header_row,
                dtype=dtype,
                engine="openpyxl",
            )
    except Exception as exc:
        error = _csv_read_error(exc) if kind == "csv" else _excel_read_error(exc)
        raise error from exc
    df.columns = [str(name).strip().replace("\u3000", "") for name in df.columns]
    return df


def _sniff_encoding(file_bytes: bytes) -> str:
    """按 utf-8-sig / gbk / gb18030 顺序探测 CSV 编码。"""
    sample = file_bytes[:65536]
    for encoding in CSV_ENCODINGS:
        try:
            sample.decode(encoding)
            return encoding
        except UnicodeDecodeError:
            continue
    raise ExcelToolError(
        "读取失败：认不出这个 CSV 的编码",
        "文件不是 utf-8 / gbk / gb18030 中的任何一种。",
        "用记事本打开 CSV，「另存为」时把编码选成 UTF-8 再上传。",
    )


def _sniff_delimiter(file_bytes: bytes, encoding: str) -> str:
    sample = file_bytes[:65536].decode(encoding, errors="replace")
    try:
        return csv.Sniffer().sniff(sample, delimiters=CSV_DELIMITERS).delimiter
    except csv.Error:
        return ","


def _cell_text(value: object) -> str:
    """把单元格转成用于比较的文本，空值统一返回空字符串。"""
    try:
        if value is None or pd.isna(value):
            return ""
    except (TypeError, ValueError):
        return ""
    return str(value).strip().replace("\u3000", "")


def _column_warnings(df: pd.DataFrame) -> list[str]:
    warnings: list[str] = []
    unnamed = [name for name in df.columns if str(name).startswith("Unnamed:")]
    if unnamed:
        warnings.append(f"检测到 {len(unnamed)} 个没有列名的列，表头行可能选错了。")
    if df.empty:
        warnings.append("文件只有表头，没有数据行。")
    return warnings


def _excel_read_error(exc: Exception) -> ExcelToolError:
    return ExcelToolError(
        "读取失败：这个文件可能不是标准 xlsx",
        f"文件被加密、已损坏，或者扩展名是 .xlsx 实际是 .xls；也可能「表头行号」超出了文件范围。（{type(exc).__name__}）",
        "用 Excel 打开后「另存为 .xlsx」再上传；如果加了打开密码，先取消密码。",
    )


def _csv_read_error(exc: Exception) -> ExcelToolError:
    return ExcelToolError(
        "读取失败：这个 CSV 解析不了",
        f"编码或分隔符不是常见格式，或者「表头行号」超出了文件范围。（{type(exc).__name__}）",
        "先把表头行号改回 1 再试；仍然失败就用 Excel 另存为 .xlsx 后上传。",
    )
