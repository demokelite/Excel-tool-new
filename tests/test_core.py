"""core 层单元测试，只用标准库 unittest，不新增依赖。

运行：python tests/test_core.py
"""

from __future__ import annotations

import io
import sys
import unittest
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from core import cleaner, loader, merger, pivot, splitter  # noqa: E402
from core.exporter import to_xlsx_bytes, to_zip_bytes  # noqa: E402


def _write_xlsx(sheets: dict[str, pd.DataFrame], header: bool = True) -> bytes:
    """在内存里造一个多 sheet 的 xlsx。"""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False, header=header)
    return buffer.getvalue()


def _title_then_header() -> bytes:
    """前两行是标题和日期，第 3 行才是表头。"""
    rows = [
        ["2024 年销售明细", None, None],
        ["导出时间：2024-01-01", None, None],
        ["订单号", "省份", "数量"],
        ["00123", "广东", 3],
        ["00124", "浙江", 5],
    ]
    return _write_xlsx({"销售": pd.DataFrame(rows)}, header=False)


def _sales_df() -> pd.DataFrame:
    provinces = [f"省份{i:02d}" for i in range(30)] + [None]
    return pd.DataFrame(
        {
            "省份": [provinces[i % 31] for i in range(1000)],
            "城市": ["广东/深圳" if i % 2 else "广州" for i in range(1000)],
            "销售额": range(1000),
        }
    )


class TestLoader(unittest.TestCase):
    def test_detect_header_row_skips_title_and_duplicates(self) -> None:
        raw = pd.DataFrame([["报表"], ["a", "a"], ["订单号", "省份"], ["1", "广东"]])
        self.assertEqual(loader.detect_header_row(raw), 2)

    def test_detect_header_row_returns_none(self) -> None:
        raw = pd.DataFrame([["只有一列"], [None]])
        self.assertIsNone(loader.detect_header_row(raw))

    def test_load_table_detects_header_and_source(self) -> None:
        table = loader.load_table(_title_then_header(), "销售.xlsx")
        self.assertEqual(table.header_row, 2)
        self.assertEqual(table.df.shape, (2, 3))
        self.assertEqual(list(table.df.columns), ["订单号", "省份", "数量"])
        self.assertEqual(table.source_name, "销售.xlsx")

    def test_guess_header_skips_two_cell_info_row(self) -> None:
        raw = pd.DataFrame(
            [
                ["2024 年销售明细", None, None],
                ["导出时间", "2024-01-01", None],
                ["订单号", "省份", "数量"],
                ["00123", "广东", 3],
                ["00124", "浙江", 5],
            ]
        )
        guess = loader.guess_header(raw)
        self.assertEqual(guess.row, 2)
        self.assertFalse(guess.suspicious)

    def test_guess_header_flags_shape_mismatch(self) -> None:
        raw = pd.DataFrame(
            [
                ["订单号", "省份", None],
                ["00123", "广东", 3],
                ["00124", "浙江", 5],
                ["00125", "江苏", 8],
            ]
        )
        guess = loader.guess_header(raw)
        self.assertEqual(guess.row, 0)
        self.assertTrue(guess.suspicious)
        self.assertIn("疑似表头在第 1 行", guess.hint)

    def test_guess_header_returns_none_without_candidate(self) -> None:
        guess = loader.guess_header(pd.DataFrame([["只有一列"], [None]]))
        self.assertIsNone(guess.row)
        self.assertTrue(guess.suspicious)

    def test_load_table_warns_on_suspicious_header(self) -> None:
        rows = [
            ["订单号", "省份", None],
            ["00123", "广东", 3],
            ["00124", "浙江", 5],
        ]
        data = _write_xlsx({"t": pd.DataFrame(rows)}, header=False)
        table = loader.load_table(data, "t.xlsx")
        self.assertEqual(table.header_row, 0)
        self.assertTrue(any("疑似表头" in text for text in table.warnings))

    def test_xls_extension_with_zip_content_reads_as_xlsx(self) -> None:
        data = _write_xlsx({"t": pd.DataFrame({"a": [1], "b": [2]})})
        table = loader.load_table(data, "renamed.xls")
        self.assertEqual(list(table.df.columns), ["a", "b"])

    def test_xls_extension_with_csv_content_reads_as_csv(self) -> None:
        data = "省份,数量\n广东,3\n".encode("utf-8-sig")
        table = loader.load_table(data, "export.xls")
        self.assertEqual(table.df["省份"].tolist(), ["广东"])

    def test_real_xls_reports_missing_engine(self) -> None:
        data = loader.OLE2_MAGIC + b"\x00" * 64
        with self.assertRaises(loader.ExcelToolError) as ctx:
            loader.load_table(data, "old.xls")
        self.assertIn("老格式", ctx.exception.message)

    def test_html_disguised_xls_is_reported(self) -> None:
        data = "<html><body><table><tr><td>1</td></tr></table></body></html>"
        with self.assertRaises(loader.ExcelToolError) as ctx:
            loader.load_table(data.encode("gbk"), "fake.xls")
        self.assertIn("网页格式", ctx.exception.message)

    def test_headers_differing_only_by_space_are_renamed(self) -> None:
        raw = pd.DataFrame(
            [["订单号", "订单号 ", "金额"], ["001", "002", 3], ["003", "004", 5]]
        )
        data = _write_xlsx({"t": raw}, header=False)
        table = loader.load_table(data, "dup.xlsx")
        self.assertEqual(table.header_row, 0)
        self.assertEqual(list(table.df.columns), ["订单号", "订单号_2", "金额"])
        self.assertTrue(any("重复列名" in text for text in table.warnings))

    def test_pivot_works_after_duplicate_headers_renamed(self) -> None:
        raw = pd.DataFrame(
            [["订单号", "订单号 ", "金额"], ["001", "002", 3], ["003", "004", 5]]
        )
        data = _write_xlsx({"t": raw}, header=False)
        table = loader.load_table(data, "dup.xlsx")
        result = pivot.build_pivot(table.df, ["订单号"], None, "金额", "求和")
        self.assertFalse(result.empty)

    def test_load_table_all_text(self) -> None:
        table = loader.load_table(_title_then_header(), "销售.xlsx", all_text=True)
        self.assertEqual(table.df["数量"].iloc[0], "3")

    def test_load_table_multi_sheet(self) -> None:
        sheets = {"销售": pd.DataFrame({"a": [1]}), "第二页": pd.DataFrame({"b": [2, 3]})}
        data = _write_xlsx(sheets)
        self.assertEqual(loader.list_sheets(data, "t.xlsx"), ("销售", "第二页"))
        table = loader.load_table(data, "t.xlsx", sheet_name="第二页")
        self.assertEqual(table.df.shape, (2, 1))

    def test_corrupt_file_reports_chinese_error(self) -> None:
        with self.assertRaises(loader.ExcelToolError) as ctx:
            loader.load_table(b"not an excel at all", "bad.xlsx")
        self.assertIn("读取失败", ctx.exception.message)

    def test_legacy_xls_text_reads_as_csv(self) -> None:
        table = loader.load_table(b"abc\n1\n", "old.xls")
        self.assertEqual(table.df.shape, (1, 1))

    def test_empty_file_reports_error(self) -> None:
        with self.assertRaises(loader.ExcelToolError):
            loader.load_table(b"", "empty.xlsx")


class TestSplitter(unittest.TestCase):
    def test_31_groups_keep_all_rows(self) -> None:
        df = _sales_df()
        groups = splitter.build_groups(df, ["省份"])
        self.assertEqual(len(groups), 31)
        self.assertEqual(sum(len(frame) for frame in groups.values()), len(df))
        self.assertIn("空值.xlsx", groups)

    def test_slash_in_value_becomes_underscore(self) -> None:
        groups = splitter.build_groups(_sales_df(), ["城市"])
        self.assertIn("广东_深圳.xlsx", groups)
        self.assertIn("广州.xlsx", groups)

    def test_same_name_gets_numeric_suffix(self) -> None:
        df = pd.DataFrame({"城市": ["广东/深圳", "广东_深圳", "A:B", "A_B"], "v": [1, 2, 3, 4]})
        names = sorted(splitter.build_groups(df, ["城市"]))
        self.assertEqual(names, ["A_B.xlsx", "A_B_2.xlsx", "广东_深圳.xlsx", "广东_深圳_2.xlsx"])

    def test_too_many_groups_is_rejected(self) -> None:
        df = pd.DataFrame({"k": [f"g{i}" for i in range(splitter.MAX_GROUPS + 1)]})
        with self.assertRaises(loader.ExcelToolError) as ctx:
            splitter.build_groups(df, ["k"])
        self.assertIn("分组过多", ctx.exception.message)

    def test_all_empty_column_is_rejected(self) -> None:
        df = pd.DataFrame({"k": [None, "  ", None], "v": [1, 2, 3]})
        with self.assertRaises(loader.ExcelToolError) as ctx:
            splitter.build_groups(df, ["k"])
        self.assertIn("整列都是空值", ctx.exception.message)

    def test_build_notes_contains_required_fields(self) -> None:
        notes = splitter.build_notes(
            "销售.xlsx", ["省份"], 31, datetime(2026, 11, 15, 9, 30, 0)
        )
        for text in ("销售.xlsx", "省份", "31", "2026-11-15 09:30:00"):
            self.assertIn(text, notes)


class TestMerger(unittest.TestCase):
    def setUp(self) -> None:
        self.first = pd.DataFrame(
            {"订单号": ["001", "002"], "省份": ["广东", "江苏"], "金额": [10, 20]}
        )
        self.second = pd.DataFrame({"省份": ["浙江"], "金额": [30], "城市": ["杭州"]})

    def test_union_aligns_columns_and_rows(self) -> None:
        merged = merger.merge_tables([self.first, self.second])
        self.assertEqual(list(merged.columns), ["订单号", "省份", "金额", "城市"])
        self.assertEqual(len(merged), 3)
        self.assertTrue(pd.isna(merged.loc[2, "订单号"]))

    def test_intersection_keeps_only_common_columns(self) -> None:
        merged = merger.merge_tables([self.first, self.second], mode=merger.ALIGN_INTERSECT)
        self.assertEqual(list(merged.columns), ["省份", "金额"])

    def test_same_header_is_plain_concat(self) -> None:
        merged = merger.merge_tables([self.first, self.first])
        expected = pd.concat([self.first, self.first], ignore_index=True)
        pd.testing.assert_frame_equal(merged, expected)

    def test_source_column_value(self) -> None:
        merged = merger.merge_tables(
            [self.first, self.second], labels=["a.xlsx", "b.xlsx"], source_column="来源"
        )
        self.assertEqual(merged.columns[0], "来源")
        self.assertEqual(merged.loc[0, "来源"], "a.xlsx")
        self.assertEqual(merged.loc[2, "来源"], "b.xlsx")

    def test_source_column_renamed_on_clash(self) -> None:
        clash = self.first.assign(来源="手工录入")
        merged = merger.merge_tables([clash, self.second], labels=["a", "b"], source_column="来源")
        self.assertEqual(merged.columns[0], "来源_2")
        self.assertEqual(merged.loc[0, "来源_2"], "a")
        self.assertEqual(merged.loc[0, "来源"], "手工录入")
        self.assertEqual(len(merged.columns), len(set(merged.columns)))

    def test_no_common_columns_raises(self) -> None:
        with self.assertRaises(loader.ExcelToolError) as ctx:
            merger.merge_tables(
                [self.first, pd.DataFrame({"x": [1]})], mode=merger.ALIGN_INTERSECT
            )
        self.assertIn("没有交集", ctx.exception.message)

    def test_duplicate_columns_in_one_table_raises(self) -> None:
        dup = pd.DataFrame([[1, 2]], columns=["金额", "金额"])
        with self.assertRaises(loader.ExcelToolError) as ctx:
            merger.merge_tables([self.first, dup])
        self.assertIn("重复的列名", ctx.exception.message)


class TestPivot(unittest.TestCase):
    def setUp(self) -> None:
        self.df = pd.DataFrame(
            {
                "省份": ["广东", "广东", "江苏", "江苏", None],
                "季度": ["Q1", "Q2", "Q1", "Q2", "Q1"],
                "销售额": [100.0, 150.0, 200.0, 50.0, 30.0],
            }
        )

    def test_sum_matches_total(self) -> None:
        result = pivot.build_pivot(self.df, ["省份"], "季度", "销售额", "求和")
        total = result.loc[result["省份"] == pivot.MARGIN_LABEL, pivot.MARGIN_LABEL].iloc[0]
        self.assertEqual(total, self.df["销售额"].sum())
        self.assertIn(pivot.MARGIN_LABEL, result.columns)
        self.assertFalse(result.isna().any().any())

    def test_blank_group_named_and_sorted(self) -> None:
        result = pivot.build_pivot(self.df, ["省份"], None, "销售额", "求和")
        self.assertEqual(result["省份"].tolist(), ["广东", "江苏", "空值", "合计"])

    def test_second_row_dimension_flattens_header(self) -> None:
        result = pivot.build_pivot(self.df, ["省份", "季度"], None, "销售额", "求和")
        self.assertTrue(all(isinstance(name, str) for name in result.columns))
        self.assertFalse(result.columns.duplicated().any())
        self.assertEqual(result.iloc[-1, 0], pivot.MARGIN_LABEL)

    def test_mean_has_no_inf(self) -> None:
        result = pivot.build_pivot(self.df, ["省份"], "季度", "销售额", "平均")
        numbers = result.select_dtypes("number").to_numpy()
        self.assertTrue((numbers != float("inf")).all())

    def test_numeric_columns_skips_text_and_bool(self) -> None:
        df = pd.DataFrame({"名称": ["a"], "数量": [1], "金额": [1.5], "标记": [True]})
        self.assertEqual(pivot.numeric_columns(df), ["数量", "金额"])

    def test_non_numeric_value_column_raises(self) -> None:
        with self.assertRaises(loader.ExcelToolError) as ctx:
            pivot.build_pivot(self.df, ["省份"], None, "季度", "求和")
        self.assertIn("不是数值列", ctx.exception.message)

    def test_same_row_and_column_dimension_raises(self) -> None:
        with self.assertRaises(loader.ExcelToolError) as ctx:
            pivot.build_pivot(self.df, ["省份"], "省份", "销售额", "求和")
        self.assertIn("不能是同一列", ctx.exception.message)

    def test_build_pivot_tolerates_duplicate_labels(self) -> None:
        df = pd.DataFrame(
            [[101, 102, 5.0], [103, 104, 7.0]],
            columns=["订单号", "订单号", "金额"],
        )
        result = pivot.build_pivot(df, ["订单号"], None, "金额", "求和")
        self.assertFalse(result.empty)
        self.assertAlmostEqual(float(result.iloc[0, 1]), 5.0)

    def test_pie_data_returns_single_series(self) -> None:
        df = pd.DataFrame({"省份": ["广东", "江苏", "广东"], "销售额": [1, 2, 3]})
        result = pivot.build_pivot(df, ["省份"], None, "销售额", "求和")
        data = pivot.pie_data(result, ["省份"])
        self.assertEqual(list(data.columns), ["维度", "数值"])
        self.assertEqual(len(data), 2)
        self.assertAlmostEqual(float(data["数值"].sum()), 6.0)

    def test_pie_data_rejects_multi_columns(self) -> None:
        df = pd.DataFrame(
            {"省份": ["广东", "江苏"], "季度": ["Q1", "Q2"], "销售额": [1.0, 2.0]}
        )
        result = pivot.build_pivot(df, ["省份"], "季度", "销售额", "求和")
        with self.assertRaises(loader.ExcelToolError) as ctx:
            pivot.pie_data(result, ["省份"])
        self.assertIn("单一指标", ctx.exception.message)

    def test_chart_data_drops_totals(self) -> None:
        result = pivot.build_pivot(self.df, ["省份"], "季度", "销售额", "求和")
        chart = pivot.chart_data(result, ["省份"])
        self.assertEqual(chart.index.tolist(), ["广东", "江苏", "空值"])
        self.assertEqual(list(chart.columns), ["Q1", "Q2"])
        self.assertEqual(float(chart.loc["广东", "Q1"]), 100.0)

    def test_chart_data_joins_two_row_dimensions(self) -> None:
        result = pivot.build_pivot(self.df, ["省份", "季度"], None, "销售额", "求和")
        chart = pivot.chart_data(result, ["省份", "季度"])
        self.assertEqual(chart.index[0], "广东 / Q1")
        self.assertEqual(list(chart.columns), ["销售额"])


class TestCleaner(unittest.TestCase):
    def test_number_text_is_converted_with_percent(self) -> None:
        df = pd.DataFrame({"销售额": ["¥1,234", "1,234.5", "12%", None]})
        result = cleaner.clean_table(df, to_numeric=True)
        values = result.df["销售额"].tolist()
        self.assertEqual(values[0], 1234.0)
        self.assertEqual(values[1], 1234.5)
        self.assertEqual(values[2], 0.12)
        self.assertTrue(pd.isna(values[3]))
        self.assertEqual(result.numeric_columns, ("销售额",))

    def test_leading_zero_ids_stay_text(self) -> None:
        df = pd.DataFrame({"订单号": ["00123", "00456"]})
        result = cleaner.clean_table(df, to_numeric=True)
        self.assertEqual(result.df["订单号"].tolist(), ["00123", "00456"])

    def test_blank_variants_and_fullwidth_space(self) -> None:
        df = pd.DataFrame({"销售员": ["　张三 ", "李四"], "备注": ["-", "无"]})
        result = cleaner.clean_table(df)
        self.assertEqual(result.df["销售员"].iloc[0], "张三")
        self.assertTrue(result.df["备注"].isna().all())
        self.assertEqual(result.blank_cells, 2)

    def test_excel_serial_and_chinese_date(self) -> None:
        df = pd.DataFrame({"下单日期": [45292, "2024/01/06", "2024年1月7日", "-"]})
        result = cleaner.clean_table(df, to_date=True)
        self.assertEqual(result.date_columns, ("下单日期",))
        dates = result.df["下单日期"]
        self.assertEqual(dates.iloc[0], pd.Timestamp("2024-01-01"))
        self.assertEqual(dates.iloc[2], pd.Timestamp("2024-01-07"))
        self.assertTrue(pd.isna(dates.iloc[3]))
        self.assertEqual(result.date_failed, 0)

    def test_date_needs_column_name_hint(self) -> None:
        df = pd.DataFrame({"订单号": ["2024-01-01", "2024-01-02"]})
        result = cleaner.clean_table(df, to_date=True)
        self.assertFalse(pd.api.types.is_datetime64_any_dtype(result.df["订单号"]))

    def test_date_failures_are_counted(self) -> None:
        df = pd.DataFrame({"下单时间": ["2024-01-01 09:00:00", "2024-01-02", "坏"]})
        result = cleaner.clean_table(df, to_date=True)
        self.assertEqual(result.date_failed, 1)
        self.assertIn("日期解析失败 1 个", result.summary)

    def test_drop_duplicates_runs_after_trim(self) -> None:
        df = pd.DataFrame({"名称": ["x", " x ", "y"]})
        result = cleaner.clean_table(df, drop_duplicates=True)
        self.assertEqual(len(result.df), 2)
        self.assertEqual(result.dropped_duplicates, 1)

    def test_long_text_is_truncated(self) -> None:
        df = pd.DataFrame({"备注": ["a" * (cleaner.MAX_CELL_LENGTH + 10)]})
        result = cleaner.clean_table(df)
        self.assertEqual(len(result.df["备注"].iloc[0]), cleaner.MAX_CELL_LENGTH)
        self.assertEqual(result.truncated_cells, 1)

    def test_summary_mentions_numeric_and_blank(self) -> None:
        df = pd.DataFrame({"销售额": ["¥1,234", "-"]})
        result = cleaner.clean_table(df, to_numeric=True)
        self.assertIn("数值化 1 列", result.summary)
        self.assertIn("识别空值 1 个", result.summary)


class TestExporter(unittest.TestCase):
    def test_xlsx_roundtrip(self) -> None:
        df = pd.DataFrame({"a": [1, 2], "b": ["x", "y"]})
        back = pd.read_excel(io.BytesIO(to_xlsx_bytes(df, "数据")), sheet_name="数据")
        pd.testing.assert_frame_equal(back, df)

    def test_zip_contains_all_files(self) -> None:
        data = to_zip_bytes({"a.xlsx": b"1", splitter.NOTES_FILENAME: b"2"})
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            self.assertEqual(archive.namelist(), ["a.xlsx", splitter.NOTES_FILENAME])
            self.assertIsNone(archive.testzip())


if __name__ == "__main__":
    unittest.main()
