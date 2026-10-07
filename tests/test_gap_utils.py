import datetime
import unittest

from gap_utils import (
    GAP_HEADERS,
    MissingColumnsError,
    build_gap_rows,
    format_reference_date,
    resolve_columns,
    split_rows_by_month,
)

# IW39（变式 /KEL）导出文件那一行的表头，10 列
EXPORT_HEADER = [
    "参考日期", "通知", "订单", "功能位置", "设备",
    "描述", "成本中心", "系统状态", "ABC 标识", "订单类型",
]


def export_row(reference_date="2026-10-15", status="REL", order="400123"):
    """一行导出数据，只有参考日期、系统状态和订单号是测试关心的。"""
    return [
        reference_date, "100001", order, "CN15-A-01", "10001",
        "给水泵", "CC01", status, "A", "PM02",
    ]


def sheet_row(year_month="202610", write_time="2026-10-07 06:00:00", order="400123"):
    """Gap 表里的一行（12 列）。"""
    return [
        "2026-10-15", "100001", order, "CN15-A-01", "10001",
        "给水泵", "CC01", "REL", "A", "PM02", write_time, year_month,
    ]


class ResolveColumnsTests(unittest.TestCase):
    def test_maps_every_required_header_to_its_index(self):
        columns = resolve_columns(EXPORT_HEADER)

        self.assertEqual(columns["参考日期"], 0)
        self.assertEqual(columns["系统状态"], 7)
        self.assertEqual(columns["订单类型"], 9)

    def test_ignores_whitespace_differences_in_headers(self):
        header = list(EXPORT_HEADER)
        header[8] = "ABC标识"

        columns = resolve_columns(header)

        self.assertEqual(columns["ABC 标识"], 8)

    def test_rejects_a_missing_header_and_lists_what_is_available(self):
        header = [name for name in EXPORT_HEADER if name != "成本中心"]

        with self.assertRaises(MissingColumnsError) as ctx:
            resolve_columns(header)

        message = str(ctx.exception)
        self.assertIn("成本中心", message)
        self.assertIn("功能位置", message)  # 实际有哪些表头，要打出来

    def test_does_not_match_a_longer_header_for_a_shorter_name(self):
        # "订单" 没找到时，绝不能拿 "订单类型" 顶上
        header = [name for name in EXPORT_HEADER if name != "订单"]

        with self.assertRaises(MissingColumnsError):
            resolve_columns(header)


class FormatReferenceDateTests(unittest.TestCase):
    def test_formats_a_datetime_cell(self):
        self.assertEqual(
            format_reference_date(datetime.datetime(2026, 10, 15, 0, 0)),
            "2026-10-15",
        )

    def test_formats_a_date_cell(self):
        self.assertEqual(format_reference_date(datetime.date(2026, 10, 15)), "2026-10-15")

    def test_keeps_an_iso_string(self):
        self.assertEqual(format_reference_date("2026-10-15"), "2026-10-15")

    def test_parses_a_us_style_string(self):
        # 脚本给 SAP 填的日期就是 %m/%d/%Y，说明这台机器的 SAP 日期格式是月在前
        self.assertEqual(format_reference_date("10/15/2026"), "2026-10-15")

    def test_parses_a_day_first_dotted_string(self):
        self.assertEqual(format_reference_date("15.10.2026"), "2026-10-15")

    def test_parses_a_year_first_dotted_string(self):
        self.assertEqual(format_reference_date("2026.10.15"), "2026-10-15")

    def test_returns_empty_for_a_blank_cell(self):
        self.assertEqual(format_reference_date(None), "")
        self.assertEqual(format_reference_date("   "), "")

    def test_keeps_an_unrecognized_string_verbatim(self):
        # 认不出来也不能丢数据，原样写进去让人能看见
        self.assertEqual(format_reference_date("待定"), "待定")


class SplitRowsByMonthTests(unittest.TestCase):
    def test_keeps_other_months_and_drops_the_target_month(self):
        header = sheet_row(year_month="月份")  # 占位，内容不重要
        all_rows = [header, sheet_row("202609"), sheet_row("202610")]

        keep, _ = split_rows_by_month(all_rows, "202610")

        self.assertEqual(keep, [header, sheet_row("202609")])

    def test_returns_the_first_write_time_of_the_target_month(self):
        all_rows = [
            sheet_row("月份"),
            sheet_row("202610", "2026-10-07 06:00:00", order="1"),
            sheet_row("202610", "2026-10-07 06:00:00", order="2"),
        ]

        _, first_write_time = split_rows_by_month(all_rows, "202610")

        self.assertEqual(first_write_time, "2026-10-07 06:00:00")

    def test_returns_no_write_time_when_the_month_is_absent(self):
        all_rows = [sheet_row("月份"), sheet_row("202609")]

        _, first_write_time = split_rows_by_month(all_rows, "202610")

        self.assertIsNone(first_write_time)

    def test_keeps_the_header_when_nothing_is_dropped(self):
        header = sheet_row("月份")
        all_rows = [header, sheet_row("202609")]

        keep, _ = split_rows_by_month(all_rows, "202610")

        self.assertEqual(keep[0], header)

    def test_handles_an_empty_sheet(self):
        self.assertEqual(split_rows_by_month([], "202610"), ([], None))

    def test_keeps_short_rows_that_have_no_month(self):
        # 列数不足的行认不出月份，一律保留，绝不误删
        short_row = ["2026-10-15", "100001"]
        all_rows = [sheet_row("月份"), short_row]

        keep, _ = split_rows_by_month(all_rows, "202610")

        self.assertIn(short_row, keep)

    def test_matches_a_numeric_month_cell(self):
        row = sheet_row("202610")
        row[11] = 202610

        keep, write_time = split_rows_by_month([sheet_row("月份"), row], "202610")

        self.assertEqual(len(keep), 1)
        self.assertEqual(write_time, "2026-10-07 06:00:00")


class BuildGapRowsTests(unittest.TestCase):
    def setUp(self):
        self.columns = resolve_columns(EXPORT_HEADER)

    def test_selects_rows_without_cnf(self):
        rows = [export_row(status="REL"), export_row(status="CNF")]

        result = build_gap_rows(rows, self.columns, "2026-10-07 06:00:00", "202610")

        self.assertEqual(len(result), 1)
        self.assertEqual(result[0][7], "REL")

    def test_treats_status_case_insensitively(self):
        rows = [export_row(status="cnf"), export_row(status="TECO")]

        result = build_gap_rows(rows, self.columns, "T", "202610")

        self.assertEqual([row[7] for row in result], ["TECO"])

    def test_skips_rows_with_a_blank_status(self):
        rows = [export_row(status=""), export_row(status=None)]

        result = build_gap_rows(rows, self.columns, "T", "202610")

        self.assertEqual(result, [])

    def test_writes_the_agreed_column_order(self):
        row = export_row(reference_date=datetime.datetime(2026, 10, 15), status="REL")

        result = build_gap_rows([row], self.columns, "2026-10-07 06:00:00", "202610")

        self.assertEqual(result[0], [
            "2026-10-15", "100001", "400123", "CN15-A-01", "10001",
            "给水泵", "CC01", "REL", "A", "PM02",
            "2026-10-07 06:00:00", "202610",
        ])
        self.assertEqual(len(GAP_HEADERS), len(result[0]))

    def test_normalizes_the_reference_date_column(self):
        row = export_row(reference_date=datetime.date(2026, 10, 15), status="REL")

        result = build_gap_rows([row], self.columns, "T", "202610")

        self.assertEqual(result[0][0], "2026-10-15")

    def test_tolerates_rows_shorter_than_the_header(self):
        row = export_row(status="REL")[:8]  # 后面的列全是空的

        result = build_gap_rows([row], self.columns, "T", "202610")

        self.assertEqual(result[0][8:10], ["", ""])


class RefreshScenarioTests(unittest.TestCase):
    """把两个函数串起来，验证「每天跑、同一个月被反复处理」的真实效果。"""

    def setUp(self):
        self.columns = resolve_columns(EXPORT_HEADER)

    def test_rerunning_a_month_refreshes_rows_but_keeps_the_first_write_time(self):
        # 第一次运行（10-07）：10 月有两条未确认工单
        keep, first_write = split_rows_by_month([GAP_HEADERS], "202610")
        self.assertIsNone(first_write)
        write_time = first_write or "2026-10-07 06:00:00"
        sheet = keep + build_gap_rows(
            [export_row(order="400123"), export_row(order="400124")],
            self.columns, write_time, "202610",
        )

        # 第二次运行（10-08）：400123 已确认，只剩 400124 未确认
        keep, first_write = split_rows_by_month(sheet, "202610")
        self.assertEqual(first_write, "2026-10-07 06:00:00")
        sheet = keep + build_gap_rows(
            [export_row(order="400123", status="CNF"), export_row(order="400124")],
            self.columns, first_write, "202610",
        )

        self.assertEqual(len(sheet), 2)  # 表头 + 1 行，旧的 2 行没有留下来
        self.assertEqual(sheet[1][2], "400124")
        self.assertEqual(sheet[1][11], "202610")
        self.assertEqual(sheet[1][10], "2026-10-07 06:00:00")  # 时间没被刷新

    def test_refreshing_one_month_leaves_the_other_months_untouched(self):
        september = build_gap_rows(
            [export_row(order="400100")], self.columns, "2026-09-07 06:00:00", "202609"
        )
        sheet = [GAP_HEADERS] + september

        keep, first_write = split_rows_by_month(sheet, "202610")
        self.assertIsNone(first_write)  # 10 月还没写过，不该沿用 9 月的时间
        october = build_gap_rows(
            [export_row(order="400200")], self.columns, "2026-10-07 06:00:00", "202610"
        )

        self.assertEqual(keep + october, [GAP_HEADERS] + september + october)


class GapHeadersTests(unittest.TestCase):
    def test_headers_are_the_agreed_twelve_columns(self):
        self.assertEqual(GAP_HEADERS, [
            "参考日期", "通知", "订单", "功能位置", "设备",
            "描述", "成本中心", "系统状态", "ABC 标识", "订单类型",
            "写入时间", "月份",
        ])


if __name__ == "__main__":
    unittest.main()
