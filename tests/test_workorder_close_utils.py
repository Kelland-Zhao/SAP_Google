import datetime
import unittest

from workorder_close_utils import (
    ActionRecord,
    ExecutionResult,
    OrderDecision,
    PlanResult,
    build_log_rows,
    build_plan,
    classify_status,
    execute_plan,
    format_orders,
    is_due_for_close,
    parse_only_orders,
    parse_sheet_date,
    parse_status,
    rows_from_values,
)


class ParseStatusTests(unittest.TestCase):
    def test_splits_codes_and_collapses_multiple_spaces(self):
        self.assertEqual(
            parse_status("TECO CNF  JBFI NMAT PRC  SETC"),
            {"TECO", "CNF", "JBFI", "NMAT", "PRC", "SETC"},
        )

    def test_uppercases_latin_codes(self):
        self.assertEqual(parse_status("rel cnf"), {"REL", "CNF"})

    def test_empty_and_none_become_empty_set(self):
        self.assertEqual(parse_status(""), set())
        self.assertEqual(parse_status(None), set())


class ParseSheetDateTests(unittest.TestCase):
    def test_parses_iso_and_compact_and_slash_forms(self):
        self.assertEqual(parse_sheet_date("2026-10-05"), datetime.date(2026, 10, 5))
        self.assertEqual(parse_sheet_date("20261005"), datetime.date(2026, 10, 5))
        self.assertEqual(parse_sheet_date("2026/10/05"), datetime.date(2026, 10, 5))

    def test_blank_and_garbage_return_none(self):
        self.assertIsNone(parse_sheet_date(""))
        self.assertIsNone(parse_sheet_date(None))
        self.assertIsNone(parse_sheet_date("not-a-date"))

    def test_accepts_date_objects(self):
        self.assertEqual(parse_sheet_date(datetime.date(2026, 10, 5)),
                         datetime.date(2026, 10, 5))


class DueForCloseTests(unittest.TestCase):
    def test_includes_cutoff_and_older(self):
        run = datetime.date(2026, 11, 2)
        self.assertTrue(is_due_for_close(datetime.date(2026, 10, 19), run))
        self.assertTrue(is_due_for_close(datetime.date(2026, 10, 18), run))
        self.assertTrue(is_due_for_close(datetime.date(2026, 1, 1), run))

    def test_rejects_recent_and_future(self):
        run = datetime.date(2026, 11, 2)
        self.assertFalse(is_due_for_close(datetime.date(2026, 10, 20), run))
        self.assertFalse(is_due_for_close(run, run))
        self.assertFalse(is_due_for_close(datetime.date(2026, 11, 3), run))


class ClassifyStatusTests(unittest.TestCase):
    def test_released_without_confirmation_needs_cnf(self):
        self.assertEqual(classify_status({"REL"}), ("CNF", ""))

    def test_released_and_confirmed_needs_teco(self):
        self.assertEqual(classify_status({"REL", "CNF"}), ("TECO", ""))

    def test_teco_and_cnf_is_managed(self):
        self.assertEqual(
            classify_status({"TECO", "CNF"}), ("SKIP", "已管理(TECO+CNF)")
        )

    def test_teco_without_cnf_is_anomaly(self):
        self.assertEqual(classify_status({"TECO"}), ("SKIP", "TECO无CNF"))

    def test_cnf_without_rel_is_anomaly(self):
        self.assertEqual(classify_status({"CNF"}), ("SKIP", "CNF无REL"))

    def test_created_is_not_released(self):
        self.assertEqual(classify_status({"CRTD"}), ("SKIP", "未释放"))


class ParseOnlyOrdersTests(unittest.TestCase):
    def test_none_and_blank_return_none(self):
        self.assertIsNone(parse_only_orders(None))
        self.assertIsNone(parse_only_orders("  "))

    def test_splits_on_commas_and_semicolons(self):
        self.assertEqual(parse_only_orders("5609259, 5611111;5620201"),
                         {"5609259", "5611111", "5620201"})


class FormatOrdersTests(unittest.TestCase):
    def test_empty_list_returns_marker(self):
        self.assertEqual(format_orders([]), "（无）")

    def test_joins_order_numbers(self):
        decisions = [
            OrderDecision("A1", "PM02", "2026-10-25", "REL", "CNF", ""),
            OrderDecision("A2", "PM03", "2026-10-25", "REL", "CNF", ""),
        ]
        self.assertEqual(format_orders(decisions), "A1, A2")


HEADER = ["参考日期", "通知", "订单", "功能位置", "设备", "描述", "成本中心",
          "系统状态", "ABC 标识", "订单类型", "基本开始日期", "基本完成日期",
          "实际开始日期", "实际完成日期"]


def sheet_row(order="5641705", status="REL", order_type="PM02",
              start="2026-10-05"):
    return ["20261020", "", order, "CN15/UTIL/ELEC", "10132240",
            "1#电容柜季保养", "CN03101660", status, "A", order_type, start,
            "2026-10-05", "2026-10-05", "2026-10-20"]


def values(*rows):
    return [HEADER] + list(rows)


class RowsFromValuesTests(unittest.TestCase):
    def test_maps_columns_by_header_name_ignoring_extra_spaces(self):
        header = list(HEADER)
        header[8] = "ABC标识"
        rows = rows_from_values([header, sheet_row()])
        self.assertEqual(rows[0].order, "5641705")
        self.assertEqual(rows[0].status_raw, "REL")
        self.assertEqual(rows[0].order_type, "PM02")
        self.assertEqual(rows[0].start_date_raw, "2026-10-05")

    def test_missing_required_header_raises(self):
        from gap_utils import MissingColumnsError
        header = [name for name in HEADER if name != "系统状态"]
        with self.assertRaises(MissingColumnsError):
            rows_from_values([header])


class BuildPlanTests(unittest.TestCase):
    RUN_DATE = datetime.date(2026, 11, 2)

    def test_classifies_groups_and_counts(self):
        plan = build_plan(values(
            sheet_row(order="A1", status="REL", order_type="PM02"),
            sheet_row(order="B1", status="REL CNF", order_type="PM03"),
            sheet_row(order="M1", status="TECO CNF", order_type="PM02"),
            sheet_row(order="T1", status="TECO", order_type="PM02"),
            sheet_row(order="P1", status="REL", order_type="PM01"),
            sheet_row(order="O1", status="REL", start="2026-10-20"),
            sheet_row(order="I1", status="REL", start="bad-date"),
        ), self.RUN_DATE)
        self.assertEqual([d.order for d in plan.cnf_orders], ["A1"])
        self.assertEqual([d.order for d in plan.teco_orders], ["B1"])
        self.assertEqual({d.order: d.reason for d in plan.skipped},
                         {"M1": "已管理(TECO+CNF)", "T1": "TECO无CNF"})
        self.assertEqual(plan.stats["pm01"], 1)
        self.assertEqual(plan.stats["too_recent"], 1)
        self.assertEqual(plan.stats["invalid"], 1)
        self.assertEqual(plan.stats["candidates"], 4)
        self.assertEqual(plan.stats["managed"], 1)
        self.assertEqual(plan.stats["anomalies"], 1)

    def test_deduplicates_and_warns_on_conflict(self):
        plan = build_plan(values(
            sheet_row(order="X1", status="REL"),
            sheet_row(order="X1", status="REL CNF"),
        ), self.RUN_DATE)
        self.assertEqual([d.order for d in plan.cnf_orders], ["X1"])
        self.assertEqual(plan.stats["duplicates"], 1)
        self.assertEqual(len(plan.warnings), 1)
        self.assertIn("X1", plan.warnings[0])

    def test_only_orders_filters_action_lists(self):
        plan = build_plan(values(
            sheet_row(order="A1", status="REL"),
            sheet_row(order="B1", status="REL CNF"),
        ), self.RUN_DATE, only_orders={"A1"})
        self.assertEqual([d.order for d in plan.cnf_orders], ["A1"])
        self.assertEqual(plan.teco_orders, [])
        self.assertEqual(plan.stats["candidates"], 2)

    def test_blank_order_or_type_is_invalid(self):
        plan = build_plan(values(
            sheet_row(order=""),
            sheet_row(order="A1", order_type=""),
        ), self.RUN_DATE)
        self.assertEqual(plan.stats["invalid"], 2)
        self.assertEqual(plan.stats["candidates"], 0)


class ExecutePlanTests(unittest.TestCase):
    def _plan(self, cnf_orders, teco_orders=()):
        def decision(order, action):
            return OrderDecision(order, "PM02", "2026-10-05", "REL", action, "")
        return PlanResult(
            [decision(o, "CNF") for o in cnf_orders],
            [decision(o, "TECO") for o in teco_orders],
            [], {}, [],
        )

    def test_confirmation_is_one_batch_call_and_all_join_teco(self):
        confirmed = []
        submitted = []
        result = execute_plan(
            self._plan(["A1", "A2"], ["B1"]),
            confirm_fn=lambda orders: (confirmed.extend(orders) or (True, "已保存")),
            teco_fn=lambda orders: (submitted.extend(orders) or (True, "已更改")),
        )
        self.assertEqual(confirmed, ["A1", "A2"])
        self.assertEqual(submitted, ["B1", "A1", "A2"])
        self.assertEqual(result.teco_submitted, ["B1", "A1", "A2"])
        self.assertEqual(result.exit_code, 0)
        self.assertEqual([r.action for r in result.records],
                         ["工时确认", "工时确认", "技术完成", "技术完成", "技术完成"])

    def test_failed_confirmation_keeps_group_b_teco_and_exit_1(self):
        submitted = []
        result = execute_plan(
            self._plan(["A1", "A2"], ["B1"]),
            confirm_fn=lambda orders: (False, "错误"),
            teco_fn=lambda orders: (submitted.extend(orders) or (True, "ok")),
        )
        self.assertEqual(submitted, ["B1"])
        self.assertEqual(result.exit_code, 1)
        self.assertEqual([r.success for r in result.records],
                         [False, False, True])

    def test_teco_failure_marks_all_teco_records_failed_and_exit_1(self):
        result = execute_plan(
            self._plan([], ["B1", "B2"]),
            confirm_fn=lambda orders: (True, "ok"),
            teco_fn=lambda orders: (False, "未选择订单"),
        )
        self.assertEqual([r.success for r in result.records], [False, False])
        self.assertEqual(result.exit_code, 1)

    def test_no_work_returns_exit_0(self):
        result = execute_plan(self._plan([]), lambda orders: (True, ""),
                              lambda orders: (True, ""))
        self.assertEqual(result.records, [])
        self.assertEqual(result.exit_code, 0)


class BuildLogRowsTests(unittest.TestCase):
    def test_skipped_and_action_rows_have_nine_columns_in_order(self):
        plan = PlanResult(
            [OrderDecision("A1", "PM02", "2026-10-05", "REL", "CNF", "")],
            [],
            [OrderDecision("S1", "PM02", "2026-10-05", "CRTD", "SKIP", "未释放")],
            {}, [],
        )
        execution = ExecutionResult(
            [ActionRecord(plan.cnf_orders[0], "工时确认", True, "已保存")],
            ["A1"], 0,
        )
        rows = build_log_rows(plan, execution, datetime.date(2026, 11, 2),
                              "2026-11-02 07:00:00")
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0], [
            "2026-11-02 07:00:00", "2026-11-02", "S1", "PM02",
            "2026-10-05", "CRTD", "跳过", "跳过", "未释放",
        ])
        self.assertEqual(rows[1], [
            "2026-11-02 07:00:00", "2026-11-02", "A1", "PM02",
            "2026-10-05", "REL", "工时确认", "成功", "已保存",
        ])
        self.assertTrue(all(len(row) == 9 for row in rows))
