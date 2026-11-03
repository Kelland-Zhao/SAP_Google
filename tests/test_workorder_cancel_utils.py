import unittest

from workorder_cancel_utils import (
    build_plan, failure_reason, failure_updates, parse_only_orders,
    parse_rows, success_updates,
)


class ParseRowsTests(unittest.TestCase):
    def test_numbers_rows_from_one_and_pads_missing_columns(self):
        rows = parse_rows([["5616063", "Done"], ["5650370"], ["", "", ""]])
        self.assertEqual(
            [(r.row_number, r.order, r.marker, r.reason) for r in rows],
            [
                (1, "5616063", "Done", ""),
                (2, "5650370", "", ""),
            ],
        )

    def test_strips_whitespace(self):
        rows = parse_rows([[" 5650370 ", " Done ", " 已锁定 "]])
        self.assertEqual(
            (rows[0].order, rows[0].marker, rows[0].reason),
            ("5650370", "Done", "已锁定"),
        )


class BuildPlanTests(unittest.TestCase):
    def test_skips_done_and_empty_orders(self):
        plan = build_plan([
            ["5616063", "Done", ""],
            ["5650370", "", ""],
            ["", "", "旧原因"],
            ["5655247", "", ""],
        ])
        self.assertEqual([o.order for o in plan.orders], ["5650370", "5655247"])
        self.assertEqual(plan.stats["total_rows"], 4)
        self.assertEqual(plan.stats["done"], 1)
        self.assertEqual(plan.stats["empty_order"], 1)
        self.assertEqual(plan.stats["pending_orders"], 2)
        self.assertEqual(plan.stats["pending_rows"], 2)

    def test_duplicate_pending_rows_merge_into_one_order(self):
        plan = build_plan([["5650370", "", ""], ["5650370", "", ""]])
        self.assertEqual(len(plan.orders), 1)
        self.assertEqual([r.row_number for r in plan.orders[0].rows], [1, 2])
        self.assertEqual(plan.stats["duplicates"], 1)

    def test_a_done_duplicate_does_not_block_a_pending_row(self):
        plan = build_plan([["5650370", "Done", ""], ["5650370", "", ""]])
        self.assertEqual([o.order for o in plan.orders], ["5650370"])
        self.assertEqual([r.row_number for r in plan.orders[0].rows], [2])

    def test_only_orders_filters_and_warns_about_missing(self):
        plan = build_plan(
            [["5650370", "", ""], ["5655247", "", ""]],
            only_orders={"5650370", "9999999"},
        )
        self.assertEqual([o.order for o in plan.orders], ["5650370"])
        self.assertEqual(plan.stats["filtered"], 1)
        self.assertTrue(any("9999999" in w for w in plan.warnings))

    def test_empty_values_produce_empty_plan(self):
        plan = build_plan([])
        self.assertEqual(plan.orders, [])
        self.assertEqual(plan.stats["pending_orders"], 0)


class UpdateTests(unittest.TestCase):
    def test_success_writes_done_and_clears_reason_for_every_row(self):
        plan = build_plan([["5650370", "", "旧原因"], ["5650370", "", ""]])
        self.assertEqual(success_updates(plan.orders[0]), [
            (1, 2, "Done"), (1, 3, ""),
            (2, 2, "Done"), (2, 3, ""),
        ])

    def test_failure_only_writes_reason(self):
        plan = build_plan([["5650370", "", ""], ["5650370", "", ""]])
        self.assertEqual(
            failure_updates(plan.orders[0], "2026-11-03 10:15:22 | 已锁定"),
            [
                (1, 3, "2026-11-03 10:15:22 | 已锁定"),
                (2, 3, "2026-11-03 10:15:22 | 已锁定"),
            ],
        )

    def test_failure_reason_format(self):
        self.assertEqual(
            failure_reason("2026-11-03 10:15:22", "已锁定"),
            "2026-11-03 10:15:22 | 已锁定",
        )


class ParseOnlyOrdersTests(unittest.TestCase):
    def test_splits_commas_and_semicolons(self):
        self.assertEqual(parse_only_orders("1, 2;3"), {"1", "2", "3"})

    def test_blank_returns_none(self):
        self.assertIsNone(parse_only_orders("  "))
