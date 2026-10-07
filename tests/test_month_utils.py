import datetime
import unittest

from month_utils import months_to_sync, should_upload_month_data


class MonthUtilsTests(unittest.TestCase):
    def test_months_to_sync_returns_previous_then_current_month(self):
        self.assertEqual(
            months_to_sync(datetime.date(2026, 9, 7)),
            ["202608", "202609"],
        )

    def test_months_to_sync_handles_january_year_boundary(self):
        self.assertEqual(
            months_to_sync(datetime.date(2026, 1, 2)),
            ["202512", "202601"],
        )

    def test_months_to_sync_rejects_negative_lookback(self):
        with self.assertRaisesRegex(ValueError, "lookback"):
            months_to_sync(datetime.date(2026, 9, 7), lookback=-1)

    def test_current_month_without_data_is_skipped(self):
        self.assertFalse(
            should_upload_month_data(None, "202609", "202609")
        )

    def test_historical_month_without_data_is_a_failure(self):
        with self.assertRaisesRegex(RuntimeError, "202608"):
            should_upload_month_data(None, "202608", "202609")

    def test_month_with_data_is_uploaded(self):
        self.assertTrue(
            should_upload_month_data({"value": 1}, "202608", "202609")
        )


if __name__ == "__main__":
    unittest.main()
