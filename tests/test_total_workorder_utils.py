import datetime
import unittest

from total_workorder_utils import get_date_range


class GetDateRangeTests(unittest.TestCase):
    def test_uses_365_days_before_and_after(self):
        start, end = get_date_range(datetime.date(2026, 11, 2))
        self.assertEqual(start, "11/02/2025")
        self.assertEqual(end, "11/02/2027")

    def test_formats_as_mm_dd_yyyy(self):
        start, end = get_date_range(datetime.date(2026, 1, 1))
        self.assertEqual(start, "01/01/2025")
        self.assertEqual(end, "01/01/2027")

    def test_handles_leap_day(self):
        start, end = get_date_range(datetime.date(2024, 3, 1))
        self.assertEqual(start, "03/02/2023")
        self.assertEqual(end, "03/01/2025")
