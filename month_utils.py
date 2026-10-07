"""Calendar-month helpers shared by the scheduled KPI sync scripts."""

import datetime
from typing import List, Optional


def _shift_month(first_of_month: datetime.date, offset: int) -> datetime.date:
    """Return the first day of the month ``offset`` months away."""
    month_index = first_of_month.year * 12 + first_of_month.month - 1 + offset
    year, month_zero_based = divmod(month_index, 12)
    return datetime.date(year, month_zero_based + 1, 1)


def months_to_sync(today: Optional[datetime.date] = None, lookback: int = 1) -> List[str]:
    """Return the previous ``lookback`` months and the current month.

    The result is chronological and uses the ``YYYYMM`` keys used in
    ``MasterData``.  With the default lookback, a run on 2026-09-07 returns
    ``["202608", "202609"]``.
    """
    if lookback < 0:
        raise ValueError("lookback must be non-negative")

    if today is None:
        today = datetime.date.today()

    current_month = today.replace(day=1)
    return [
        _shift_month(current_month, -offset).strftime("%Y%m")
        for offset in range(lookback, -1, -1)
    ]


def should_upload_month_data(data, year_month: str, current_month: str) -> bool:
    """Decide whether a month result may be written to Google Sheets.

    A missing current-month result is normal early in a month and should leave
    the existing row unchanged.  A missing historical-month result is a
    failed catch-up and must stop the job so the next run retries it.
    """
    if data is not None:
        return True
    if year_month == current_month:
        return False
    raise RuntimeError(f"历史月份 {year_month} 没有可上传的数据")
