"""04 - Total_Workorder 的纯逻辑。

放在仓库根、且不 import win32com / pandas，是为了能在任何机器上跑单元测试。
"""

from __future__ import annotations

import datetime

LOOKBACK_DAYS = 365
LOOKAHEAD_DAYS = 365


def get_date_range(today=None):
    """返回 IW39 的起止日期（MM/DD/YYYY）：today-365 天 ~ today+365 天。"""
    if today is None:
        today = datetime.date.today()
    start = today - datetime.timedelta(days=LOOKBACK_DAYS)
    end = today + datetime.timedelta(days=LOOKAHEAD_DAYS)
    return start.strftime("%m/%d/%Y"), end.strftime("%m/%d/%Y")
