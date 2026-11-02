"""09 - 工单自动关闭的纯逻辑。

放在仓库根、且不 import win32com / gspread，是为了能在任何机器上跑单元测试。
main.py 只负责 SAP / Google Sheets 的管路，筛选、分组、执行编排和日志行都在这里。
"""

from __future__ import annotations

import datetime
from dataclasses import dataclass

from gap_utils import resolve_columns

WINDOW_DAYS = 14
REQUIRED_COLUMNS = ["订单", "系统状态", "订单类型", "基本开始日期"]
LOG_HEADERS = [
    "执行时间", "运行日期", "订单", "订单类型", "基本开始日期",
    "处理前系统状态", "动作", "结果", "原因 / SAP 消息",
]

# 和 gap_utils 保持一致的日期写法集合。%Y%m%d 放在带分隔符的写法之后，
# 互不冲突；全认不出来就返回 None。
_DATE_FORMATS = (
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%Y%m%d",
    "%m/%d/%Y", "%d.%m.%Y",
)


@dataclass(frozen=True)
class OrderRow:
    order: str
    order_type: str
    start_date_raw: str
    status_raw: str


@dataclass(frozen=True)
class OrderDecision:
    order: str
    order_type: str
    start_date_raw: str
    status_raw: str
    action: str  # "CNF" | "TECO" | "SKIP"
    reason: str  # SKIP 的原因；CNF/TECO 为空串


@dataclass
class PlanResult:
    cnf_orders: list[OrderDecision]
    teco_orders: list[OrderDecision]
    skipped: list[OrderDecision]
    stats: dict[str, int]
    warnings: list[str]


@dataclass(frozen=True)
class ActionRecord:
    decision: OrderDecision
    action: str  # "工时确认" | "技术完成"
    success: bool
    message: str


@dataclass
class ExecutionResult:
    records: list[ActionRecord]
    teco_submitted: list[str]
    exit_code: int


def parse_status(raw):
    """把系统状态串拆成整词集合；大小写不敏感。"""
    if raw is None:
        return set()
    text = str(raw).strip()
    if not text:
        return set()
    return {token.upper() for token in text.split()}


def parse_sheet_date(raw):
    """解析 Total_Workorder 的日期单元格；认不出来返回 None。"""
    if isinstance(raw, datetime.datetime):
        return raw.date()
    if isinstance(raw, datetime.date):
        return raw
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    for fmt in _DATE_FORMATS:
        try:
            return datetime.datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def within_window(start, run_date, window_days=WINDOW_DAYS):
    """闭区间：run_date 往前 window_days 天 到 run_date。"""
    return run_date - datetime.timedelta(days=window_days) <= start <= run_date


def classify_status(statuses):
    """按状态集合决定动作：CNF、TECO 或 SKIP（带原因）。"""
    has_rel = "REL" in statuses
    has_cnf = "CNF" in statuses
    has_teco = "TECO" in statuses

    if has_teco and has_cnf:
        return ("SKIP", "已管理(TECO+CNF)")
    if has_rel and not has_cnf and not has_teco:
        return ("CNF", "")
    if has_rel and has_cnf and not has_teco:
        return ("TECO", "")
    if has_teco and not has_cnf:
        return ("SKIP", "TECO无CNF")
    if has_cnf and not has_rel and not has_teco:
        return ("SKIP", "CNF无REL")
    if has_rel and has_teco:
        return ("SKIP", "REL+TECO")
    return ("SKIP", "未释放")


def parse_only_orders(raw):
    """解析 --only-orders；逗号和分号都接受。空值返回 None。"""
    if raw is None:
        return None
    text = str(raw).strip()
    if not text:
        return None
    parts = {
        part.strip()
        for part in text.replace(";", ",").split(",")
        if part.strip()
    }
    return parts or None


def rows_from_values(values):
    """把 Total_Workorder 的原始行转成 OrderRow；全空行跳过。"""
    if not values:
        raise ValueError("工作表为空")

    columns = resolve_columns(values[0], REQUIRED_COLUMNS)
    rows = []
    for raw in values[1:]:
        if not any(str(cell).strip() for cell in raw):
            continue

        def cell(name):
            index = columns[name]
            return str(raw[index]).strip() if index < len(raw) else ""

        rows.append(OrderRow(
            order=cell("订单"),
            order_type=cell("订单类型"),
            start_date_raw=cell("基本开始日期"),
            status_raw=cell("系统状态"),
        ))
    return rows


def build_plan(values, run_date, window_days=WINDOW_DAYS, only_orders=None):
    """筛选、分组并统计；only_orders 只过滤动作列表，不改变统计。"""
    rows = rows_from_values(values)
    stats = {
        "total_rows": 0,
        "pm01": 0,
        "out_of_window": 0,
        "invalid": 0,
        "duplicates": 0,
        "candidates": 0,
        "managed": 0,
        "anomalies": 0,
    }
    cnf_orders, teco_orders, skipped = [], [], []
    warnings = []
    seen = {}

    for row in rows:
        stats["total_rows"] += 1

        if not row.order:
            stats["invalid"] += 1
            continue
        if not row.order_type:
            stats["invalid"] += 1
            continue
        if row.order_type.upper() == "PM01":
            stats["pm01"] += 1
            continue

        start = parse_sheet_date(row.start_date_raw)
        if start is None:
            stats["invalid"] += 1
            continue
        if not within_window(start, run_date, window_days):
            stats["out_of_window"] += 1
            continue

        if row.order in seen:
            stats["duplicates"] += 1
            if seen[row.order] != row.status_raw:
                warnings.append(
                    f"订单 {row.order} 重复且状态不一致："
                    f"{seen[row.order]!r} / {row.status_raw!r}，保留第一条"
                )
            continue
        seen[row.order] = row.status_raw

        action, reason = classify_status(parse_status(row.status_raw))
        decision = OrderDecision(
            order=row.order,
            order_type=row.order_type,
            start_date_raw=row.start_date_raw,
            status_raw=row.status_raw,
            action=action,
            reason=reason,
        )
        if action == "CNF":
            cnf_orders.append(decision)
        elif action == "TECO":
            teco_orders.append(decision)
        else:
            skipped.append(decision)
        stats["candidates"] += 1

    stats["managed"] = sum(
        1 for decision in skipped if decision.reason == "已管理(TECO+CNF)"
    )
    stats["anomalies"] = len(skipped) - stats["managed"]

    if only_orders is not None:
        cnf_orders = [d for d in cnf_orders if d.order in only_orders]
        teco_orders = [d for d in teco_orders if d.order in only_orders]
        skipped = [d for d in skipped if d.order in only_orders]

    return PlanResult(cnf_orders, teco_orders, skipped, stats, warnings)


def execute_plan(plan, confirm_fn, teco_fn):
    """批量执行：先一次工时确认，再一次批量技术完成。"""
    records = []
    teco_decisions = list(plan.teco_orders)
    any_failure = False

    if plan.cnf_orders:
        ok, message = confirm_fn([d.order for d in plan.cnf_orders])
        for decision in plan.cnf_orders:
            records.append(ActionRecord(decision, "工时确认", ok, message))
        if ok:
            teco_decisions.extend(plan.cnf_orders)
        else:
            any_failure = True

    submitted = [d.order for d in teco_decisions]
    if teco_decisions:
        ok, message = teco_fn(submitted)
        for decision in teco_decisions:
            records.append(ActionRecord(decision, "技术完成", ok, message))
        if not ok:
            any_failure = True

    return ExecutionResult(records, submitted, 1 if any_failure else 0)


def build_log_rows(plan, execution, run_date, timestamp):
    """按 WorkOrderClose_Log 的 9 列约定组装追加行。"""
    run_date_text = run_date.isoformat()
    rows = []

    for decision in plan.skipped:
        rows.append([
            timestamp, run_date_text, decision.order, decision.order_type,
            decision.start_date_raw, decision.status_raw,
            "跳过", "跳过", decision.reason,
        ])

    for record in execution.records:
        decision = record.decision
        rows.append([
            timestamp, run_date_text, decision.order, decision.order_type,
            decision.start_date_raw, decision.status_raw,
            record.action, "成功" if record.success else "失败",
            record.message,
        ])

    return rows
