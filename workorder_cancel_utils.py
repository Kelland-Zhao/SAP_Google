"""10 - 工单取消技术完成的纯逻辑。

放在仓库根、且不 import win32com / gspread，是为了能在任何机器上跑单元测试。
main.py 只负责 SAP / Google Sheets 的管路；读行、去重、计划和单元格更新都在这里。
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Callable

DONE = "Done"
COL_ORDER = 1
COL_MARKER = 2
COL_REASON = 3


@dataclass(frozen=True)
class SheetRow:
    row_number: int
    order: str
    marker: str
    reason: str


@dataclass(frozen=True)
class PendingOrder:
    order: str
    rows: tuple[SheetRow, ...]


@dataclass
class PlanResult:
    orders: list[PendingOrder]
    stats: dict[str, int]
    warnings: list[str]


@dataclass(frozen=True)
class ActionRecord:
    order: str
    success: bool
    message: str
    row_numbers: tuple[int, ...]


@dataclass
class ExecutionResult:
    records: list[ActionRecord]
    pending_writes: list[tuple[int, int, str]]
    aborted: bool
    exit_code: int


class SessionLostError(RuntimeError):
    """SAP 会话不可用，整轮应中止。"""


def _cell(raw, index):
    if index >= len(raw):
        return ""
    return str(raw[index]).strip()


def parse_rows(values):
    """把工作表原始行转成 SheetRow；A/B/C 全空的行跳过。"""
    rows = []
    for index, raw in enumerate(values or []):
        order = _cell(raw, 0)
        marker = _cell(raw, 1)
        reason = _cell(raw, 2)
        if not (order or marker or reason):
            continue
        rows.append(SheetRow(index + 1, order, marker, reason))
    return rows


def build_plan(values, only_orders=None):
    """跳过已处理行，按订单号去重，保留 Sheet 中的原始顺序。"""
    rows = parse_rows(values)
    stats = {
        "total_rows": len(rows),
        "done": 0,
        "empty_order": 0,
        "filtered": 0,
        "duplicates": 0,
        "pending_orders": 0,
        "pending_rows": 0,
    }
    grouped = {}
    order_sequence = []
    warnings = []

    for row in rows:
        if not row.order:
            stats["empty_order"] += 1
            continue
        if only_orders is not None and row.order not in only_orders:
            stats["filtered"] += 1
            continue
        if row.marker:
            stats["done"] += 1
            continue

        if row.order not in grouped:
            grouped[row.order] = []
            order_sequence.append(row.order)
        else:
            stats["duplicates"] += 1
        grouped[row.order].append(row)

    orders = [
        PendingOrder(order, tuple(grouped[order])) for order in order_sequence
    ]
    stats["pending_orders"] = len(orders)
    stats["pending_rows"] = sum(len(order.rows) for order in orders)

    if only_orders is not None:
        present = {row.order for row in rows if row.order}
        missing = sorted(only_orders - present)
        if missing:
            warnings.append(f"这些订单在工作表中不存在: {', '.join(missing)}")

    return PlanResult(orders, stats, warnings)


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


def format_orders(plan):
    """把待处理订单列表格式化成一行，便于 dry-run 核对。"""
    if not plan.orders:
        return "（无）"
    return ", ".join(order.order for order in plan.orders)


def failure_reason(timestamp, message):
    return f"{timestamp} | {message}"


def success_updates(pending):
    """成功：B 写 Done，C 清空（清掉上一次的失败原因）。"""
    updates = []
    for row in pending.rows:
        updates.append((row.row_number, COL_MARKER, DONE))
        updates.append((row.row_number, COL_REASON, ""))
    return updates


def failure_updates(pending, reason):
    """失败：只写 C，B 保持空。"""
    return [(row.row_number, COL_REASON, reason) for row in pending.rows]


def execute_plan(
    plan: PlanResult,
    cancel_fn: Callable[[str], tuple[bool, str]],
    write_fn: Callable[[list[tuple[int, int, str]]], bool],
    timestamp: str,
) -> ExecutionResult:
    """逐单执行；会话级失败中止，单笔失败继续。"""
    records = []
    pending_writes = []
    aborted = False

    for pending in plan.orders:
        try:
            ok, message = cancel_fn(pending.order)
        except SessionLostError:
            aborted = True
            break

        if ok:
            updates = success_updates(pending)
        else:
            message = message or "SAP 未返回消息"
            updates = failure_updates(
                pending, failure_reason(timestamp, message),
            )

        records.append(ActionRecord(
            order=pending.order,
            success=ok,
            message=message,
            row_numbers=tuple(row.row_number for row in pending.rows),
        ))
        if not write_fn(updates):
            pending_writes.extend(updates)

    exit_code = 1 if (
        aborted
        or pending_writes
        or any(not record.success for record in records)
    ) else 0
    return ExecutionResult(records, pending_writes, aborted, exit_code)
