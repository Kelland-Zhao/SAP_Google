# 10 - 工单取消技术完成 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增 `10 - WorkOrderCancelTECO`：读同一 spreadsheet 的 `取消技术完成` 工作表，对 B 列为空的订单逐单执行 IW32「取消技术完成」，成功写 `Done` 到 B 列，失败原因写 C 列。一次性工具，不加入 `run_all.ps1`。

**Architecture:** 纯逻辑（读行、去重、计划、单元格更新、执行编排）放仓库根的 `workorder_cancel_utils.py`，可在 macOS 上用 `unittest` 覆盖；SAP GUI 操作放 `workorder_cancel_sap.py`，只依赖传入的 session，并从 09 的 `workorder_close_sap.py` import 状态栏/等待函数（**不修改 09**）；`10 - WorkOrderCancelTECO/main.py` 只负责 argparse、gspread 读写、SAP 会话生命周期和把三者串起来。

**Tech Stack:** Python 3.11、`pywin32`（仅 Windows）、`gspread` + `google-auth`、`unittest`。不使用 Excel/pandas。

**Spec:** `docs/superpowers/specs/2026-11-03-workorder-cancel-teco-design.md`

## Global Constraints

- 运行机：Windows + SAP GUI Scripting + Python 3.11；纯逻辑与 SAP 按键顺序可在 macOS 上跑单元测试。
- 依赖：`pywin32`、`gspread`、`google-auth`；不新增其他依赖。
- Google 服务账号文件：仓库根 `pyreadsp-b5b9c1909de6.json`。
- Spreadsheet：`1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98`；工作表 `取消技术完成`。
- 工作表**没有表头**，A1 就是第一单：A=订单号，B=`Done`，C=失败原因。
- B 列去空格后非空 → 跳过；A 列为空 → 忽略；同一订单多行 → 只调 SAP 一次，所有待处理行一起更新。
- C 列格式：`YYYY-MM-DD HH:MM:SS | 原因`；成功时清空 C。
- 成功判定：`S`/`W` + 文本非空 + 不含否定词 + 含 `保存`/`saved`；宁可误报失败，不可误报成功。
- 退出码：全部成功/无事可做 `0`；任意失败、会话中止、结果写不进 Sheet `1`。
- **不修改** `workorder_close_sap.py`、`workorder_close_utils.py` 和 `09 - WorkOrderAutoClose/main.py`；10 只 import 09 的 `read_status` / `is_success` / `wait_for_control` / `NEGATIVE_KEYWORDS`。
- 不修改 `run_all.ps1`；10 不参与每日运行。
- 控制台中文输出；订单号按文本处理。
- 录制文件 `取消技术完成IW32.txt` 复制进 `10 - WorkOrderCancelTECO/` 作为参考，不执行。
- 提交信息格式：`VYYYY-MM-DD.XX_简短描述`；实现阶段直接提交到 `main`（用户工作流是本地 `main` → push → 运行机 pull）。

## Review Focus

按最可能出问题的顺序：

1. **SAP 返回「像成功」的消息但其实没取消 TECO**：这会把 B 写成 `Done`，以后永远不再重试。Task 3 的 `SuccessIsProvenTests` 把判定锁成「必须含 `保存`/`saved` 且不含否定词」；Task 5 要求在 Windows 上先单笔实跑、用 IW33 核对状态，再跑全量。
2. **SAP 会话中途失效**：如果继续逐单处理，剩余每单都会白等两次 30 秒超时并写一堆假失败。Task 2 的 `test_abort_stops_remaining_orders` 和 Task 3 的 `SessionLostError` 测试锁定「会话级失败 → 整轮中止、剩余行完全不动」。
3. **SAP 已成功但 Sheet 写入失败**：行会停在 B 空，下次重跑会再取消一次（SAP 会报未技术完成）。Task 2 的 `test_failed_write_is_returned_for_retry` 锁定「失败写入进 `pending_writes`」；Task 4 的主程序做 3 次重试 + 结尾补写 + 明确列出需手工补写的订单。
4. **错误弹窗被回车关掉后继续保存**：可能把失败当成功。Task 3 的 `test_popup_error_skips_save` 锁定「`E`/`A` 弹窗 → 不按保存、该单失败」。
5. **同一订单多行，其中一行已 Done、另一行待处理**：必须仍处理该订单一次，并把所有待处理行写成 Done，不能因为看到一行 Done 就整单跳过。Task 1 的 `test_a_done_duplicate_does_not_block_a_pending_row` 锁定。

---

### Task 0: 提交设计与计划文档

**Files:**
- Add: `docs/superpowers/specs/2026-11-03-workorder-cancel-teco-design.md`
- Add: `docs/superpowers/plans/2026-11-03-workorder-cancel-teco.md`

**Interfaces:**
- Consumes: 已批准的设计文档。
- Produces: 仓库里可追溯的设计与计划；后续任务从干净的工作区开始。

- [ ] **Step 1: 提交文档**

```bash
git add docs/superpowers/specs/2026-11-03-workorder-cancel-teco-design.md docs/superpowers/plans/2026-11-03-workorder-cancel-teco.md
git commit -m "V2026-11-03.00_新增10取消技术完成设计与实现计划" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

- [ ] **Step 2: 确认工作区干净**

Run: `git status --porcelain`
Expected: 无输出。

---

### Task 1: 纯逻辑模块（解析、计划、单元格更新）

**Files:**
- Create: `workorder_cancel_utils.py`
- Test: `tests/test_workorder_cancel_utils.py`

**Interfaces:**
- Produces:
  - `DONE = "Done"`
  - `COL_ORDER = 1`、`COL_MARKER = 2`、`COL_REASON = 3`
  - `SheetRow(row_number: int, order: str, marker: str, reason: str)`
  - `PendingOrder(order: str, rows: tuple[SheetRow, ...])`
  - `PlanResult(orders: list[PendingOrder], stats: dict[str, int], warnings: list[str])`
  - `SessionLostError(RuntimeError)`
  - `parse_rows(values: list[list[str]]) -> list[SheetRow]`
  - `build_plan(values: list[list[str]], only_orders: set[str] | None = None) -> PlanResult`
  - `parse_only_orders(raw: str | None) -> set[str] | None`
  - `format_orders(plan: PlanResult) -> str`
  - `failure_reason(timestamp: str, message: str) -> str`
  - `success_updates(pending: PendingOrder) -> list[tuple[int, int, str]]`
  - `failure_updates(pending: PendingOrder, reason: str) -> list[tuple[int, int, str]]`
- 单元格更新元组是 `(row_number, column_number, value)`；`column_number` 使用 `COL_MARKER` / `COL_REASON`。
- `stats` 的键固定为：`total_rows`、`done`、`empty_order`、`filtered`、`duplicates`、`pending_orders`、`pending_rows`。

- [ ] **Step 1: 写失败测试**

创建 `tests/test_workorder_cancel_utils.py`：

```python
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
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `python3 -m unittest tests.test_workorder_cancel_utils -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'workorder_cancel_utils'`。

- [ ] **Step 3: 实现 `workorder_cancel_utils.py`**

要求：

- `parse_rows`：逐行取 A/B/C，缺失补空串，三者都 `strip()`；A/B/C 全空的行跳过；`row_number` 从 1 开始，等于原始行号。
- `build_plan`：按原始顺序遍历 `parse_rows` 的结果。
  - `total_rows` 统计 `parse_rows` 返回的行数。
  - A 空 → `empty_order += 1`，忽略。
  - `only_orders` 不是 `None` 且订单不在集合里 → `filtered += 1`，忽略。
  - B 非空 → `done += 1`，忽略。
  - 其余按订单号分组：第一次出现创建 `PendingOrder`；再出现把行追加进 `rows`，`duplicates += 1`。
  - `pending_orders = len(orders)`，`pending_rows = sum(len(o.rows))`。
  - `only_orders` 里有、但整张表都没出现过的订单号 → 加一条 warning，文本包含该订单号。
- `parse_only_orders`：与 09 相同——`None`/空白返回 `None`；逗号和分号都接受；返回 `set[str]`。
- `format_orders`：返回 `", ".join(o.order ...)`，空计划返回 `"（无）"`。
- `failure_reason`：`f"{timestamp} | {message}"`。
- `success_updates`：对每个 row 产生 `(row_number, COL_MARKER, DONE)` 和 `(row_number, COL_REASON, "")`，顺序按 rows。
- `failure_updates`：对每个 row 只产生 `(row_number, COL_REASON, reason)`。
- `SessionLostError` 只是一个 `RuntimeError` 子类，供 Task 2/3 使用。

- [ ] **Step 4: 运行测试，确认通过**

Run: `python3 -m unittest tests.test_workorder_cancel_utils -v`
Expected: PASS（至少 10 个测试）。

- [ ] **Step 5: 提交**

```bash
git add workorder_cancel_utils.py tests/test_workorder_cancel_utils.py
git commit -m "V2026-11-03.01_新增10取消技术完成纯逻辑" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 2: 执行编排（`execute_plan`）

**Files:**
- Modify: `workorder_cancel_utils.py`
- Modify: `tests/test_workorder_cancel_utils.py`

**Interfaces:**
- Consumes: Task 1 的 `PlanResult` / `PendingOrder` / `success_updates` / `failure_updates` / `failure_reason` / `SessionLostError`。
- Produces:
  - `ActionRecord(order: str, success: bool, message: str, row_numbers: tuple[int, ...])`
  - `ExecutionResult(records: list[ActionRecord], pending_writes: list[tuple[int, int, str]], aborted: bool, exit_code: int)`
  - `execute_plan(plan: PlanResult, cancel_fn: Callable[[str], tuple[bool, str]], write_fn: Callable[[list[tuple[int, int, str]]], bool], timestamp: str) -> ExecutionResult`
- `cancel_fn` 只对单笔订单操作，返回 `(ok, message)`；会话级失败必须抛 `SessionLostError`。
- `write_fn` 返回 `True` 表示这一批单元格已写入；`False` 表示写入失败（调用方负责最后重试）。

- [ ] **Step 1: 追加失败测试**

在 `tests/test_workorder_cancel_utils.py` 顶部 import 中加上 `execute_plan`、`SessionLostError`，并追加：

```python
class ExecutePlanTests(unittest.TestCase):
    def setUp(self):
        self.plan = build_plan([["5650370", "", ""], ["5655247", "", ""]])
        self.written = []

    def test_success_writes_done(self):
        result = execute_plan(
            self.plan,
            lambda order: (True, "已保存"),
            lambda updates: self.written.extend(updates) or True,
            "T",
        )
        self.assertEqual(result.exit_code, 0)
        self.assertEqual(self.written[0], (1, 2, "Done"))
        self.assertEqual(len(result.records), 2)

    def test_failure_writes_reason_and_continues(self):
        calls = []

        def cancel(order):
            calls.append(order)
            if order == "5650370":
                return False, "已锁定"
            return True, "已保存"

        result = execute_plan(self.plan, cancel, lambda updates: True, "T")
        self.assertEqual(calls, ["5650370", "5655247"])
        self.assertEqual(result.exit_code, 1)
        self.assertFalse(result.records[0].success)
        self.assertEqual(result.records[0].message, "已锁定")

    def test_abort_stops_remaining_orders(self):
        def cancel(order):
            raise SessionLostError("会话没了")

        result = execute_plan(self.plan, cancel, lambda updates: [], "T")
        self.assertTrue(result.aborted)
        self.assertEqual(result.records, [])
        self.assertEqual(result.exit_code, 1)
        self.assertEqual(result.pending_writes, [])

    def test_failed_write_is_returned_for_retry(self):
        result = execute_plan(
            self.plan, lambda order: (True, "已保存"),
            lambda updates: False, "T",
        )
        self.assertEqual(len(result.pending_writes), 4)
        self.assertEqual(result.exit_code, 1)
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `python3 -m unittest tests.test_workorder_cancel_utils.ExecutePlanTests -v`
Expected: FAIL，`ImportError: cannot import name 'execute_plan'`。

- [ ] **Step 3: 实现 `execute_plan`**

在 `workorder_cancel_utils.py` 追加：

- 按 `plan.orders` 顺序循环：
  - `try: ok, message = cancel_fn(pending.order)`
  - `except SessionLostError:` 设 `aborted = True` 并 `break`（**当前单和剩余单都不写**）。
  - `ok=True` → `updates = success_updates(pending)`；`ok=False` → `updates = failure_updates(pending, failure_reason(timestamp, message or "SAP 未返回消息"))`。
  - 追加一条 `ActionRecord(...)`，`row_numbers` 是该订单所有行号。
  - `if not write_fn(updates): pending_writes.extend(updates)`。
- `exit_code = 1 if (aborted or any(not r.success for r in records) or pending_writes) else 0`。
- 返回 `ExecutionResult(records, pending_writes, aborted, exit_code)`。

- [ ] **Step 4: 运行全部单元测试**

Run: `python3 -m unittest discover -s tests -v`
Expected: PASS，原有 82 个测试 + 新增测试全部通过。

- [ ] **Step 5: 提交**

```bash
git add workorder_cancel_utils.py tests/test_workorder_cancel_utils.py
git commit -m "V2026-11-03.02_新增10执行编排" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 3: SAP 操作（IW32 取消技术完成）

**Files:**
- Create: `workorder_cancel_sap.py`
- Test: `tests/test_workorder_cancel_sap.py`

**Interfaces:**
- Consumes: `workorder_cancel_utils.SessionLostError`；`workorder_close_sap` 的 `read_status` / `is_success` / `wait_for_control` / `NEGATIVE_KEYWORDS`（只 import，不修改 09）。
- Produces:
  - `OKCODE = "wnd[0]/tbar[0]/okcd"`
  - `WINDOW = "wnd[0]"`
  - `ORDER_FIELD = "wnd[0]/usr/ctxtCAUFVD-AUFNR"`
  - `CANCEL_MENU = "wnd[0]/mbar/menu[0]/menu[8]/menu[2]/menu[1]"`
  - `SAVE_BUTTON = "wnd[0]/tbar[0]/btn[11]"`
  - `POPUP = "wnd[1]"`、`POPUP_STATUS = "wnd[1]/sbar"`
  - `EXTRA_NEGATIVE_KEYWORDS: tuple[str, ...]`
  - `SAVED_KEYWORDS = ("保存", "saved")`
  - `success_is_proven(message_type: str, text: str) -> bool`
  - `prepare_session(session) -> None`（尝试 `maximize()`，失败忽略）
  - `cancel_technical_completion(session, order: str) -> tuple[bool, str]`
- 会话级失败（`wnd[0]` / okcode 找不到，或 `/nIW32` 后连订单号输入框都等不到）必须抛 `SessionLostError`，不能返回失败元组。

- [ ] **Step 1: 写失败测试**

创建 `tests/test_workorder_cancel_sap.py`：

```python
import unittest
from unittest.mock import patch

import workorder_cancel_sap as sap
from workorder_cancel_utils import SessionLostError


class FakeControl:
    def __init__(self, cid, session):
        self.id = cid
        self.session = session
        self.text = ""
        self.selected = False

    def press(self):
        self.session.pressed.append(self.id)
        self.session.maybe_popup("save")

    def select(self):
        self.session.selected.append(self.id)
        self.session.maybe_popup("menu")

    def sendVKey(self, key):
        self.session.sent.append((self.id, key))
        if self.id == sap.WINDOW and key == 0:
            field = self.session.controls.get(sap.ORDER_FIELD)
            if field is not None and field.text:
                self.session.maybe_popup("enter")

    def maximize(self):
        self.session.maximized.append(self.id)


class FakeStatusBar:
    def __init__(self, message_type="", text=""):
        self.messageType = message_type
        self.text = text


class FakePopup:
    def __init__(self, message_type="", text=""):
        self.messageType = message_type
        self.text = text

    def sendVKey(self, key):
        self.session.popup = None


class FakeSession:
    def __init__(self, message_type="S", text="订单 5650370 已保存",
                 missing=(), popup_after=None, popup_type="", popup_text=""):
        self.pressed = []
        self.selected = []
        self.sent = []
        self.maximized = []
        self.message_type = message_type
        self.text = text
        self.missing = set(missing)
        self.popup_after = popup_after
        self.popup_type = popup_type
        self.popup_text = popup_text
        self.popup = None
        self.controls = {}

    def maybe_popup(self, moment):
        if self.popup_after == moment:
            self.popup = FakePopup(self.popup_type, self.popup_text)
            self.popup.session = self

    def findById(self, cid):
        if cid in self.missing:
            raise RuntimeError(f"not found: {cid}")
        if cid == "wnd[0]/sbar":
            return FakeStatusBar(self.message_type, self.text)
        if cid == sap.POPUP:
            if self.popup is None:
                raise RuntimeError("not found: wnd[1]")
            return self.popup
        if cid == sap.POPUP_STATUS:
            if self.popup is None:
                raise RuntimeError("not found: wnd[1]/sbar")
            return FakeStatusBar(self.popup.messageType, self.popup.text)
        return self.controls.setdefault(cid, FakeControl(cid, self))


class CancelTests(unittest.TestCase):
    def test_presses_recorded_sequence(self):
        session = FakeSession()
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertTrue(ok, message)
        self.assertEqual(session.controls[sap.OKCODE].text, "/nIW32")
        self.assertEqual(session.controls[sap.ORDER_FIELD].text, "5650370")
        self.assertEqual(session.selected, [sap.CANCEL_MENU])
        self.assertEqual(session.pressed, [sap.SAVE_BUTTON])
        self.assertEqual(session.sent, [(sap.WINDOW, 0), (sap.WINDOW, 0)])

    def test_error_message_fails(self):
        session = FakeSession(message_type="E", text="订单已锁定")
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertIn("订单已锁定", message)

    def test_success_without_saved_keyword_fails(self):
        session = FakeSession(message_type="S", text="处理完成")
        ok, _ = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)

    def test_extra_negative_keyword_fails(self):
        session = FakeSession(message_type="S", text="订单 5650370 未技术完成")
        ok, _ = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)

    def test_popup_error_skips_save(self):
        session = FakeSession(
            popup_after="enter", popup_type="E", popup_text="订单已锁定",
        )
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertEqual(session.pressed, [])
        self.assertIn("订单已锁定", message)

    def test_confirmation_popup_is_dismissed_then_save(self):
        session = FakeSession(popup_after="menu")
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertTrue(ok, message)
        self.assertEqual(session.pressed, [sap.SAVE_BUTTON])

    def test_missing_okcode_raises_session_lost(self):
        session = FakeSession(missing={sap.OKCODE})
        with self.assertRaises(SessionLostError):
            sap.cancel_technical_completion(session, "5650370")

    def test_order_field_timeout_raises_session_lost(self):
        session = FakeSession(missing={sap.ORDER_FIELD})
        with patch(
            "workorder_cancel_sap.wait_for_control",
            side_effect=TimeoutError("等待控件超时"),
        ):
            with self.assertRaises(SessionLostError):
                sap.cancel_technical_completion(session, "5650370")

    def test_menu_timeout_uses_status_bar_text(self):
        session = FakeSession(
            message_type="E", text="订单不存在", missing={sap.CANCEL_MENU},
        )

        def fake_wait(session, control_id, timeout=30.0, interval=0.5):
            if control_id == sap.ORDER_FIELD:
                return session.findById(control_id)
            raise TimeoutError("等待控件超时: " + control_id)

        with patch("workorder_cancel_sap.wait_for_control", side_effect=fake_wait):
            ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertEqual(message, "订单不存在")


class SuccessIsProvenTests(unittest.TestCase):
    def test_requires_saved_keyword(self):
        self.assertTrue(sap.success_is_proven("S", "订单 5650370 已保存"))
        self.assertTrue(sap.success_is_proven("W", "已保存，有警告"))
        self.assertFalse(sap.success_is_proven("S", "处理完成"))
        self.assertFalse(sap.success_is_proven("S", "订单 未技术完成"))
        self.assertFalse(sap.success_is_proven("E", "已保存"))
        self.assertFalse(sap.success_is_proven("S", "not changed"))
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `python3 -m unittest tests.test_workorder_cancel_sap -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'workorder_cancel_sap'`。

- [ ] **Step 3: 实现 `workorder_cancel_sap.py`**

模块开头：

```python
from workorder_close_sap import (
    NEGATIVE_KEYWORDS, is_success, read_status, wait_for_control,
)
from workorder_cancel_utils import SessionLostError
```

实现要点：

- `success_is_proven`：文本先转小写；命中 `NEGATIVE_KEYWORDS` 或 `EXTRA_NEGATIVE_KEYWORDS` → `False`；`is_success(message_type, text)` 为 `False` → `False`；文本为空 → `False`；必须包含 `SAVED_KEYWORDS` 中任意一个（大小写不敏感）→ `True`。
- `EXTRA_NEGATIVE_KEYWORDS = ("未技术完成", "没有技术完成", "无法取消", "未执行", "not technically completed", "not possible", "no change")`。
- `prepare_session`：`try: session.findById(WINDOW).maximize() except Exception: pass`。
- `cancel_technical_completion` 的流程：
  1. `_dismiss_stray_popup`：最多 2 次，如果 `wnd[1]` 存在就 `sendVKey(0)`。
  2. `_start_iw32`：设置 `OKCODE` 为 `/nIW32`，`WINDOW.sendVKey(0)`；任何异常都包成 `SessionLostError`。
  3. `wait_for_control(session, ORDER_FIELD, timeout=30.0)`，`TimeoutError` → `SessionLostError`；写入订单号；`WINDOW.sendVKey(0)`。
  4. 读弹窗：`E`/`A` 或文本命中否定词 → 关闭并返回 `(False, 弹窗文本或"SAP 弹窗报错")`；其他弹窗 → 关闭后继续。
  5. `wait_for_control(session, CANCEL_MENU, timeout=30.0)`；超时先 `read_status`，有文本就返回该文本，否则返回 `等待订单屏幕超时: ...`。
  6. `menu.select()`。
  7. 再读弹窗，规则同第 4 步。
  8. `session.findById(SAVE_BUTTON).press()`。
  9. 再读弹窗；`E`/`A` → 关闭并返回失败；否则记下弹窗的 `(type, text)` 并关闭。
  10. `read_status`；如果状态栏文本为空，用弹窗的 `(type, text)` 兜底。
  11. `success_is_proven` 为真 → `(True, text)`；否则 `(False, text or f"消息类型: {message_type}")`。
- 除 `SessionLostError` 外，任何异常都 `return False, f"{type(error).__name__}: {error}"`。
- 弹窗读取辅助：`wnd[1]` 不存在 → `(False, "", "")`；存在但 `wnd[1]/sbar` 读不到 → `(True, "", "")`。
- 弹窗判定：`message_type` 大写后是 `E`/`A`，或文本命中 `NEGATIVE_KEYWORDS + EXTRA_NEGATIVE_KEYWORDS` → 错误弹窗；空类型 + 空文本视为确认框。

- [ ] **Step 4: 运行测试**

Run: `python3 -m unittest tests.test_workorder_cancel_sap -v`
Expected: PASS。

再跑全部：`python3 -m unittest discover -s tests -v`，Expected: PASS。

- [ ] **Step 5: 提交**

```bash
git add workorder_cancel_sap.py tests/test_workorder_cancel_sap.py
git commit -m "V2026-11-03.03_新增10的IW32取消技术完成操作" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 4: 主程序、录制文件与 README

**Files:**
- Create: `10 - WorkOrderCancelTECO/main.py`
- Add: `10 - WorkOrderCancelTECO/取消技术完成IW32.txt`（从 `~/Downloads/取消技术完成IW32.txt` 复制）
- Modify: `README.md`

**Interfaces:**
- Consumes: Task 1–3 的全部接口；`sap_common` 的 `ensure_sap_running` / `open_session` / `close_session` / `close_sap` / `get_resource_path`。
- Produces: 可执行的 `main.py`；README 的「一次性工具」说明。

- [ ] **Step 1: 创建目录并复制录制文件**

```bash
mkdir -p "10 - WorkOrderCancelTECO"
cp ~/Downloads/取消技术完成IW32.txt "10 - WorkOrderCancelTECO/取消技术完成IW32.txt"
```

- [ ] **Step 2: 实现 `main.py`**

结构与 `09 - WorkOrderAutoClose/main.py` 一致：

- 常量：`SPREADSHEET_URL`（同 09）、`WORKSHEET = "取消技术完成"`、`SERVICE_ACCOUNT_FILE = get_resource_path("pyreadsp-b5b9c1909de6.json")`、`WRITE_ATTEMPTS = 3`。
- `parse_args`：`--dry-run`、`--only-orders`（逗号/分号分隔，复用 `parse_only_orders`）。
- `open_spreadsheet()`：`gspread.service_account(filename=...).open_by_url(...)`。
- `load_values(sh)`：`sh.worksheet(WORKSHEET).get_all_values()`。
- `print_plan(plan, only_orders)`：打印 `total_rows` / `done` / `empty_order` / `filtered` / `pending_orders` / `pending_rows`，以及 `format_orders(plan)` 和 `plan.warnings`。
- `write_updates(ws, updates, attempts=WRITE_ATTEMPTS) -> bool`：
  - `cells = [gspread.Cell(row, col, value) for row, col, value in updates]`
  - 循环 `attempts` 次调用 `ws.update_cells(cells, value_input_option="RAW")`；失败时 `time.sleep(1)` 然后 `time.sleep(2)`；最后一次仍失败打印 `❌ 写入失败: ...` 并返回 `False`。
- `run(argv=None)`：
  1. 读 Sheet、`build_plan`；失败打印 `❌ ...` 返回 `1`。
  2. `print_plan`。
  3. `--dry-run` → 打印 `dry-run：不启动 SAP、不写任何单元格。` 返回 `0`。
  4. 没有待处理订单 → 打印 `没有需要处理的工单。` 返回 `0`。
  5. `timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")`。
  6. `ensure_sap_running()` → `open_session()` → `prepare_session(session)`；失败走 `except` 返回 `1`。
  7. `ws = sh.worksheet(WORKSHEET)`。
  8. 调 `execute_plan(plan, cancel_fn, write_fn, timestamp)`；`cancel_fn` 包装 `cancel_technical_completion`，并在每单前后打印：
     ```
     [10] 5650370 ...
     [10] 5650370 成功：订单 5650370 已保存
     ```
     失败时打印 `[10] 5650370 失败：<原因>`。
  9. `execution.aborted` → 打印 `❌ SAP 会话不可用，已中止；第 N 个待处理订单起未处理。`
  10. `execution.pending_writes` 非空 → 再调用一次 `write_updates`；仍失败则打印 `❌ 仍有结果写不进 Sheet，请按上面的订单清单手工补写 B/C 列。`
  11. 打印汇总：成功单数、失败单数、`plan.stats["done"]` 行已跳过；失败清单逐行列出。
  12. 退出码：`1` 如果 `aborted`、存在失败记录、或最终写入仍失败；否则 `0`。
  13. `finally`：`close_session(session)`；`we_started` 为真才 `close_sap()`。
- 不写本地日志文件；C 列就是失败记录。

- [ ] **Step 3: 更新 `README.md`**

- 「项目清单」新增一节「不在 `run_all.ps1` 中（一次性工具）」：

  | 目录 | SAP 事务码 | 说明 |
  |---|---|---|
  | `10 - WorkOrderCancelTECO` | IW32 | 按 `取消技术完成` 工作表逐单撤销技术完成；B 列写 `Done`，C 列写失败原因 |

- 新增 `### 10` 的说明：
  - 数据源：同一 spreadsheet 的 `取消技术完成` 表；无表头，A=订单、B=`Done`、C=失败原因。
  - 运行命令：
    ```powershell
    python "10 - WorkOrderCancelTECO\main.py" --dry-run
    python "10 - WorkOrderCancelTECO\main.py" --only-orders 5650370
    python "10 - WorkOrderCancelTECO\main.py"
    ```
  - 成功/失败行为、退出码、逐单 IW32、不批量。
  - **09 关系提醒**：取消后回到 `REL CNF`；这些单的基本开始日期超过 14 天后，09 会再次技术完成它们；10 不负责阻止。
  - 不在 `run_all.ps1` 中。
- 「目录里还有什么」增加两行：
  - `workorder_cancel_utils.py` | `10` 的纯逻辑（解析、计划、执行编排），不 import `win32com` / `gspread`，单元测试可在任何机器上跑
  - `workorder_cancel_sap.py` | `10` 的 SAP GUI 操作（IW32 取消技术完成），只依赖传入的 session

- [ ] **Step 4: 语法检查与全部单元测试**

Run: `python3 -m py_compile "10 - WorkOrderCancelTECO/main.py" workorder_cancel_utils.py workorder_cancel_sap.py`
Expected: 无输出（退出码 0）。

Run: `python3 -m unittest discover -s tests -v`
Expected: PASS，全部测试通过。

- [ ] **Step 5: 提交**

```bash
git add "10 - WorkOrderCancelTECO" README.md
git commit -m "V2026-11-03.04_新增10主程序并更新README" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 5: Windows 实机验收（由用户在运行机执行）

**Files:**
- 无代码改动；只在运行机上验证。

**Interfaces:**
- Consumes: Task 4 完成并 `git pull` 到运行机。
- Produces: 真实取消结果；确认成功消息判定是否需要放宽。

- [ ] **Step 1: 干跑核对名单**

```powershell
cd "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
python "10 - WorkOrderCancelTECO\main.py" --dry-run
```

Expected：读到 59 行；已处理(Done) 1 行；待处理 58 单。与 Sheet 对照。

- [ ] **Step 2: 单笔实跑**

```powershell
python "10 - WorkOrderCancelTECO\main.py" --only-orders 5650370
```

Expected：
- SAP 里该订单从 `TECO` 变为 `REL CNF`（IW33 查看）；
- `取消技术完成!B2` = `Done`，`C2` 为空；
- 控制台的成功消息实际形态被记录下来。**如果成功消息不含「保存」/`saved`，先停下反馈，不要继续全量。**

- [ ] **Step 3: 全量运行**

```powershell
python "10 - WorkOrderCancelTECO\main.py" ; "退出码 = $LASTEXITCODE"
```

Expected：剩余 57 单逐单处理；B/C 列符合预期；退出码 `0`（或失败单对应 `1`）。

- [ ] **Step 4: 重跑验证幂等**

再跑一次同一条命令。

Expected：所有待处理行已 `Done`，脚本打印 `没有需要处理的工单。`，退出码 `0`。

- [ ] **Step 5: 失败路径（可选）**

把某张已成功取消的订单的 B 列临时清空，再跑 `--only-orders <该单>`。

Expected：SAP 报未技术完成/未更改；B 保持空；C 显示 `时间 | 原因`。
