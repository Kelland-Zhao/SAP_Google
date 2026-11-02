# 09 - 工单自动关闭 实现计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 新增 `09 - WorkOrderAutoClose`：从 `Total_Workorder` 筛出基本开始日期在 14 天前或更早、非 PM01 的工单，在 SAP 里自动做工时确认（IW48）和技术完成（IW38），结果打印到控制台并追加到 `WorkOrderClose_Log`。

**Architecture:** 把可测试的纯逻辑放在仓库根的 `workorder_close_utils.py`；把 SAP GUI 操作放在 `workorder_close_sap.py`（只依赖传入的 session，不 import win32com，便于在 macOS 上用假 session 测试）；`09 - WorkOrderAutoClose/main.py` 只负责 argparse、gspread 读写、SAP 会话生命周期和把三者串起来。这样筛选/分组/日志行可以在 macOS 上跑单元测试，SAP 操作序列也能用假 session 验证按键顺序。

**Tech Stack:** Python 3.11、`pywin32`（仅 Windows）、`gspread` + `google-auth`、`unittest`。不使用 Excel/pandas。

**Spec:** `docs/superpowers/specs/2026-11-02-workorder-close-design.md`

## Global Constraints

- 运行机：Windows + SAP GUI Scripting + Python 3.11；开发/测试可在 macOS 上跑纯逻辑单元测试。
- 依赖：`pywin32`、`gspread`、`google-auth`；不新增其他依赖。
- Google 服务账号文件：仓库根 `pyreadsp-b5b9c1909de6.json`。
- Spreadsheet：`1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98`；数据表 `Total_Workorder`，日志表 `WorkOrderClose_Log`。
- 日志表固定 9 列：执行时间、运行日期、订单、订单类型、基本开始日期、处理前系统状态、动作、结果、原因 / SAP 消息。
- 日志只追加；日志写入失败只打印警告，不改变退出码。
- 退出码：`0` 成功/无事可做；`1` 读 Sheet 失败、SAP 失败、IW48 批量工时确认失败、或 IW38 批量技术完成失败。
- 09 依赖 04 刚刷新过 `Total_Workorder`；单独运行 09 前必须先跑 04，启动时打印提醒。
- 控制台中文输出；订单号按文本处理，保留前导零。
- 两个 SAP 录制文件复制进 `09 - WorkOrderAutoClose/` 作为参考，不执行。
- 按仓库现有习惯直接提交到 `main`（用户的工作流是本地 `main` → push → 运行机 pull）。

## Review Focus

按最可能出问题的顺序：

1. **IW38 批量技术完成前是否真的选中了订单**：`btn[5]`/`btn[45]` 中一个是全选。如果按键顺序或选中状态不对，SAP 可能返回成功但一单没关。Task 2 的 `CompleteTechnicallyTests.test_uploads_orders_and_presses_recorded_sequence` 锁定按键顺序；Task 5 要求 Windows 上先做单量实测。
2. **IW48 批量提交的结果粒度**：整批只有一个状态栏消息；一张订单有问题可能让整批失败。`btn[11]` 是保存，保存后立即读状态栏；空消息/错误消息都按整批失败处理，组 A 全部不进入 TECO，下一次 04 刷新后按实际状态重试。Task 2 的 `ConfirmTimesTests.test_*` 覆盖成功、错误、空消息三种情况。
3. **重复订单号/状态冲突**：同一订单出现多行且状态不同，可能重复确认或选错状态。Task 1 的 `BuildPlanTests.test_deduplicates_and_warns_on_conflict` 锁定「保留第一条 + 警告」。
4. **基本开始日期的格式**：K 列可能是 `YYYY-MM-DD`、`YYYYMMDD`、斜杠格式或空值。解析错误会把工单放进/漏出「14 天前或更早」的筛选。Task 1 的 `test_parse_sheet_date_*` 覆盖。
5. **系统状态的双空格与中文状态**：状态串里有连续空格和中文（如 `结算`），必须按整词匹配。Task 1 的 `test_parse_status_*` 和 `test_classify_status_*` 覆盖。

---

### Task 0: 提交设计与计划文档

**Files:**
- Add: `docs/superpowers/specs/2026-11-02-workorder-close-design.md`
- Add: `docs/superpowers/plans/2026-11-02-workorder-close.md`

**Interfaces:**
- Consumes: 已批准的设计文档。
- Produces: 仓库里可追溯的设计与计划；后续任务从干净的工作区开始。

- [ ] **Step 1: 提交文档**

```bash
git add docs/superpowers/specs/2026-11-02-workorder-close-design.md docs/superpowers/plans/2026-11-02-workorder-close.md
git commit -m "V2026-11-02.00_新增09工单自动关闭设计与实现计划" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

- [ ] **Step 2: 确认工作区干净**

Run: `git status --porcelain`
Expected: 无输出。

---

### Task 1: 纯逻辑模块（筛选、分组、执行编排、日志行）

**Files:**
- Create: `workorder_close_utils.py`
- Test: `tests/test_workorder_close_utils.py`

**Interfaces:**
- Produces:
  - `WINDOW_DAYS = 14`
  - `REQUIRED_COLUMNS = ["订单", "系统状态", "订单类型", "基本开始日期"]`
  - `LOG_HEADERS: list[str]`（9 列，顺序见 Global Constraints）
  - `OrderRow(order: str, order_type: str, start_date_raw: str, status_raw: str)`
  - `OrderDecision(order: str, order_type: str, start_date_raw: str, status_raw: str, action: str, reason: str)`，`action` ∈ `{"CNF","TECO","SKIP"}`
  - `PlanResult(cnf_orders: list[OrderDecision], teco_orders: list[OrderDecision], skipped: list[OrderDecision], stats: dict[str,int], warnings: list[str])`
  - `ActionRecord(decision: OrderDecision, action: str, success: bool, message: str)`，`action` ∈ `{"工时确认","技术完成"}`
  - `ExecutionResult(records: list[ActionRecord], teco_submitted: list[str], exit_code: int)`
  - `parse_status(raw) -> set[str]`
  - `parse_sheet_date(raw) -> datetime.date | None`
  - `is_due_for_close(start: datetime.date, run_date: datetime.date, window_days: int = WINDOW_DAYS) -> bool`
  - `classify_status(statuses: set[str]) -> tuple[str, str]`
  - `parse_only_orders(raw: str | None) -> set[str] | None`
  - `rows_from_values(values: list[list[str]]) -> list[OrderRow]`
  - `build_plan(values, run_date, window_days=WINDOW_DAYS, only_orders=None) -> PlanResult`
  - `execute_plan(plan, confirm_fn: Callable[[list[str]], tuple[bool, str]], teco_fn: Callable[[list[str]], tuple[bool, str]]) -> ExecutionResult`
  - `build_log_rows(plan, execution, run_date, timestamp) -> list[list[str]]`
- Consumes: `gap_utils.resolve_columns` / `MissingColumnsError`（复用表头定位，不重复写）。

- [ ] **Step 1: 写第一批失败测试（状态/日期/窗口/分类/参数）**

创建 `tests/test_workorder_close_utils.py`，至少包含：

```python
import datetime
import unittest

from workorder_close_utils import (
    classify_status, is_due_for_close, parse_only_orders, parse_sheet_date,
    parse_status,
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

    def test_teco_and_cnf_is_closed(self):
        self.assertEqual(
            classify_status({"TECO", "CNF"}), ("SKIP", "已关闭(TECO+CNF)")
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
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `python -m unittest tests.test_workorder_close_utils -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'workorder_close_utils'`。

- [ ] **Step 3: 实现第一批函数**

在 `workorder_close_utils.py` 中实现 dataclasses 与 `parse_status` / `parse_sheet_date` / `is_due_for_close` / `classify_status` / `parse_only_orders`。

要求：
- `parse_status` 用 `str(raw).split()`，每个 token 做 `.upper()`；`None`/空串返回 `set()`。
- `parse_sheet_date` 支持 `%Y-%m-%d`、`%Y/%m/%d`、`%Y.%m.%d`、`%Y%m%d`、`%m/%d/%Y`、`%d.%m.%Y`；直接传 `datetime.date`/`datetime.datetime` 时返回 `date`；其余返回 `None`。
- `is_due_for_close` 的判定是 `start <= run_date - timedelta(days=window_days)`（含边界，没有下限）。
- `classify_status` 按此顺序：`TECO+CNF` → 已关闭；`REL & !CNF & !TECO` → `CNF`；`REL & CNF & !TECO` → `TECO`；`TECO & !CNF` → 异常；`CNF & !REL & !TECO` → 异常；`REL & TECO` → 异常；其余 `未释放`。
- `parse_only_orders` 同时接受逗号和分号，返回 `set[str]` 或 `None`。

- [ ] **Step 4: 运行测试，确认通过**

Run: `python -m unittest tests.test_workorder_close_utils -v`
Expected: PASS。

- [ ] **Step 5: 写第二批失败测试（表头映射、build_plan、execute_plan、日志行）**

在同一个测试文件追加；同时把 `rows_from_values, OrderDecision, PlanResult, ActionRecord, ExecutionResult` 加进文件顶部的 import：

```python
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
                         {"M1": "已关闭(TECO+CNF)", "T1": "TECO无CNF"})
        self.assertEqual(plan.stats["pm01"], 1)
        self.assertEqual(plan.stats["too_recent"], 1)
        self.assertEqual(plan.stats["invalid"], 1)
        self.assertEqual(plan.stats["candidates"], 4)
        self.assertEqual(plan.stats["closed"], 1)
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
            [OrderDecision("S1", "PM02", "2026-10-05", "TECO", "SKIP", "TECO无CNF")],
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
            "2026-10-05", "TECO", "跳过", "跳过", "TECO无CNF",
        ])
        self.assertEqual(rows[1], [
            "2026-11-02 07:00:00", "2026-11-02", "A1", "PM02",
            "2026-10-05", "REL", "工时确认", "成功", "已保存",
        ])
        self.assertTrue(all(len(row) == 9 for row in rows))
```

- [ ] **Step 6: 运行测试，确认失败**

Run: `python -m unittest tests.test_workorder_close_utils -v`
Expected: FAIL，`ImportError` 或 `AttributeError`（`rows_from_values` / `build_plan` 等未定义）。

- [ ] **Step 7: 实现第二批函数**

实现 `rows_from_values`、`build_plan`、`execute_plan`、`build_log_rows`，并补齐 dataclasses / `LOG_HEADERS`。

要求：
- `rows_from_values` 复用 `gap_utils.resolve_columns(values[0], REQUIRED_COLUMNS)`；`values` 为空时抛 `ValueError("工作表为空")`；缺列时让 `MissingColumnsError` 冒泡；所有单元格都为空的行直接跳过，不产生 `OrderRow`。
- `build_plan` 的 `stats` 至少包含 `total_rows`、`pm01`、`too_recent`、`invalid`、`duplicates`、`candidates`、`closed`、`anomalies`；统计在 `only_orders` 过滤前算好。
- 只有通过 J/K 初筛（非空订单、非空订单类型、非 PM01、K 可解析且在窗口内）的行才进入 `cnf_orders` / `teco_orders` / `skipped`；其余只计数。
- 重复订单保留第一条，后续计入 `duplicates`；状态不同时往 `warnings` 追加一条含订单号的警告。
- `execute_plan` 的顺序：先对全部组 A 调用一次 `confirm_fn(orders)`；为每个组 A 订单生成一条「工时确认」记录（共用批次结果/消息）；成功则把所有组 A 加入 TECO 名单，失败则都不加入；然后对「组 B + 成功组 A」调用一次 `teco_fn`，为每个 TECO 订单生成一条「技术完成」记录；`teco_fn` 失败时所有 TECO 记录都记失败。
- `build_log_rows` 只输出 `plan.skipped` 里的异常（不在 `NORMAL_SKIP_REASONS` 里），再按 `execution.records` 顺序输出；每行正好 9 列。`NORMAL_SKIP_REASONS = {"已关闭(TECO+CNF)", "未释放"}`。

- [ ] **Step 8: 运行测试，确认通过**

Run: `python -m unittest tests.test_workorder_close_utils -v`
Expected: PASS。

- [ ] **Step 9: 提交**

```bash
git add workorder_close_utils.py tests/test_workorder_close_utils.py
git commit -m "V2026-11-02.01_新增09工单筛选与执行编排逻辑" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 2: SAP GUI 操作模块（IW48 工时确认、IW38 批量技术完成）

**Files:**
- Create: `workorder_close_sap.py`
- Test: `tests/test_workorder_close_sap.py`

**Interfaces:**
- Produces:
  - `read_status(session) -> tuple[str, str]`：返回 `(message_type, text)`，读不到时返回 `("", "")`。
  - `is_success(message_type: str, text: str) -> bool`：`S`/`W` 成功；`E`/`A`、空、未知都失败；文本含「未选择 / 未选中 / 没有选择 / not selected」也失败。
  - `confirm_times(session, orders: list[str]) -> tuple[bool, str]`
  - `complete_technically(session, orders: list[str]) -> tuple[bool, str]`
  - `copy_to_clipboard(text: str) -> None`（内部 lazy import `win32clipboard`）
- Consumes: 只使用传入的 SAP GUI `session` 对象；不 import `win32com` / `sap_common`，保证 macOS 上可 import 并测试。

- [ ] **Step 1: 写失败的假 session 测试**

创建 `tests/test_workorder_close_sap.py`，至少包含：

```python
import unittest
from unittest.mock import patch

import workorder_close_sap as sap


class FakeControl:
    def __init__(self, cid, session):
        self.id = cid
        self.session = session
        self.text = ""
        self.selected = False

    def press(self):
        self.session.pressed.append(self.id)

    def sendVKey(self, key):
        self.session.sent.append((self.id, key))

    def setFocus(self):
        pass

    def resizeWorkingPane(self, *args):
        self.session.resized.append(self.id)


class FakeStatusBar:
    def __init__(self, message_type, text):
        self.messageType = message_type
        self.text = text


class FakeSession:
    def __init__(self, message_type="S", text="已保存", missing=()):
        self.pressed = []
        self.sent = []
        self.resized = []
        self.controls = {}
        self.message_type = message_type
        self.text = text
        self.missing = set(missing)

    def findById(self, cid):
        if cid in self.missing:
            raise RuntimeError(f"not found: {cid}")
        if cid == "wnd[0]/sbar":
            return FakeStatusBar(self.message_type, self.text)
        return self.controls.setdefault(cid, FakeControl(cid, self))


class ConfirmTimesTests(unittest.TestCase):
    def test_uploads_orders_and_presses_recorded_sequence(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard") as clipboard:
            ok, message = sap.confirm_times(session, ["5609259", "5611111"])
        self.assertTrue(ok)
        clipboard.assert_called_once_with("5609259\r\n5611111")
        for cid in (
            "wnd[0]/tbar[1]/btn[19]",
            "wnd[1]/tbar[0]/btn[24]",
            "wnd[0]/tbar[1]/btn[8]",
            "wnd[0]/tbar[1]/btn[9]",
            "wnd[0]/tbar[1]/btn[38]",
            "wnd[0]/tbar[0]/btn[11]",
        ):
            self.assertIn(cid, session.pressed)
        for cid in (
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-AUERU[9,0]",
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-AUSOR[10,0]",
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-LEKNW[11,0]",
        ):
            self.assertTrue(session.controls[cid].selected)

    def test_no_orders_does_not_touch_sap(self):
        session = FakeSession()
        ok, message = sap.confirm_times(session, [])
        self.assertTrue(ok)
        self.assertEqual(session.pressed, [])

    def test_error_message_returns_false(self):
        session = FakeSession(message_type="E", text="订单已锁定")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("订单已锁定", message)

    def test_empty_message_returns_false(self):
        session = FakeSession(message_type="", text="")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, _ = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)

    def test_missing_confirmation_table_returns_false(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"), \
             patch("workorder_close_sap.wait_for_control",
                   side_effect=TimeoutError("等待控件超时")):
            ok, message = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("TimeoutError", message)


class CompleteTechnicallyTests(unittest.TestCase):
    def test_uploads_orders_and_presses_recorded_sequence(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard") as clipboard:
            ok, message = sap.complete_technically(session, ["5609259", "5611111"])
        self.assertTrue(ok)
        clipboard.assert_called_once_with("5609259\r\n5611111")
        for cid in (
            "wnd[1]/tbar[0]/btn[24]",
            "wnd[1]/tbar[0]/btn[8]",
            "wnd[0]/tbar[1]/btn[8]",
            "wnd[0]/tbar[1]/btn[5]",
            "wnd[0]/tbar[1]/btn[45]",
            "wnd[1]/usr/btnSPOP-VAROPTION1",
        ):
            self.assertIn(cid, session.pressed)
        self.assertEqual(session.controls["wnd[0]/usr/ctxtVARIANT"].text, "/KEL")

    def test_no_orders_does_not_touch_sap(self):
        session = FakeSession()
        ok, message = sap.complete_technically(session, [])
        self.assertTrue(ok)
        self.assertEqual(session.pressed, [])

    def test_error_message_returns_false(self):
        session = FakeSession(message_type="E", text="未选择订单")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.complete_technically(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("未选择订单", message)

    def test_warning_that_says_no_orders_selected_is_failure(self):
        session = FakeSession(message_type="W", text="未选择订单")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.complete_technically(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("未选择订单", message)


class IsSuccessTests(unittest.TestCase):
    def test_success_and_warning_are_success(self):
        self.assertTrue(sap.is_success("S", "已保存"))
        self.assertTrue(sap.is_success("W", "警告"))

    def test_error_abort_empty_and_unknown_are_failure(self):
        for message_type in ("E", "A", "", "X"):
            self.assertFalse(sap.is_success(message_type, "消息"))

    def test_warning_with_negative_text_is_failure(self):
        self.assertFalse(sap.is_success("W", "未选择订单"))
        self.assertFalse(sap.is_success("S", "not selected"))
```

- [ ] **Step 2: 运行测试，确认失败**

Run: `python -m unittest tests.test_workorder_close_sap -v`
Expected: FAIL，`ModuleNotFoundError: No module named 'workorder_close_sap'`。

- [ ] **Step 3: 实现 `workorder_close_sap.py`**

只 import `time`；`win32clipboard` 在 `copy_to_clipboard()` 内部 import。实现：

- `read_status(session)`：读 `wnd[0]/sbar` 的 `.text` 和 `.messageType`；兼容 `messageType` / `MessageType`；异常返回 `("", "")`。
- `is_success(message_type, text)`：先检查文本是否含 `未选择` / `未选中` / `没有选择` / `not selected` / `no orders`，含则失败；否则 `message_type.upper() in {"S","W"}` 才算成功。
- `wait_for_control(session, control_id, timeout=30.0, interval=0.5)`：轮询 `findById`，超时抛 `TimeoutError`。
- `confirm_times(session, orders)`：
  1. 空列表直接返回 `(True, "没有需要工时确认的工单")`
  2. `session.findById("wnd[0]").resizeWorkingPane(232, 29, False)`
  3. `wnd[0]/tbar[0]/okcd` 填 `/nIW48`，`wnd[0]/sendVKey(0)`
  4. 按 `wnd[0]/tbar[1]/btn[19]`
  5. 依次清空 `%_ARBPL_%`、`%_WERKS_%`、`%_AUART_%`：按 VALU_PUSH → 弹窗 `tbar[0]/btn[16]` → `tbar[0]/btn[8]`
  6. 打开 `%_AUFNR_%`，先 btn[16] 清空；`copy_to_clipboard("\r\n".join(orders))`，按 btn[24]（自剪贴板上载），按 btn[8]
  7. 按 `wnd[0]/tbar[1]/btn[8]` 执行
  8. `wait_for_control` 等 `chkAFRD-AUERU[9,0]`，然后把三个复选框 `.selected = True`
  9. 依次按 `wnd[0]/tbar[1]/btn[9]`、`wnd[0]/tbar[1]/btn[38]`、`wnd[0]/tbar[0]/btn[11]`
  10. `read_status`，`is_success` 判定；异常捕获后返回 `(False, f"{type(e).__name__}: {e}")`
- `complete_technically(session, orders)`：
  1. 空列表直接返回 `(True, "没有需要技术完成的工单")`
  2. resize；`/nIW38`
  3. 打开 `%_AUFNR_%`，btn[16] 清空
  4. `copy_to_clipboard("\r\n".join(orders))`，按 btn[24]（自剪贴板上载），按 btn[8]
  5. `wnd[0]/usr/ctxtVARIANT` = `/KEL`
  6. 按 `wnd[0]/tbar[1]/btn[8]` 执行
  7. `time.sleep(3)` 等列表
  8. 按 `wnd[0]/tbar[1]/btn[5]`，再按 `wnd[0]/tbar[1]/btn[45]`
  9. 按 `wnd[1]/usr/btnSPOP-VAROPTION1`
  10. `read_status` 判定；异常捕获返回 `(False, ...)`

- [ ] **Step 4: 运行测试，确认通过**

Run: `python -m unittest tests.test_workorder_close_sap -v`
Expected: PASS。

- [ ] **Step 5: 提交**

```bash
git add workorder_close_sap.py tests/test_workorder_close_sap.py
git commit -m "V2026-11-02.02_新增09的SAP工时确认与技术完成操作" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 3: 主程序（argparse、gspread、SAP 会话、控制台输出）

**Files:**
- Create: `09 - WorkOrderAutoClose/main.py`

**Interfaces:**
- Consumes:
  - `workorder_close_utils`: `WINDOW_DAYS`、`LOG_HEADERS`、`parse_only_orders`、`build_plan`、`execute_plan`、`build_log_rows`、`PlanResult`、`ExecutionResult`
  - `workorder_close_sap`: `confirm_times(session, orders)`、`complete_technically(session, orders)`
  - `sap_common`: `get_resource_path`、`ensure_sap_running`、`open_session`、`close_session`、`close_sap`
- Produces: 可执行入口 `main.py`，支持 `--dry-run` 和 `--only-orders`。

- [ ] **Step 1: 写 `main.py`**

要点：

- 顶部沿用现有项目写法：
  ```python
  sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
  from sap_common import (get_resource_path, ensure_sap_running,
                          open_session, close_session, close_sap)
  from workorder_close_utils import (...)
  from workorder_close_sap import confirm_times, complete_technically
  ```
- 常量：
  ```python
  GOOGLE_SHEET_URL = "https://docs.google.com/spreadsheets/d/1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98/edit"
  TOTAL_WORKSHEET = "Total_Workorder"
  LOG_WORKSHEET = "WorkOrderClose_Log"
  SERVICE_ACCOUNT_FILE = get_resource_path("pyreadsp-b5b9c1909de6.json")
  ```
- `parse_args(argv=None)`：`--dry-run`（store_true）、`--only-orders`（字符串，逗号/分号分隔）。
- `open_spreadsheet()`：`gspread.service_account(filename=SERVICE_ACCOUNT_FILE).open_by_url(GOOGLE_SHEET_URL)`。
- `load_values(sh)`：`sh.worksheet(TOTAL_WORKSHEET).get_all_values()`。
- `get_or_create_log_sheet(sh)`：找不到就 `add_worksheet(title=LOG_WORKSHEET, rows=1000, cols=len(LOG_HEADERS))`。
- `ensure_log_header(ws)`：`ws.row_values(1)` 全空时 `append_row(LOG_HEADERS, value_input_option="RAW")`；非空但与 `LOG_HEADERS` 不一致时打印警告，不覆盖。
- `append_log_rows(sh, rows)`：`ws.append_rows(rows, value_input_option="RAW")`；任何异常只 `print("⚠️ 日志写入失败: ...")`。
- `print_plan(plan, run_date)`：打印筛选条件、统计、`plan.warnings`、组 A/组 B 数量、已关闭数量；跳过按原因汇总，只有不在 `NORMAL_SKIP_REASONS` 里的异常才逐单打印。
- `print_execution(execution)`：先打印 `[CNF] 批量提交 N 单: ...` 和批次消息，再逐条打印 `[CNF] 订单 成功/失败（整批结果）`；TECO 同样先打印 `[TECO] 批量提交 N 单: ...`，再逐条打印。
- `run(argv=None)`：
  1. 打印提醒：`⚠️ 注意：09 依赖 Total_Workorder 是 04 刚刷新过的；单独运行前请先跑 04。`
  2. `args = parse_args(argv)`；`run_date = datetime.date.today()`；`only_orders = parse_only_orders(args.only_orders)`
  3. 读 Sheet；失败打印错误并 `return 1`
  4. `plan = build_plan(values, run_date, only_orders=only_orders)`；打印
  5. `--dry-run` → 打印「dry-run，不启动 SAP、不写日志」并 `return 0`
  6. 如果 `not plan.cnf_orders and not plan.teco_orders`：写 skip 日志（`build_log_rows(plan, ExecutionResult([], [], 0), ...)`），打印「没有需要处理的工单」，`return 0`
  7. `we_started = ensure_sap_running()`；`session = open_session()`；`try/finally` 保证 `close_session(session)` 和必要时 `close_sap()`
  8. `execution = execute_plan(plan, lambda orders: confirm_times(session, orders), lambda orders: complete_technically(session, orders))`
  9. `print_execution(execution)`
  10. `rows = build_log_rows(plan, execution, run_date, datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))`；`append_log_rows(sh, rows)`
  11. `return execution.exit_code`
  - SAP 启动/会话异常：打印错误并 `return 1`。
  - `--only-orders` 指定的订单不在计划里：打印警告，继续。
- `if __name__ == "__main__": sys.exit(run())`

- [ ] **Step 2: 编译检查**

Run: `python -m py_compile "09 - WorkOrderAutoClose/main.py"`
Expected: 无输出，退出码 0。

- [ ] **Step 3: 检查导入边界**

Run: `python -m py_compile workorder_close_utils.py workorder_close_sap.py`
Expected: 无输出。
确认 `workorder_close_sap.py` 顶部没有 `import win32com` / `import win32clipboard`（后者只能在 `copy_to_clipboard` 内部）。

- [ ] **Step 4: 提交**

```bash
git add "09 - WorkOrderAutoClose/main.py"
git commit -m "V2026-11-02.03_新增09主程序与Google日志写入" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 4: 接入 run_all、更新 README、复制录制脚本

**Files:**
- Modify: `run_all.ps1:13-21`（`$Projects` 数组）、`run_all.ps1:30`（标题）、`run_all.ps1:51`（7 个脚本文案）
- Modify: `README.md`（运行方式、项目清单、日志表、目录说明）
- Create: `09 - WorkOrderAutoClose/工时确认IW48.txt`（复制自 `~/Downloads/工时确认IW48.txt`）
- Create: `09 - WorkOrderAutoClose/PM工单技术完成.txt`（复制自 `~/Downloads/PM工单技术完成.txt`）

**Interfaces:**
- Consumes: Task 3 的可执行 `main.py`。
- Produces: `run_all.ps1` 中 09 的执行步骤；README 与代码一致的说明。

- [ ] **Step 1: 修改 `run_all.ps1`**

把 `$Projects` 改为：

```powershell
$Projects = @(
    "00 - Maintenance_Plan_Adherence_Sync",
    "01 - Maintenance_Effectiveness_Sync",
    "02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync",
    "03 - Safety_Stock_ZSE16",
    "04 - Total_Workorder",
    "09 - WorkOrderAutoClose",
    "04 - Total_Workorder",
    "06 - IM_Equipment_Number",
    "07 - Inventory_MB52"
)
```

同时把开头 `共 $Total 个项目` 改为 `共 $Total 个步骤`，把 `不终止的话 7 个脚本会各等 60 秒` 改成 `8 个脚本`。

在循环里加一道保护：如果第一次 `04` 失败（`$FailedList` 含 `04 - Total_Workorder`），跳过 `09` 并把 `09` 记为失败；第二次 `04` 仍执行。

- [ ] **Step 2: 更新 README**

至少改这些位置：

- 第 5 行「依次执行 7 个独立脚本」→ 8 个独立脚本 / 9 个步骤。
- 「运行方式 → 全量运行（7 个项目）」→ 8 个项目 / 9 个步骤。
- 执行顺序改为：`00 → 01 → 02 → 03 → 04 → 09 → 04 → 06 → 07`，并解释：第一次 04 刷新 Sheet，09 关单，第二次 04 把新状态刷回 Sheet。
- 项目清单表增加：
  `| 09 - WorkOrderAutoClose | IW48 + IW38 | 工时确认 + 技术完成 | WorkOrderClose_Log |`
- 在项目清单后加一段 09 的说明：筛选规则（非 PM01、K 列 14 天前或更早）、组 A/组 B、批量 TECO、日志 9 列、退出码、`REL` 与 `CNF` 可共存、单独运行前必须先跑 04。
- 「环境要求」的依赖说明里补充 09 不需要 Excel/pandas。
- 如果第一次 04 失败，run_all 会跳过 09（避免用旧 Sheet 关单），第二次 04 仍执行。
- 记录已知限制：09 现在取 `K <= 今天-14天`（没有下限），上一年度及更早的行都可能被选中，而当前 04 只刷新当年数据；后续把 04 改为动态 365 天后解决。
- 「目录里还有什么」表增加 `workorder_close_utils.py` 和 `workorder_close_sap.py` 两行，说明前者是纯逻辑、后者是 SAP 操作。

- [ ] **Step 3: 复制录制脚本**

```bash
cp ~/Downloads/工时确认IW48.txt "09 - WorkOrderAutoClose/工时确认IW48.txt"
cp ~/Downloads/PM工单技术完成.txt "09 - WorkOrderAutoClose/PM工单技术完成.txt"
```

- [ ] **Step 4: 检查改动**

Run: `git diff -- run_all.ps1 README.md`
Expected: 只包含上述改动；`git status` 显示两个新录制文件。

- [ ] **Step 5: 提交**

```bash
git add run_all.ps1 README.md "09 - WorkOrderAutoClose/工时确认IW48.txt" "09 - WorkOrderAutoClose/PM工单技术完成.txt"
git commit -m "V2026-11-02.04_接入09并更新README" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

---

### Task 5: 全量验证与 Windows 实测交接

**Files:**
- 不新增文件；只运行验证命令并记录结果。

**Interfaces:**
- Consumes: Task 1-4 的全部产物。
- Produces: 可交给用户在 Windows 上实测的说明。

- [ ] **Step 1: 跑全部单元测试**

Run: `python -m unittest discover -s tests -v`
Expected: 全部 PASS，包括原有的 `test_gap_utils` / `test_month_utils` 和新增的 09 测试。

- [ ] **Step 2: 编译所有新增 Python 文件**

Run: `python -m py_compile workorder_close_utils.py workorder_close_sap.py "09 - WorkOrderAutoClose/main.py"`
Expected: 无输出，退出码 0。

- [ ] **Step 3: 检查工作区**

Run: `git status --porcelain`
Expected: 干净（除非有意保留的未提交文件）。

- [ ] **Step 4: 把 Windows 实测步骤交给用户**

在回复里给出：

```powershell
cd "C:\Users\kelland zhao\Projects\SAP_Google_AutoRun"
git pull

# 1) 干跑：只看名单，不碰 SAP、不写日志
python "09 - WorkOrderAutoClose\main.py" --dry-run

# 2) 干跑 + 指定订单：从第 1 步的组 A / 组 B 名单里挑 1–2 张
python "09 - WorkOrderAutoClose\main.py" --dry-run --only-orders <订单1>,<订单2>

# 3) 单量实测：批量确认组 A + 批量技术完成组 B
python "09 - WorkOrderAutoClose\main.py" --only-orders <订单1>,<订单2> ; "退出码 = $LASTEXITCODE"

# 4) 通过后再跑完整 run_all（会依次 04 → 09 → 04）
.\run_all.ps1
```

并说明：如果第 3 步失败，把控制台里的 `[CNF]` / `[TECO]` 行和完整 traceback 发回来；不要直接跑第 4 步。

- [ ] **Step 5: 最终提交（如有验证中产生的改动）**

如果验证过程改了代码或文档，把实际改动的文件逐个 `git add`（不要用 `-A`）：

```bash
git add <改动的文件...>
git commit -m "V2026-11-02.05_09实现验证修正" -m "Co-Authored-By: Claude Code <noreply@anthropic.com>"
```

如果没有改动，跳过这一步。
