"""09 - 工单自动关闭。

从 Total_Workorder 筛出基本开始日期在 14 天前或更早、非 PM01 的工单：
  * 组 A（REL 且无 CNF/TECO）→ IW48 批量工时确认；
  * 组 B（REL+CNF 且无 TECO）→ IW38 批量技术完成；
  * 组 A 确认成功后，也并入同一轮 IW38。

结果打印到控制台，并追加到 WorkOrderClose_Log。
"""

import argparse
import datetime
import os
import sys

import gspread

# 公共部分在仓库根，不在本脚本所在目录。
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sap_common import (get_resource_path, ensure_sap_running,
                        open_session, close_session, close_sap)
from workorder_close_utils import (
    ExecutionResult,
    LOG_HEADERS,
    NORMAL_SKIP_REASONS,
    WINDOW_DAYS,
    build_log_rows,
    build_plan,
    execute_plan,
    format_orders,
    parse_only_orders,
)
from workorder_close_sap import confirm_times, complete_technically

GOOGLE_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98/edit"
)
TOTAL_WORKSHEET = "Total_Workorder"
LOG_WORKSHEET = "WorkOrderClose_Log"
SERVICE_ACCOUNT_FILE = get_resource_path("pyreadsp-b5b9c1909de6.json")


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="09 - 工单自动关闭")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="只读 Sheet 并打印计划，不启动 SAP、不写日志",
    )
    parser.add_argument(
        "--only-orders", default=None,
        help="只处理指定订单号，逗号/分号分隔（测试用）",
    )
    return parser.parse_args(argv)


def open_spreadsheet():
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    return gc.open_by_url(GOOGLE_SHEET_URL)


def load_values(sh):
    return sh.worksheet(TOTAL_WORKSHEET).get_all_values()


def get_or_create_log_sheet(sh):
    try:
        return sh.worksheet(LOG_WORKSHEET)
    except gspread.exceptions.WorksheetNotFound:
        return sh.add_worksheet(
            title=LOG_WORKSHEET, rows=1000, cols=len(LOG_HEADERS),
        )


def ensure_log_header(ws):
    first_row = ws.row_values(1)
    if not any(str(cell).strip() for cell in first_row):
        ws.append_row(LOG_HEADERS, value_input_option="RAW")
        return

    existing = [str(cell).strip() for cell in first_row[:len(LOG_HEADERS)]]
    if existing != LOG_HEADERS:
        print("⚠️ WorkOrderClose_Log 的表头与预期不一致，仍按现有表头追加。")


def append_log_rows(sh, rows):
    if not rows:
        return
    try:
        ws = get_or_create_log_sheet(sh)
        ensure_log_header(ws)
        ws.append_rows(rows, value_input_option="RAW")
        print(f"✅ 已写入 WorkOrderClose_Log {len(rows)} 行")
    except Exception as error:
        print(f"⚠️ 日志写入失败: {type(error).__name__}: {error}")


def print_plan(plan, run_date):
    cutoff = run_date - datetime.timedelta(days=WINDOW_DAYS)
    stats = plan.stats

    print(f"筛选条件: 基本开始日期 <= {cutoff}（今天往前 {WINDOW_DAYS} 天或更早）")
    print(
        f"读到 {stats.get('total_rows', 0)} 行；"
        f"排除 PM01 {stats.get('pm01', 0)} 行；"
        f"未满 14 天 {stats.get('too_recent', 0)} 行；"
        f"无效 {stats.get('invalid', 0)} 行；"
        f"重复 {stats.get('duplicates', 0)} 行；"
        f"候选 {stats.get('candidates', 0)} 单"
    )
    print(f"组 A（待工时确认）: {len(plan.cnf_orders)} 单")
    print(f"[CNF] 名单: {format_orders(plan.cnf_orders)}")
    print(f"组 B（待技术完成）: {len(plan.teco_orders)} 单")
    print(f"[TECO] 名单: {format_orders(plan.teco_orders)}")
    print(f"已关闭（TECO+CNF）: {stats.get('closed', 0)} 单")
    print(f"异常/其他跳过: {stats.get('anomalies', 0)} 单")

    for warning in plan.warnings:
        print(f"⚠️ {warning}")

    skip_counts = {}
    for decision in plan.skipped:
        skip_counts[decision.reason] = skip_counts.get(decision.reason, 0) + 1
    for reason, count in sorted(skip_counts.items()):
        print(f"[SKIP] {reason}: {count} 单")

    for decision in plan.skipped:
        if decision.reason not in NORMAL_SKIP_REASONS:
            print(f"[SKIP] {decision.order} {decision.reason}")


def print_execution(execution):
    cnf_records = [r for r in execution.records if r.action == "工时确认"]
    if cnf_records:
        orders = [r.decision.order for r in cnf_records]
        print(f"[CNF] 批量提交 {len(orders)} 单: {', '.join(orders)}")
        print(f"[CNF] SAP 返回: {cnf_records[0].message}")
        for record in cnf_records:
            status = "成功" if record.success else "失败"
            print(f"[CNF] {record.decision.order} {status}（整批结果）")

    teco_records = [r for r in execution.records if r.action == "技术完成"]
    if teco_records:
        orders = [r.decision.order for r in teco_records]
        print(f"[TECO] 批量提交 {len(orders)} 单: {', '.join(orders)}")
        print(f"[TECO] SAP 返回: {teco_records[0].message}")
        for record in teco_records:
            status = "成功" if record.success else "失败"
            print(f"[TECO] {record.decision.order} {status}")


def run(argv=None):
    print("⚠️ 注意：09 依赖 Total_Workorder 是 04 刚刷新过的；"
          "单独运行前请先跑 04。")

    args = parse_args(argv)
    run_date = datetime.date.today()
    only_orders = parse_only_orders(args.only_orders)

    try:
        sh = open_spreadsheet()
        values = load_values(sh)
    except Exception as error:
        print(f"❌ 读取 Google Sheet 失败: {type(error).__name__}: {error}")
        return 1

    try:
        plan = build_plan(values, run_date, only_orders=only_orders)
    except Exception as error:
        print(f"❌ 解析 Total_Workorder 失败: {type(error).__name__}: {error}")
        return 1

    print_plan(plan, run_date)

    if only_orders is not None:
        planned = {
            d.order
            for d in plan.cnf_orders + plan.teco_orders + plan.skipped
        }
        missing = sorted(only_orders - planned)
        if missing:
            print(f"⚠️ --only-orders 中这些订单不在本次计划里: "
                  f"{', '.join(missing)}")

    if args.dry_run:
        print("dry-run：不启动 SAP、不写日志。")
        return 0

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not plan.cnf_orders and not plan.teco_orders:
        append_log_rows(
            sh,
            build_log_rows(plan, ExecutionResult([], [], 0),
                           run_date, timestamp),
        )
        print("没有需要处理的工单。")
        return 0

    session = None
    we_started = False
    try:
        we_started = ensure_sap_running()
        session = open_session()
        execution = execute_plan(
            plan,
            lambda orders: confirm_times(session, orders),
            lambda orders: complete_technically(session, orders),
        )
        print_execution(execution)
        append_log_rows(
            sh,
            build_log_rows(plan, execution, run_date, timestamp),
        )
        return execution.exit_code
    except Exception as error:
        print(f"❌ SAP 操作失败: {type(error).__name__}: {error}")
        return 1
    finally:
        if session is not None:
            close_session(session)
        if we_started:
            close_sap()


if __name__ == "__main__":
    sys.exit(run())
