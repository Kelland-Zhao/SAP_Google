"""10 - 工单取消技术完成。

读同一 spreadsheet 的 `取消技术完成` 工作表：
  * A 列订单号，B 列 Done，C 列失败原因；没有表头。
  * 只处理 B 列为空的行；逐张工单用 IW32 取消技术完成。
  * 成功写 B=Done 并清空 C；失败写 C=时间 | 原因，B 保持空。

这是一次性工具，不加入 run_all.ps1。
"""

import argparse
import datetime
import os
import sys
import time

import gspread

# 公共部分在仓库根，不在本脚本所在目录。
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from sap_common import (ensure_sap_running, open_session, close_session,
                        close_sap, get_resource_path)
from workorder_cancel_utils import (
    build_plan,
    execute_plan,
    format_orders,
    parse_only_orders,
    pending_write_orders,
)
from workorder_cancel_sap import cancel_technical_completion, prepare_session

SPREADSHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98/edit"
)
WORKSHEET = "取消技术完成"
SERVICE_ACCOUNT_FILE = get_resource_path("pyreadsp-b5b9c1909de6.json")
WRITE_ATTEMPTS = 3


def parse_args(argv=None):
    parser = argparse.ArgumentParser(description="10 - 工单取消技术完成")
    parser.add_argument(
        "--dry-run", action="store_true",
        help="只读 Sheet 并打印计划，不启动 SAP、不写任何单元格",
    )
    parser.add_argument(
        "--only-orders", default=None,
        help="只处理指定订单号，逗号/分号分隔（测试用）",
    )
    return parser.parse_args(argv)


def open_spreadsheet():
    gc = gspread.service_account(filename=SERVICE_ACCOUNT_FILE)
    return gc.open_by_url(SPREADSHEET_URL)


def load_values(sh):
    return sh.worksheet(WORKSHEET).get_all_values()


def print_plan(plan, only_orders):
    stats = plan.stats
    print(
        f"读到 {stats['total_rows']} 行；"
        f"已处理(Done) {stats['done']} 行；"
        f"空订单 {stats['empty_order']} 行；"
        f"筛选掉 {stats['filtered']} 行；"
        f"待处理 {stats['pending_orders']} 张工单（{stats['pending_rows']} 行）"
    )
    print(f"待处理订单: {format_orders(plan)}")
    if only_orders is not None:
        print(f"只处理指定订单: {', '.join(sorted(only_orders))}")
    for warning in plan.warnings:
        print(f"⚠️ {warning}")


def write_updates(ws, updates, attempts=WRITE_ATTEMPTS):
    """把 (row, col, value) 写入工作表；失败重试，最终返回是否成功。"""
    if not updates:
        return True

    cells = [gspread.Cell(row, col, value) for row, col, value in updates]
    for attempt in range(1, attempts + 1):
        try:
            ws.update_cells(cells, value_input_option="RAW")
            return True
        except Exception as error:
            if attempt == attempts:
                print(f"❌ 写入失败: {type(error).__name__}: {error}")
                return False
            time.sleep(1 if attempt == 1 else 2)
    return False


def print_summary(plan, execution):
    successes = [record for record in execution.records if record.success]
    failures = [record for record in execution.records if not record.success]
    print(
        f"==== 完成：成功 {len(successes)} 张工单，"
        f"失败 {len(failures)} 张工单，"
        f"已处理跳过 {plan.stats['done']} 行 ===="
    )
    if failures:
        print("失败清单:")
        for record in failures:
            print(f"  - {record.order}: {record.message}")


def run(argv=None):
    args = parse_args(argv)
    only_orders = parse_only_orders(args.only_orders)

    try:
        sh = open_spreadsheet()
        values = load_values(sh)
    except Exception as error:
        print(f"❌ 读取 Google Sheet 失败: {type(error).__name__}: {error}")
        return 1

    try:
        plan = build_plan(values, only_orders=only_orders)
    except Exception as error:
        print(f"❌ 解析工作表失败: {type(error).__name__}: {error}")
        return 1

    print_plan(plan, only_orders)

    if args.dry_run:
        print("dry-run：不启动 SAP、不写任何单元格。")
        return 0

    if not plan.orders:
        print("没有需要处理的工单。")
        return 0

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    session = None
    we_started = False

    try:
        we_started = ensure_sap_running()
        session = open_session()
        prepare_session(session)
        ws = sh.worksheet(WORKSHEET)

        def cancel_fn(order):
            print(f"[10] {order} ...")
            ok, message = cancel_technical_completion(session, order)
            status = "成功" if ok else "失败"
            print(f"[10] {order} {status}：{message}")
            return ok, message

        execution = execute_plan(
            plan, cancel_fn, lambda updates: write_updates(ws, updates),
            timestamp,
        )

        if execution.aborted:
            processed = len(execution.records)
            print(
                f"❌ SAP 会话不可用，已中止；"
                f"第 {processed + 1} 个待处理订单起未处理。"
            )

        final_write_ok = True
        if execution.pending_writes:
            print(f"⚠️ 重新写入 {len(execution.pending_writes)} 个单元格...")
            final_write_ok = write_updates(ws, execution.pending_writes)
            if not final_write_ok:
                print(
                    "❌ 仍有结果写不进 Sheet；请手工补写以下行"
                    "（成功的 B 写 Done、C 清空；失败的 C 写原因）："
                )
                for order, rows in pending_write_orders(execution):
                    row_text = ", ".join(str(row) for row in rows)
                    print(f"  - 订单 {order}，行 {row_text}")

        print_summary(plan, execution)

        failed = any(not record.success for record in execution.records)
        if execution.aborted or failed or not final_write_ok:
            return 1
        return 0
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
