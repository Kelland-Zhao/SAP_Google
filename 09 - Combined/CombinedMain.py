import warnings
warnings.filterwarnings('ignore', category=FutureWarning)

import calendar
import io
import datetime
import importlib.util
import os
import subprocess
import sys
import time
import traceback
from contextlib import contextmanager
from pathlib import Path
from typing import Callable, Dict, List, Optional, Tuple, TypeVar

import win32com.client  # type: ignore


if hasattr(sys, "_MEIPASS"):
    ROOT_DIR = Path(sys._MEIPASS)
else:
    ROOT_DIR = Path(__file__).resolve().parent.parent
GLOBAL_CODE = "GLOBAL"
MODULE_CACHE: Dict[str, object] = {}
TaskRunner = Callable[[object, str], Tuple[bool, str]]
T = TypeVar("T")


@contextmanager
def temporary_cwd(path: Path):
    previous_cwd = os.getcwd()
    os.chdir(str(path))
    try:
        yield
    finally:
        os.chdir(previous_cwd)


class _PrefixedStream(io.TextIOBase):
    """Wraps stdout so that every line printed by sub-modules is prefixed with the task code."""

    def __init__(self, underlying, code: str) -> None:
        self._underlying = underlying
        self._code = code
        self._buf = ""

    def write(self, s: str) -> int:
        self._buf += s
        while "\n" in self._buf:
            line, self._buf = self._buf.split("\n", 1)
            if line:
                ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                self._underlying.write(f"[{ts}] [{self._code}] [OUT] {line}\n")
                self._underlying.flush()
        return len(s)

    def flush(self) -> None:
        if self._buf:
            ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._underlying.write(f"[{ts}] [{self._code}] [OUT] {self._buf}\n")
            self._buf = ""
        self._underlying.flush()


@contextmanager
def prefixed_stdout(code: str):
    original = sys.__stdout__
    sys.stdout = _PrefixedStream(original, code)
    try:
        yield
    finally:
        sys.stdout.flush()
        sys.stdout = original


def log(code: str, level: str, message: str) -> None:
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] [{code}] [{level}] {message}\n"
    sys.__stdout__.write(line)
    sys.__stdout__.flush()


def is_control_not_found(exc: Exception) -> bool:
    return "the control could not be found by id" in str(exc).lower()


def run_with_retry(code: str, description: str, action: Callable[[], T], prepare_retry: Optional[Callable[[], None]] = None) -> T:
    last_exc: Optional[Exception] = None
    for attempt in range(2):
        try:
            return action()
        except Exception as exc:  # noqa: BLE001
            last_exc = exc
            if attempt == 0 and is_control_not_found(exc):
                log(code, "WARN", f"{description} 时控件定位失败，正在重试 / Control not found, retrying")
                try:
                    if prepare_retry is not None:
                        prepare_retry()
                    else:
                        reset_sap_session()
                except Exception as reset_exc:  # noqa: BLE001
                    log(code, "WARN", f"重置 SAP 会话失败，终止重试: {reset_exc}")
                    break
                time.sleep(1)
                continue
            break

    assert last_exc is not None
    raise last_exc


def load_module(alias: str, file_path: Path) -> object:
    if alias in MODULE_CACHE:
        return MODULE_CACHE[alias]

    if alias in sys.modules:
        MODULE_CACHE[alias] = sys.modules[alias]
        return MODULE_CACHE[alias]

    if not file_path.exists():
        raise FileNotFoundError(f"无法找到模块文件: {file_path}")

    spec = importlib.util.spec_from_file_location(alias, str(file_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块 {alias} ({file_path})")

    module = importlib.util.module_from_spec(spec)
    with temporary_cwd(file_path.parent):
        spec.loader.exec_module(module)
    sys.modules[alias] = module
    MODULE_CACHE[alias] = module
    return module


def wait_for_file(path: Path, timeout: int = 180, interval: int = 2) -> bool:
    deadline = time.time() + timeout
    lock_suffix = ".lockcheck"
    failed_close_attempts = 0
    max_force_kill_attempts = 3

    def try_lock() -> bool:
        temp_path = path.with_name(path.name + lock_suffix)
        if temp_path.exists():
            try:
                temp_path.unlink()
            except Exception:
                return False

        try:
            path.rename(temp_path)
            temp_path.rename(path)
            return True
        except (PermissionError, OSError):
            return False
        finally:
            if temp_path.exists() and not path.exists():
                try:
                    temp_path.rename(path)
                except Exception:
                    pass

    def close_excel_workbook(target: Path) -> bool:
        normalized = str(target.resolve()).lower()
        try:
            excel = win32com.client.GetActiveObject("Excel.Application")
        except Exception:
            return False

        workbook_found = False
        try:
            count = excel.Workbooks.Count
            for idx in range(1, count + 1):
                workbook = excel.Workbooks(idx)
                try:
                    try:
                        full_name = workbook.FullName  # type: ignore[attr-defined]
                    except Exception:
                        continue
                    if str(Path(full_name).resolve()).lower() == normalized:
                        workbook.Close(SaveChanges=False)
                        workbook_found = True
                        break
                finally:
                    workbook = None

            if workbook_found:
                time.sleep(0.5)
                if excel.Workbooks.Count == 0:
                    excel.Quit()
                return True
        finally:
            excel = None

        return False

    while time.time() < deadline:
        if path.exists():
            if try_lock():
                return True
            if close_excel_workbook(path):
                failed_close_attempts = 0
                continue

            failed_close_attempts += 1
            if failed_close_attempts >= max_force_kill_attempts:
                log(GLOBAL_CODE, "WARN", f"文件被 Excel 占用，执行强制关闭: {path}")
                os.system('taskkill /im excel.exe /t /f')
                failed_close_attempts = 0
                time.sleep(1)
                continue
        time.sleep(interval)

    if path.exists():
        if try_lock():
            return True
        if close_excel_workbook(path):
            if try_lock():
                return True
        log(GLOBAL_CODE, "WARN", f"等待超时，最后尝试强制关闭 Excel: {path}")
        os.system('taskkill /im excel.exe /t /f')
        time.sleep(1)
        if try_lock():
            return True

    return False


def get_sap_session():
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.Children(0)
    return connection.Children(0)


def reset_sap_session() -> object:
    session = get_sap_session()

    for idx in range(1, 4):
        try:
            window = session.findById(f"wnd[{idx}]")
        except Exception:
            break
        try:
            window.close()
        except Exception:
            try:
                session.findById(f"wnd[{idx}]/tbar[0]/btn[0]").press()
            except Exception:
                pass
        time.sleep(0.2)

    for _ in range(3):
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(0.5)
            break
        except Exception:
            try:
                session.findById("wnd[0]").sendVKey(12)
                time.sleep(0.5)
            except Exception:
                break

    return session


def run_task_00(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    today = datetime.date.today()
    year_month = today.strftime("%Y%m")
    output_path = Path(module.OUTPUT_DIR) / f"{year_month}_Maintenance_Plan_Adherence.xlsx"

    log(code, "INFO", "执行 IW39，获取维护计划遵守率数据 / Running IW39 for maintenance adherence")
    run_with_retry(
        code,
        "执行 IW39",
        lambda: module.get_maintenance_plan_iw39(),
        reset_sap_session,
    )

    if not wait_for_file(output_path):
        return False, f"未生成维护计划遵守率文件: {output_path}"

    return True, f"已生成 {output_path.name}"


def run_task_01(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    today = datetime.date.today()
    year_month = today.strftime("%Y%m")
    output_path = Path(module.OUTPUT_DIR) / f"{year_month}_Maintenance_Effectiveness.xlsx"

    log(code, "INFO", "执行 IW47，获取维护有效性数据 / Running IW47 for maintenance effectiveness")
    run_with_retry(
        code,
        "执行 IW47",
        lambda: module.get_maintenance_effectiveness_iw47(),
        reset_sap_session,
    )

    if not wait_for_file(output_path):
        return False, f"未生成维护有效性文件: {output_path}"

    return True, f"已生成 {output_path.name}"


def run_task_02(module: object, code: str) -> Tuple[bool, str]:
    today = datetime.date.today()
    first_day = today.replace(day=1)
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1])

    first_day_str = first_day.strftime("%m/%d/%Y")
    last_day_str = last_day.strftime("%m/%d/%Y")
    year_month = today.strftime("%Y%m")

    log(code, "INFO", "执行 IH08，导出 A 级关键设备 / Running IH08 for critical equipment")
    log(code, "INFO", "执行 IP18，导出含维护计划设备 / Running IP18 for equipment with maintenance plan")

    def execute_export() -> Tuple[str, str]:
        session_local = reset_sap_session()
        ih08_path = module.get_a_equipments_ih08(session_local, year_month, first_day_str, last_day_str)
        ip18_path = module.get_equipments_with_plan_ip18(session_local, year_month, ih08_path)
        return ih08_path, ip18_path

    try:
        ih08_file, ip18_file = run_with_retry(code, "执行 IH08/IP18", execute_export)
    except Exception as exc:  # noqa: BLE001
        return False, f"IH08/IP18 执行失败: {exc}"

    if not ih08_file or not Path(ih08_file).exists():
        return False, "IH08 导出失败，未找到 A 级关键设备文件"

    if not ip18_file or not Path(ip18_file).exists():
        return False, "IP18 导出失败，未找到有维护计划的设备文件"

    data = module.process_critical_equipment_data(ih08_file, ip18_file, year_month)
    if not data:
        return False, "数据处理失败"

    upload_ok = module.upload_to_google_sheets(
        data=data,
        sheet_id=module.GOOGLE_SHEET_ID,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
    )
    if not upload_ok:
        return False, "Google Sheets 更新失败"

    return True, f"IH08: {Path(ih08_file).name}, IP18: {Path(ip18_file).name}"


def run_task_03(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    excel_path = Path(module.OUTPUT_DIR) / module.OUTPUT_FILENAME

    log(code, "INFO", "执行 ZSE16，导出安全库存 / Running ZSE16 for safety stock")
    run_with_retry(
        code,
        "执行 ZSE16",
        lambda: module.get_safety_stock_zse16(),
        reset_sap_session,
    )

    if not wait_for_file(excel_path):
        return False, f"未生成安全库存文件: {excel_path}"

    upload_ok = module.upload_to_google_sheets(
        excel_file_path=str(excel_path),
        sheet_id=module.GOOGLE_SHEET_ID,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
    )
    if not upload_ok:
        return False, "Google Sheets 上传失败"

    return True, f"安全库存数据已上传: {excel_path.name}"


def run_task_04(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    excel_path = Path(module.OUTPUT_DIR) / module.OUTPUT_FILENAME

    log(code, "INFO", "执行 IW39，导出全年工单 / Running IW39 for total workorder")
    run_with_retry(
        code,
        "执行 IW39",
        lambda: module.get_work_order(),
        reset_sap_session,
    )

    if not wait_for_file(excel_path):
        return False, f"未生成工单文件: {excel_path}"

    upload_ok = module.write_to_google_sheet(
        excel_file_path=str(excel_path),
        sheet_url=module.GOOGLE_SHEET_URL,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
        start_row=2,
    )
    if not upload_ok:
        return False, "Google Sheets 写入失败"

    return True, f"工单数据已写入: {excel_path.name}"


def run_task_05(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    today = datetime.date.today()
    end_period = today.strftime("%m/%Y")
    start_period = today.replace(year=today.year - 1).strftime("%m/%Y")

    sap_dir = module.OUTPUT_DIR.replace('\\', '/')
    file_name = module.OUTPUT_FILENAME
    txt_path = Path(module.OUTPUT_DIR) / file_name
    xlsx_path = Path(module.OUTPUT_DIR) / file_name.replace(".txt", ".xlsx")

    log(code, "INFO", "执行 MC.7，获取库存周转 / Running MC.7 for stock turnover")
    run_with_retry(
        code,
        "执行 MC.7",
        lambda: module.run_sap_automation(start_period, end_period, sap_dir, file_name),
        reset_sap_session,
    )

    if not wait_for_file(txt_path):
        return False, f"未生成库存周转 TXT 文件: {txt_path}"

    df = module.parse_sap_list_report_to_dataframe(str(txt_path))
    if df is None or df.empty:
        return False, "TXT 解析失败或无数据"

    try:
        if xlsx_path.exists():
            xlsx_path.unlink()
        df.to_excel(xlsx_path, index=False, engine='openpyxl')
    except Exception as exc:
        return False, f"写入 Excel 失败: {exc}"

    target_value = df.iloc[0]['周转评估-V'] if '周转评估-V' in df.columns else ''
    if module.pd.isna(target_value):
        target_value = ''

    month_key = today.strftime("%Y%m")
    try:
        module.write_to_google_sheet(
            value=target_value,
            sheet_url=module.GOOGLE_SHEET_URL,
            worksheet_name=module.WORKSHEET_NAME,
            target_month_key=month_key,
            auth_file=module.SERVICE_ACCOUNT_FILE,
        )
    except Exception as exc:
        return False, f"Google Sheets 写入失败: {exc}"

    try:
        if txt_path.exists():
            txt_path.unlink()
    except Exception:
        pass

    return True, f"已更新 {month_key} 周转率，值: {target_value}"


def run_task_06(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    excel_path = Path(module.OUTPUT_DIR) / module.OUTPUT_FILENAME

    log(code, "INFO", "执行 IH08，导出 IM 设备编号 / Running IH08 for IM equipment numbers")
    run_with_retry(
        code,
        "执行 IH08",
        lambda: module.get_equipment_number(),
        reset_sap_session,
    )

    if not wait_for_file(excel_path):
        return False, f"未生成 IM 设备编号文件: {excel_path}"

    upload_ok = module.write_to_google_sheet(
        excel_file_path=str(excel_path),
        sheet_url=module.GOOGLE_SHEET_URL,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
        start_row=2,
    )
    if not upload_ok:
        return False, "Google Sheets 写入失败"

    return True, f"IM 设备编号已同步: {excel_path.name}"


def run_task_07(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    excel_path = Path(module.OUTPUT_DIR) / module.OUTPUT_FILENAME

    log(code, "INFO", "执行 MB52，导出库存清单 / Running MB52 for inventory list")
    run_with_retry(
        code,
        "执行 MB52",
        lambda: module.get_inventory_mb52(),
        reset_sap_session,
    )

    if not wait_for_file(excel_path):
        return False, f"未生成库存文件: {excel_path}"

    upload_ok = module.upload_to_google_sheets(
        excel_file_path=str(excel_path),
        sheet_id=module.GOOGLE_SHEET_ID,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
    )
    if not upload_ok:
        return False, "Google Sheets 上传失败"

    return True, f"库存数据已同步: {excel_path.name}"


def run_task_08(module: object, code: str) -> Tuple[bool, str]:
    try:
        reset_sap_session()
    except Exception as exc:
        log(code, "WARN", f"重置 SAP 会话失败: {exc}")

    excel_path = Path(module.OUTPUT_DIR) / module.OUTPUT_FILENAME

    log(code, "INFO", "执行 IW39，获取重点工单明细 / Running IW39 for detailed workorders")
    run_with_retry(
        code,
        "执行 IW39",
        lambda: module.get_work_order(),
        reset_sap_session,
    )

    if not wait_for_file(excel_path):
        return False, f"未生成工单明细文件: {excel_path}"

    upload_ok = module.write_to_google_sheet(
        excel_file_path=str(excel_path),
        sheet_url=module.GOOGLE_SHEET_URL,
        worksheet_name=module.WORKSHEET_NAME,
        auth_file=module.SERVICE_ACCOUNT_FILE,
        start_row=2,
    )
    if not upload_ok:
        return False, "Google Sheets 写入失败"

    return True, f"工单明细已写入: {excel_path.name}"


TASK_DEFINITIONS: List[Dict[str, object]] = [
    {
        "code": "T0",
        "title": "维护计划遵守率 Maintenance Plan Adherence",
        "path": ROOT_DIR / "00 - Maintenance_Plan_Adherence_Sync" / "main.py",
        "alias": "module00_main",
        "runner": run_task_00,
    },
    {
        "code": "T1",
        "title": "维护有效性 Maintenance Effectiveness",
        "path": ROOT_DIR / "01 - Maintenance_Effectiveness_Sync" / "main.py",
        "alias": "module01_main",
        "runner": run_task_01,
    },
    {
        "code": "T2",
        "title": "关键设备维护计划 Critical Equipment Maintenance Plan",
        "path": ROOT_DIR / "02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync" / "main.py",
        "alias": "module02_main",
        "runner": run_task_02,
    },
    {
        "code": "T3",
        "title": "安全库存 Safety Stock",
        "path": ROOT_DIR / "03 - Safety_Stock_ZSE16" / "main.py",
        "alias": "module03_main",
        "runner": run_task_03,
    },
    {
        "code": "T4",
        "title": "全年工单 Total Workorder",
        "path": ROOT_DIR / "04 - Total_Workorder" / "main.py",
        "alias": "module04_main",
        "runner": run_task_04,
    },
    {
        "code": "T5",
        "title": "库存周转 Stock Turnover",
        "path": ROOT_DIR / "05 - Stock_Turnover" / "main.py",
        "alias": "module05_main",
        "runner": run_task_05,
    },
    {
        "code": "T6",
        "title": "IM 设备编号 IM Equipment Number",
        "path": ROOT_DIR / "06 - IM_Equipment_Number" / "main.py",
        "alias": "module06_main",
        "runner": run_task_06,
    },
    {
        "code": "T7",
        "title": "库存清单 Inventory MB52",
        "path": ROOT_DIR / "07 - Inventory_MB52" / "main.py",
        "alias": "module07_main",
        "runner": run_task_07,
    },
    {
        "code": "T8",
        "title": "重点工单 WorkOrder Auto Acquisition",
        "path": ROOT_DIR / "08 - WorkOrderAutoAcquisition" / "main.py",
        "alias": "module08_main",
        "runner": run_task_08,
    },
]


def execute_task(task: Dict[str, object]) -> Dict[str, object]:
    code = task["code"]
    title = task["title"]
    module = load_module(task["alias"], task["path"])
    runner: TaskRunner = task["runner"]  # type: ignore

    log(code, "INFO", f"🚀 {title} 开始执行 / Starting")
    start_time = time.time()

    success = False
    detail = ""

    try:
        with prefixed_stdout(code):
            success, detail = runner(module, code)
        icon = "✅" if success else "❌"
        level = "SUCCESS" if success else "ERROR"
        message = detail if detail else ("任务完成 / Task completed" if success else "任务失败 / Task failed")
        log(code, level, f"{icon} {message}")
    except Exception:
        success = False
        traceback_lines = traceback.format_exc().strip().splitlines()
        log(code, "ERROR", "❌ 出现未处理异常 / Unhandled exception raised")
        for line in traceback_lines:
            log(code, "TRACE", line)
        detail = "\n".join(traceback_lines)

    duration = time.time() - start_time
    return {
        "code": code,
        "title": title,
        "success": success,
        "detail": detail,
        "duration": duration,
    }


def print_summary(results: List[Dict[str, object]]) -> None:
    print("\n" + "=" * 80)
    print("执行摘要 / Execution Summary")
    print("=" * 80)

    for item in results:
        status = "✅ 成功 / Success" if item["success"] else "❌ 失败 / Failed"
        print(f"{status} | {item['code']} | {item['title']} | 用时 / Duration: {item['duration']:.1f}s")
        if not item["success"] and item["detail"]:
            print(f"   失败原因 / Reason: {item['detail']}")


def main() -> None:
    log(GLOBAL_CODE, "INFO", "组合流程启动 / Combined workflow starting")

    try:
        module0 = load_module("module00_main", ROOT_DIR / "00 - Maintenance_Plan_Adherence_Sync" / "main.py")
    except Exception as exc:
        log(GLOBAL_CODE, "ERROR", f"无法加载 00 模块: {exc}")
        return

    log(GLOBAL_CODE, "INFO", "清理残留 SAP 进程 / Closing residual SAP processes")
    try:
        module0.close_SAP()
        time.sleep(2)
    except Exception as exc:
        log(GLOBAL_CODE, "WARN", f"关闭 SAP 进程时出现警告: {exc}")

    log(GLOBAL_CODE, "INFO", "启动 SAP GUI / Launching SAP GUI")
    try:
        module0.sap_auto_logo()
    except subprocess.CalledProcessError as exc:
        log(GLOBAL_CODE, "ERROR", f"SAP GUI 启动失败: {exc}")
        return

    log(GLOBAL_CODE, "INFO", "等待 SAP GUI 完全加载 / Waiting for SAP GUI to be ready")
    time.sleep(5)

    results: List[Dict[str, object]] = []
    try:
        for task in TASK_DEFINITIONS:
            result = execute_task(task)
            results.append(result)
    finally:
        log(GLOBAL_CODE, "INFO", "统一关闭 SAP GUI / Closing SAP GUI")
        try:
            module0.close_SAP()
        except Exception as exc:
            log(GLOBAL_CODE, "WARN", f"关闭 SAP 进程出现警告: {exc}")

    print_summary(results)

    failures = [item for item in results if not item["success"]]
    if failures:
        log(GLOBAL_CODE, "INFO", f"存在 {len(failures)} 个失败任务，请根据摘要排查 / {len(failures)} task(s) failed")
    else:
        log(GLOBAL_CODE, "INFO", "全部任务成功完成 / All tasks succeeded")


if __name__ == "__main__":
    main()
