"""Maintenance_Plan_Adherence_Gap 表的纯逻辑。

放在仓库根、且不 import win32com，是为了能在任何机器上跑测试 ——
00 的 main.py 顶部就 import win32com，在非 Windows 上根本 import 不进来，
所以凡是能被测试覆盖的逻辑都放这里，main.py 只留 SAP/Excel/Sheets 的管路。

这里做三件事：
  * 在导出文件的表头里按名字找出需要的列（不写死列号）
  * 把「参考日期」单元格统一成 YYYY-MM-DD
  * 按月拆分 Gap 表已有的行，并按 12 列的约定组装新行
"""

import datetime

# 前 10 列取自 IW39（变式 /KEL）的导出文件，顺序即写入顺序；
# 后 2 列由脚本生成。这份顺序要和 Gap 工作表的表头一致。
GAP_SOURCE_COLUMNS = [
    "参考日期", "通知", "订单", "功能位置", "设备",
    "描述", "成本中心", "系统状态", "ABC 标识", "订单类型",
]
GAP_HEADERS = GAP_SOURCE_COLUMNS + ["写入时间", "月份"]

WRITE_TIME_COLUMN_INDEX = GAP_HEADERS.index("写入时间")
MONTH_COLUMN_INDEX = GAP_HEADERS.index("月份")

# 「参考日期」的可能写法。分隔符决定顺序：斜杠按月在前（脚本给 SAP 填的就是
# %m/%d/%Y），点号按日在先（欧洲写法）。两种都认不出来就原样返回，不丢数据。
_DATE_FORMATS = ("%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d", "%m/%d/%Y", "%d.%m.%Y")


class MissingColumnsError(ValueError):
    """导出文件里缺少需要的列。"""


def _normalize_header(value):
    """去掉所有空白后再比，避开 'ABC 标识' 与 'ABC标识' 这类差异。"""
    return "".join(str(value).split())


def resolve_columns(header_row, required=None):
    """在表头行里找出每一列的 0 基下标。

    只做精确匹配（忽略空白），**不做包含匹配** —— 否则「订单」没找到时
    会拿「订单类型」顶上，把错的列写进表里。找不到就抛错，并把导出文件
    实际有哪些表头一起打出来。
    """
    required = GAP_SOURCE_COLUMNS if required is None else required

    found = {}
    for index, value in enumerate(header_row):
        key = _normalize_header(value)
        if key and key not in found:
            found[key] = index

    columns = {}
    missing = []
    for name in required:
        key = _normalize_header(name)
        if key in found:
            columns[name] = found[key]
        else:
            missing.append(name)

    if missing:
        available = "、".join(
            str(value).strip() for value in header_row if str(value).strip()
        )
        raise MissingColumnsError(
            f"导出文件里找不到这些列：{'、'.join(missing)}。"
            f"该文件实际的表头是：{available}"
        )

    return columns


def format_reference_date(value):
    """把「参考日期」单元格转成 YYYY-MM-DD。"""
    if value is None:
        return ""
    # datetime 是 date 的子类，必须先判
    if isinstance(value, datetime.datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, datetime.date):
        return value.strftime("%Y-%m-%d")

    text = str(value).strip()
    if not text:
        return ""
    for date_format in _DATE_FORMATS:
        try:
            parsed = datetime.datetime.strptime(text, date_format)
        except ValueError:
            continue
        return parsed.strftime("%Y-%m-%d")
    return text


def _cell_text(row, index):
    """取一格的文本；越界、空值都当空字符串。"""
    if index >= len(row):
        return ""
    value = row[index]
    if value is None:
        return ""
    return str(value).strip()


def split_rows_by_month(all_rows, year_month):
    """把 Gap 表的现有内容拆成「要保留的」和该月的首次写入时间。

    all_rows 是 worksheet.get_all_values() 的结果，第 0 行是表头。
    第 12 列等于 year_month 的行会被丢掉（调用方随后用最新结果重写），
    其余行原样保留 —— 其他月份的历史数据不受影响。

    认不出月份的行（列数不足等）一律保留，绝不误删。
    """
    if not all_rows:
        return [], None

    target = str(year_month)
    keep = [all_rows[0]]
    first_write_time = None

    for row in all_rows[1:]:
        if _cell_text(row, MONTH_COLUMN_INDEX) == target:
            if first_write_time is None:
                first_write_time = _cell_text(row, WRITE_TIME_COLUMN_INDEX) or None
        else:
            keep.append(row)

    return keep, first_write_time


def build_gap_rows(export_rows, columns, write_time, year_month):
    """从导出数据里挑出「系统状态不含 CNF」的行，组装成 12 列。

    系统状态为空的行走不进 MasterData 的分母，也就不是 gap，直接跳过。
    """
    status_index = columns["系统状态"]
    rows = []

    for row in export_rows:
        status = _cell_text(row, status_index)
        if not status or "CNF" in status.upper():
            continue

        values = []
        for name in GAP_SOURCE_COLUMNS:
            index = columns[name]
            if name == "参考日期":
                value = row[index] if index < len(row) else None
                values.append(format_reference_date(value))
            else:
                values.append(_cell_text(row, index))

        rows.append(values + [write_time, str(year_month)])

    return rows
