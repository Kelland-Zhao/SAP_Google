import warnings
warnings.filterwarnings('ignore', category=FutureWarning)

import ssl
ssl._create_default_https_context = ssl._create_unverified_context

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

import requests
_original_request = requests.Session.request
def _patched_request(self, *args, **kwargs):
    kwargs['verify'] = False
    return _original_request(self, *args, **kwargs)
requests.Session.request = _patched_request

import time
import win32com.client
import sys
import os
import datetime
import pandas as pd
import numpy as np
import re
import openpyxl
import gspread
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession

# 公共部分在仓库根的 sap_common.py，不在本脚本所在目录。
# 先把仓库根加进 sys.path，这样单个脚本仍然可以独立运行。
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sap_common import get_resource_path, start_sap, wait_for_sap_session, close_sap

# --- 核心配置 ---
# 文件路径配置
OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\IM_Equipment_Number"
OUTPUT_FILENAME_PREFIX = "IM_Equipment_Number"


def get_output_filename():
    """按当月 YYYYMM 生成文件名，当月文件覆盖，历史文件保留"""
    return f"{OUTPUT_FILENAME_PREFIX}_{datetime.date.today().strftime('%Y%m')}.xlsx"

# Google Sheets 配置
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM/edit?gid=151672918#gid=151672918'
WORKSHEET_NAME = 'Equipment_Number_EAM'
SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')

# 由 D 列前 8 个字符生成的新增列，写入位置为导出文件现有末列+1
TAG_COLUMN_HEADER = '机台号 - Tag'

# IH08 筛选参数
WORK_CENTERS = ['PMMSXFAC', 'PMMSXTF1', 'PMMSXWHS', 'PMMSXPK1', 'PMMSXTF2', 'PMMSXIN2', 'PMMSXIN1']
MULTI_SELECT_ROW = ('wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010'
                    '/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I')


def get_date_range():
    """
    基于函数运行时间输出两个日期（Start_date, End_date）
    Start_date = 运行月的第一天，End_date = 运行日期
    日期格式为 MM/DD/YYYY
    
    Returns:
        tuple: (Start_date, End_date) 格式为 (MM/DD/YYYY, MM/DD/YYYY)
    """
    today = datetime.date.today()
    
    # 计算当月第一天
    start_date = today.replace(day=1)
    
    # 格式化日期为 MM/DD/YYYY
    start_date_str = start_date.strftime("%m/%d/%Y")
    end_date_str = today.strftime("%m/%d/%Y")
    
    return start_date_str, end_date_str


def get_equipment_number(session):
    """
    执行 IH08 事务码，获取设备编号数据。
    会话由调用方通过 sap_common.wait_for_sap_session() 取得后传入。
    """
    print("正在执行 IH08...")
    
    try:
        # 最大化窗口
        session.findById("wnd[0]").maximize()
        
        # 执行 IH08 事务码
        session.findById("wnd[0]/tbar[0]/okcd").text = "IH08"
        session.findById("wnd[0]").sendVKey(0)
        
        # 设置日期区间
        start_date_str, end_date_str = get_date_range()
        print(f"日期范围: {start_date_str} - {end_date_str}")
        session.findById("wnd[0]/usr/ctxtDATUV").text = start_date_str
        session.findById("wnd[0]/usr/ctxtDATUB").text = end_date_str
        session.findById("wnd[0]/usr/ctxtDATUB").setFocus()
        session.findById("wnd[0]/usr/ctxtDATUB").caretPosition = 10
        
        # 系统状态多选：标记
        session.findById("wnd[0]/usr/btn%_STAE1_%_APP_%-VALU_PUSH").press()
        session.findById(f"{MULTI_SELECT_ROW}[1,0]").text = "标记"
        session.findById(f"{MULTI_SELECT_ROW}[1,0]").setFocus()
        session.findById(f"{MULTI_SELECT_ROW}[1,0]").caretPosition = 2
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(1)
        
        # 主工作中心多选
        session.findById("wnd[0]/usr/ctxtGEWRK-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtGEWRK-LOW").caretPosition = 0
        session.findById("wnd[0]/usr/btn%_GEWRK_%_APP_%-VALU_PUSH").press()
        
        for index, work_center in enumerate(WORK_CENTERS):
            session.findById(f"{MULTI_SELECT_ROW}[1,{index}]").text = work_center
        last_index = len(WORK_CENTERS) - 1
        session.findById(f"{MULTI_SELECT_ROW}[1,{last_index}]").setFocus()
        session.findById(f"{MULTI_SELECT_ROW}[1,{last_index}]").caretPosition = len(WORK_CENTERS[last_index])
        print(f"主工作中心筛选: {', '.join(WORK_CENTERS)}")
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        time.sleep(1)
        
        # 设置工厂
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = "CN15"
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").caretPosition = 4
        
        # 设置输出布局
        session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KEL"
        session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
        session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 4
        
        # 执行查询
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(3)
        
        # 导出数据
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        time.sleep(1)
        
        # 确认导出对话框
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            pass
        
        # 选择导出格式（本地文件）
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select()
        session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus()
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        time.sleep(1)
        
        # 确认导出
        try:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
        except Exception:
            pass
        
        print("✅ 设备编号数据查询完成。")
        
        # 等待 Excel 打开
        print("等待 Excel 文件打开...")
        time.sleep(5)
        
        # 确保输出目录存在
        if not os.path.exists(OUTPUT_DIR):
            try:
                os.makedirs(OUTPUT_DIR)
                print(f"创建目录: {OUTPUT_DIR}")
            except Exception as e:
                print(f"错误: 无法创建输出目录。{e}")
                raise
        
        # 构建完整文件路径
        output_path = os.path.join(OUTPUT_DIR, get_output_filename())
        
        # 如果文件已存在，先删除
        if os.path.exists(output_path):
            try:
                os.remove(output_path)
                print(f"已删除旧文件: {output_path}")
            except Exception as e:
                print(f"警告: 无法删除旧文件，可能正在被使用。{e}")
        
        # 保存 Excel 文件（在保存前会自动处理 D 列到 M 列）
        print(f"正在保存 Excel 文件到: {output_path}")
        if save_and_rename_active_excel(output_path):
            print(f"✅ 文件已成功保存: {output_path}")
        else:
            raise RuntimeError("文件保存失败，请检查 Excel 是否已打开")

    except Exception as e:
        print(f"❌ 执行 SAP 操作时发生错误: {e}")
        import traceback
        traceback.print_exc()
        raise

def save_and_rename_active_excel(new_full_path, original_window_title="Worksheet in excel (1)"):
    """
    查找屏幕上活动的 Excel 实例，并将其另存为到指定路径。
    
    Args:
        new_full_path (str): 带有新文件名和路径的完整路径 (例如: C:/NewFolder/FinalReport.xlsx)
        original_window_title (str): SAP 导出的 Excel 窗口的标题
    """
    
    # 转换为 Windows 兼容路径 (win32com 需要反斜杠，但正斜杠通常也能工作)
    new_full_path = new_full_path.replace('/', '\\')
    
    # 确保目录存在
    output_dir = os.path.dirname(new_full_path)
    if output_dir and not os.path.exists(output_dir):
        try:
            os.makedirs(output_dir)
            print(f"创建目录: {output_dir}")
        except Exception as e:
            print(f"错误: 无法创建目录 {output_dir}: {e}")
            return False
    
    try:
        # 1. 连接到活动的 Excel 应用程序
        # 尝试使用 GetObject 连接到当前正在运行的 Excel 实例
        ExcelApp = win32com.client.GetActiveObject("Excel.Application")
        
        # 2. 找到正确的工作簿 (通过标题或直接使用 ActiveWorkbook)
        # 依赖 ActiveWorkbook 不总是可靠，但对于刚从 SAP 导出的文件，通常有效
        Workbook = ExcelApp.ActiveWorkbook
        
        # 3. 在保存前处理数据：提取 D 列前8个字符到 M 列
        if not process_excel_d_to_m_column():
            raise RuntimeError("处理 Excel 数据（提取 Tag 列）失败")
        
        # 4. 执行另存为操作
        # FileFormat=51 是用于 .xlsx 格式的数字代码
        Workbook.SaveAs(new_full_path, FileFormat=51) 
        
        # 5. 关闭原工作簿
        Workbook.Close(SaveChanges=False)
        
        # 6. 退出 Excel 实例（完全关闭 Excel 程序）
        ExcelApp.Quit() 
        
        print(f"✅ Excel 文件成功另存为: {new_full_path}")
        return True
        
    except Exception as e:
        print(f"❌ 自动化 Excel 操作失败: {e}")
        print("请确保 Excel 实例正在运行且可见，并且 'win32com' 已安装。")
        return False


def process_excel_d_to_m_column():
    """
    在活动的 Excel 工作簿中，将 D 列每个单元格值的前8个字符提取到 M 列
    此函数在 Excel 另存为之前执行，直接操作活动的 Excel 应用程序
    
    Returns:
        bool: 操作是否成功
    """
    try:
        # 连接到活动的 Excel 应用程序
        ExcelApp = win32com.client.GetActiveObject("Excel.Application")
        Workbook = ExcelApp.ActiveWorkbook
        Worksheet = Workbook.ActiveSheet
        
        print("正在处理 Excel 数据：提取 B 列前8个字符到末列+1...")
        
        # 检查工作表是否被保护，如果被保护则先解除保护
        was_protected = False
        try:
            if Worksheet.ProtectContents:
                print("检测到工作表被保护，正在解除保护...")
                Worksheet.Unprotect()
                was_protected = True
        except Exception as e:
            print(f"⚠️ 检查工作表保护状态时出错（继续执行）: {e}")
        
        # 确保 Excel 应用程序处于可编辑状态
        ExcelApp.ScreenUpdating = False  # 关闭屏幕更新以提高性能
        ExcelApp.EnableEvents = False    # 禁用事件以提高性能
        
        try:
            # 获取工作表的最后一行，并定位到现有末列的下一列
            used_range = Worksheet.UsedRange
            last_row = used_range.Rows.Count
            target_col = used_range.Columns.Count + 1
            
            Worksheet.Cells(1, target_col).Value = TAG_COLUMN_HEADER
            print(f"Tag 列写入位置: 第 {target_col} 列，表头 {TAG_COLUMN_HEADER}")
            
            if last_row < 2:
                print("⚠️ 警告: 工作表只有表头行，无数据需要处理")
                return True
            
            # 批量读取 B 列「描述」数据（从第 2 行开始，跳过表头）
            d_range = Worksheet.Range(f"B2:B{last_row}")
            d_values = d_range.Value
            
            # Range.Value 返回的可能是元组或列表，需要统一处理
            # 如果是单行，可能是单个值；如果是多行，可能是元组的元组
            if not isinstance(d_values, (list, tuple)):
                d_values = [[d_values]]
            elif len(d_values) > 0 and not isinstance(d_values[0], (list, tuple)):
                # 如果是一维列表，转换为二维
                d_values = [[v] for v in d_values]
            
            # 准备 M 列的值列表
            m_values = []
            for d_row in d_values:
                # 获取该行的 B 列值（可能是元组或列表的第一个元素）
                d_value = d_row[0] if isinstance(d_row, (list, tuple)) else d_row
                
                # 提取前8个字符
                if d_value is not None and str(d_value).strip() != '':
                    str_value = str(d_value)
                    m_value = str_value[:8] if len(str_value) >= 8 else str_value
                else:
                    m_value = ''
                m_values.append([m_value])
            
            # 批量写入目标列（使用 Range 批量写入，更高效且更可靠）
            target_range = Worksheet.Range(Worksheet.Cells(2, target_col),
                                          Worksheet.Cells(last_row, target_col))
            target_range.Value = m_values
            
        finally:
            # 恢复 Excel 应用程序设置
            ExcelApp.ScreenUpdating = True
            ExcelApp.EnableEvents = True
        
        # 如果之前工作表被保护，可以选择重新保护（这里不重新保护，因为后续要保存）
        # if was_protected:
        #     Worksheet.Protect()
        
        print("✅ 数据处理完成")
        return True
        
    except Exception as e:
        print(f"⚠️ 处理 Excel 数据时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def write_to_google_sheet(excel_file_path, sheet_url, worksheet_name, auth_file, start_row=1):
    """
    读取 Excel 文件（含表头），清空整表后写入 Google Sheets
    
    Args:
        excel_file_path (str): Excel 文件的完整路径
        sheet_url (str): Google Sheets 的 URL
        worksheet_name (str): 目标工作表的名称
        auth_file (str): Google 服务账户 JSON 文件的路径
        start_row (int): 开始写入的行号（默认从第1行开始）
    """
    try:
        # 1. 读取 Excel 文件
        print(f"正在读取 Excel 文件: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"❌ 错误: 找不到文件 {excel_file_path}")
            return False
        
        # 不把首行当列名，保留原始表头（避开 pandas 对重复列名的自动重命名）
        df = pd.read_excel(excel_file_path, engine='openpyxl', header=None)
        
        df_data = df.copy()
        
        # 清理数据：替换 NaN、Infinity 等不符合 JSON 规范的值
        # 将 NaN 替换为空字符串
        df_data = df_data.fillna('')
        
        # 将 Infinity 和 -Infinity 替换为空字符串
        df_data = df_data.replace([np.inf, -np.inf], '')
        
        # 转换为列表
        data_to_write = df_data.values.tolist()
        
        # 进一步清理：确保所有值都是 JSON 兼容的
        def clean_value(value):
            """清理单个值，确保符合 JSON 规范"""
            # 优先检查是否为 NaN/NaT/空值
            if pd.isna(value) or value == '' or value is None:
                return ''
            # 处理 datetime 对象（包括 pandas Timestamp）
            if isinstance(value, (pd.Timestamp, datetime.datetime)):
                return value.strftime("%Y-%m-%d")
            # 处理 date 对象
            if isinstance(value, datetime.date):
                return value.strftime("%Y-%m-%d")
            if isinstance(value, (float, int)):
                if np.isinf(value) or np.isnan(value):
                    return ''
            # 将 numpy 类型转换为 Python 原生类型
            if isinstance(value, (np.integer, np.floating)):
                return value.item()
            return value
        
        # 清理所有值
        data_to_write = [[clean_value(cell) for cell in row] for row in data_to_write]
        
        if not data_to_write:
            print("⚠️ 警告: Excel 文件中没有数据可写入，跳过上传（不算失败）")
            return True  # 空数据不是故障，让 run_all 记为成功
        
        print(f"读取到 {len(data_to_write)} 行数据（含表头行）")
        
        # 2. 连接 Google Sheets
        print("正在连接 Google Sheets...")
        _scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        _credentials = Credentials.from_service_account_file(auth_file, scopes=_scopes)
        _authed_session = AuthorizedSession(_credentials)
        _authed_session.verify = False
        gc = gspread.Client(auth=_credentials, session=_authed_session)
        
        # 3. 打开工作簿和工作表
        sh = gc.open_by_url(sheet_url)
        worksheet = sh.worksheet(worksheet_name)
        
        # 4. 确定写入范围（从指定行开始）
        # 计算需要写入的行数和列数
        num_rows = len(data_to_write)
        num_cols = len(data_to_write[0]) if data_to_write else 0
        
        # 将列号转换为字母（支持超过26列）
        def col_num_to_letter(n):
            """将列号转换为 Excel 列字母（1 -> A, 27 -> AA）"""
            result = ""
            while n > 0:
                n -= 1
                result = chr(65 + (n % 26)) + result
                n //= 26
            return result
        
        # 构建范围字符串，例如 "A2:Z100"
        end_row = start_row + num_rows - 1
        end_col_letter = col_num_to_letter(num_cols)
        range_name = f"A{start_row}:{end_col_letter}{end_row}"
        
        # 5. 清空整表后写入数据
        print("正在清空工作表...")
        worksheet.clear()
        
        print(f"正在将数据写入 Google Sheets: {worksheet_name}，范围: {range_name}")
        worksheet.update(range_name=range_name, values=data_to_write)
        
        print(f"✅ 成功将 {num_rows} 行数据写入 Google Sheet: {worksheet_name}（从第 {start_row} 行开始）")
        return True
        
    except FileNotFoundError:
        print(f"❌ 错误: 找不到文件 {excel_file_path}")
        return False
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作表 '{worksheet_name}'。")
        return False
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作簿。请务必确认已将服务账户邮箱添加为该表格的 '编辑者'。")
        return False
    except Exception as e:
        print(f"❌ Google Sheets 写入失败。错误: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        # 1. 清理可能存在的 SAP 进程
        print("正在清理可能存在的 SAP 进程...")
        close_sap()
        time.sleep(2)
        
        # 2. 自动登录 SAP
        print("正在启动 SAP GUI...")
        start_sap()
        
        # 3. 等待 SAP 会话就绪
        print("等待 SAP 会话就绪...")
        session = wait_for_sap_session()

        # 4. 执行 SAP 操作 - 获取设备编号数据
        print("开始执行 SAP 操作...")
        get_equipment_number(session)
        
        # 5. 将 Excel 数据复制到 Google Sheets
        print("\n" + "="*50)
        print("开始将 Excel 数据上传到 Google Sheets...")
        print("="*50)
        
        excel_file_path = os.path.join(OUTPUT_DIR, get_output_filename())
        time.sleep(2)  # 等待文件完全保存
        
        if not write_to_google_sheet(
            excel_file_path=excel_file_path,
            sheet_url=GOOGLE_SHEET_URL,
            worksheet_name=WORKSHEET_NAME,
            auth_file=SERVICE_ACCOUNT_FILE,
            start_row=1
        ):
            close_sap()
            sys.exit(1)
        
        # 6. 操作完成后关闭 SAP
        print("\nSAP 操作完成，正在关闭 SAP...")
        time.sleep(2)  # 等待操作完成
        close_sap()
        
        print("\n" + "="*50)
        print("✅ 所有操作已完成！")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 程序执行过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        # 确保出错时也关闭 SAP
        close_sap()
        sys.exit(1)
    finally:
        print("\n执行完毕 / Completed.")

