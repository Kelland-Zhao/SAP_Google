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

import subprocess
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

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- 核心配置 ---
# 文件路径配置
OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\IM_Equipment_Number"
OUTPUT_FILENAME = "Temporary_File.xlsx"

# Google Sheets 配置
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/12MXO53wJC8s_J-IE2uGY5jx35rnUE7rxW1xvwVU-FxM/edit?gid=151672918#gid=151672918'
WORKSHEET_NAME = 'Equipment_Number_EAM'
SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')

def sap_auto_logo():
    subprocess.check_call(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', '-system=CAP', '-client=321',
                           '-user=USERNAME', '-pw=PASSWORD', '-language=ZH'])  # Login to CAP
    time.sleep(15)
    print("sap open successfully")


def close_SAP():
    """使用 Windows taskkill 命令强制关闭所有 SAP GUI 进程 (saplogon.exe)"""
    try:
        os.system('taskkill /im saplogon.exe /t /f')
        print("SAP GUI 进程已成功关闭。")
    except Exception as e:
        print(f"关闭 SAP 进程失败: {e}")


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


def get_equipment_number():
    """
    连接到 SAP GUI 并执行 IH08 事务码，获取设备编号数据
    基于 IM_Equipment_Number.vbs 脚本转换而来
    """
    try:
        # 连接到 SAP GUI
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print(f"错误: 无法连接到 SAP GUI Scripting Engine 或找不到活动会话。请确保 SAP GUI 已登录。{e}")
        return
    
    print("成功连接到 SAP 会话。正在执行 IH08...")
    
    try:
        # 最大化窗口
        session.findById("wnd[0]").maximize()
        
        # 执行 IH08 事务码
        session.findById("wnd[0]/tbar[0]/okcd").text = "IH08"
        session.findById("wnd[0]").sendVKey(0)
        
        # 设置工作中心筛选
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = "CN15"
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").setFocus()
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").caretPosition = 4
        session.findById("wnd[0]/usr/btn%_STORT_%_APP_%-VALU_PUSH").press()
        
        # 输入工作中心类型：TB1M, TB2M, TB3M
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "TB1M"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "TB2M"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "TB3M"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        
        # 执行查询
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)
        
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
        
        # 设置当前单元格
        session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell").setCurrentCell(20, "EQKTX")
        
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
                return
        
        # 构建完整文件路径
        output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
        
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
            print(f"❌ 文件保存失败，请检查 Excel 是否已打开")
        
    except Exception as e:
        print(f"❌ 执行 SAP 操作时发生错误: {e}")
        import traceback
        traceback.print_exc()

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
        process_excel_d_to_m_column()
        
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
        
        print("正在处理 Excel 数据：提取 D 列前8个字符到 M 列...")
        
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
            # 获取工作表的最后一行
            last_row = Worksheet.UsedRange.Rows.Count
            
            # 批量读取 D 列的值
            d_range = Worksheet.Range(f"D1:D{last_row}")
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
                # 获取该行的 D 列值（可能是元组或列表的第一个元素）
                d_value = d_row[0] if isinstance(d_row, (list, tuple)) else d_row
                
                # 提取前8个字符
                if d_value is not None and str(d_value).strip() != '':
                    str_value = str(d_value)
                    m_value = str_value[:8] if len(str_value) >= 8 else str_value
                else:
                    m_value = ''
                m_values.append([m_value])
            
            # 批量写入 M 列（使用 Range 批量写入，更高效且更可靠）
            m_range = Worksheet.Range(f"M1:M{last_row}")
            m_range.Value = m_values
            
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


def extract_first_8_chars_to_m_column(excel_file_path):
    """
    将 D 列中每个单元格值的前8个字符提取出来，放入 M 列对应的行
    
    Args:
        excel_file_path (str): Excel 文件的完整路径
    
    Returns:
        bool: 操作是否成功
    """
    try:
        print(f"正在处理 Excel 文件: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"❌ 错误: 找不到文件 {excel_file_path}")
            return False
        
        # 读取 Excel 文件
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        # 确保 D 列和 M 列存在（如果不存在则创建）
        if len(df.columns) < 4:
            print("❌ 错误: Excel 文件至少需要4列（D列）")
            return False
        
        # 获取 D 列的列名（索引为3，因为从0开始）
        d_column = df.columns[3] if len(df.columns) > 3 else None
        
        if d_column is None:
            print("❌ 错误: 找不到 D 列")
            return False
        
        # 确保 M 列存在（索引为12）
        while len(df.columns) < 13:
            df[f'Column_{len(df.columns) + 1}'] = ''
        
        m_column = df.columns[12] if len(df.columns) > 12 else None
        
        # 处理 D 列，提取前8个字符并放入 M 列
        def extract_first_8(value):
            """提取值的前8个字符"""
            if pd.isna(value) or value is None:
                return ''
            # 转换为字符串并提取前8个字符
            str_value = str(value)
            return str_value[:8] if len(str_value) >= 8 else str_value
        
        # 应用函数到 D 列，结果放入 M 列
        df[m_column] = df[d_column].apply(extract_first_8)
        
        # 保存文件
        df.to_excel(excel_file_path, index=False, engine='openpyxl')
        
        print(f"✅ 成功将 D 列的前8个字符提取到 M 列")
        return True
        
    except Exception as e:
        print(f"❌ 处理 Excel 文件时发生错误: {e}")
        import traceback
        traceback.print_exc()
        return False


def write_to_google_sheet(excel_file_path, sheet_url, worksheet_name, auth_file, start_row=2):
    """
    读取 Excel 文件（从第二行开始）并将数据写入 Google Sheets（从第二行开始）
    
    Args:
        excel_file_path (str): Excel 文件的完整路径
        sheet_url (str): Google Sheets 的 URL
        worksheet_name (str): 目标工作表的名称
        auth_file (str): Google 服务账户 JSON 文件的路径
        start_row (int): 开始写入的行号（默认从第2行开始）
    """
    try:
        # 1. 读取 Excel 文件
        print(f"正在读取 Excel 文件: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"❌ 错误: 找不到文件 {excel_file_path}")
            return False
        
        # 读取 Excel 文件，第一行作为列名
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        # 从第二行开始获取数据（索引从0开始，所以是 iloc[1:]）
        df_data = df.iloc[0:].copy()
        
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
            print("⚠️ 警告: Excel 文件中没有数据可写入（从第二行开始）")
            return False
        
        print(f"读取到 {len(data_to_write)} 行数据（从第二行开始）")
        
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
        
        # 5. 写入数据
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
        close_SAP()
        time.sleep(2)
        
        # 2. 自动登录 SAP
        print("正在启动 SAP GUI...")
        sap_auto_logo()
        
        # 3. 等待 SAP 完全加载
        print("等待 SAP GUI 完全加载...")
        time.sleep(5)  # 额外等待时间，确保 SAP 完全就绪
        
        # 4. 执行 SAP 操作 - 获取设备编号数据
        print("开始执行 SAP 操作...")
        get_equipment_number()
        
        # 5. 将 Excel 数据复制到 Google Sheets
        print("\n" + "="*50)
        print("开始将 Excel 数据上传到 Google Sheets...")
        print("="*50)
        
        excel_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
        time.sleep(2)  # 等待文件完全保存
        
        write_to_google_sheet(
            excel_file_path=excel_file_path,
            sheet_url=GOOGLE_SHEET_URL,
            worksheet_name=WORKSHEET_NAME,
            auth_file=SERVICE_ACCOUNT_FILE,
            start_row=2
        )
        
        # 6. 操作完成后关闭 SAP
        print("\nSAP 操作完成，正在关闭 SAP...")
        time.sleep(2)  # 等待操作完成
        close_SAP()
        
        print("\n" + "="*50)
        print("✅ 所有操作已完成！")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 程序执行过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        # 确保出错时也关闭 SAP
        close_SAP()
    finally:
        input("\n按 Enter 键退出 / Press Enter to exit...")

