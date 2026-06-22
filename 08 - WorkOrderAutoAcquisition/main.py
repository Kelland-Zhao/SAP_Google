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
import calendar
import pandas as pd
import numpy as np
import re
import openpyxl
import gspread

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- 核心配置 ---
# 文件路径配置
OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\WorkOrder"
OUTPUT_FILENAME = "Temporary_File.xlsx"

# Google Sheets 配置
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1YzMGIQ2RcBlGIadWh5yfxlCmOpCuOBHpgKfEVz8_W98/edit?gid=0#gid=0'
WORKSHEET_NAME = 'Database'
SERVICE_ACCOUNT_FILE = get_resource_path('../pyreadsp-b5b9c1909de6.json')

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
    Start_date = 运行月的第一天，End_date = 运行月的最后一天
    日期格式为 MM/DD/YYYY
    
    Returns:
        tuple: (Start_date, End_date) 格式为 (MM/DD/YYYY, MM/DD/YYYY)
    """
    today = datetime.date.today()
    
    # 计算当月第一天
    start_date = today.replace(day=1)
    
    # 计算当月最后一天
    last_day = calendar.monthrange(today.year, today.month)[1]
    end_date = today.replace(day=last_day)
    
    # 格式化日期为 MM/DD/YYYY
    start_date_str = start_date.strftime("%m/%d/%Y")
    end_date_str = end_date.strftime("%m/%d/%Y")
    
    return start_date_str, end_date_str


def get_work_order():
    """
    连接到 SAP GUI 并执行 IW39 事务码，获取工单数据
    基于 VBS 脚本转换而来
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
    
    print("成功连接到 SAP 会话。正在执行 IW39...")
    
    # 获取日期范围
    start_date, end_date = get_date_range()
    
    try:
        # 调整工作窗格大小
        session.findById("wnd[0]").resizeWorkingPane(232, 29, False)
        
        # 执行 IW39 事务码
        session.findById("wnd[0]/tbar[0]/okcd").text = "IW39"
        session.findById("wnd[0]").sendVKey(0)
        
        # 选中复选框
        session.findById("wnd[0]/usr/chkDY_OFN").selected = True
        session.findById("wnd[0]/usr/chkDY_IAR").selected = True
        session.findById("wnd[0]/usr/chkDY_IAR").setFocus()
        
        # 设置订单类型筛选
        session.findById("wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH").press()
        
        # 输入订单类型：PM02, PM03, PM10, ZPM4, ZPM8
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "PM02"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "PM03"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,2]").text = "PM10"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,3]").text = "ZPM4"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").text = "ZPM8"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,4]").caretPosition = 4
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        
        # 设置功能位置
        session.findById("wnd[0]/usr/ctxtSTRNO-LOW").text = "CN15*"
        
        # 设置日期范围（动态获取当月）
        session.findById("wnd[0]/usr/ctxtDATUV").text = start_date
        session.findById("wnd[0]/usr/ctxtDATUB").text = end_date
        
        # 设置计划工厂和维护工厂
        session.findById("wnd[0]/usr/ctxtIWERK-LOW").text = "CN15"
        session.findById("wnd[0]/usr/ctxtSWERK-LOW").text = "CN15"
        
        # 设置变式
        session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KEL"
        session.findById("wnd[0]/usr/ctxtVARIANT").setFocus()
        session.findById("wnd[0]/usr/ctxtVARIANT").caretPosition = 4
        
        # 执行查询
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)
        
        # 导出数据（使用btn[16]）
        session.findById("wnd[0]/tbar[1]/btn[16]").press()
        time.sleep(1)
        
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
        
        print(f"✅ 工单数据查询完成。日期范围: {start_date} 至 {end_date}，工单类型: PM02, PM03, PM10, ZPM4, ZPM8")
        
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
        
        # 保存 Excel 文件
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
        
        # 3. 执行另存为操作
        # FileFormat=51 是用于 .xlsx 格式的数字代码
        Workbook.SaveAs(new_full_path, FileFormat=51) 
        
        # 4. 关闭原工作簿 (可选)
        Workbook.Close(SaveChanges=False)
        
        # 5. 退出 Excel 实例 (如果这是唯一打开的工作簿)
        ExcelApp.Quit() 
        
        print(f"✅ Excel 文件成功另存为: {new_full_path}")
        return True
        
    except Exception as e:
        print(f"❌ 自动化 Excel 操作失败: {e}")
        print("请确保 Excel 实例正在运行且可见，并且 'win32com' 已安装。")
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
            if pd.isna(value) or value == '':
                return ''
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
        gc = gspread.service_account(filename=auth_file)
        
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
        
        # 5. 清空旧数据（从第2行开始）
        print(f"正在清空 Google Sheets 中的旧数据...")
        try:
            # 获取工作表的总行数和列数
            all_values = worksheet.get_all_values()
            total_rows = len(all_values)
            
            if total_rows > 1:  # 如果有数据行（不只是表头）
                # 清空从第2行到最后一行的所有数据
                clear_range = f"A{start_row}:{end_col_letter}{total_rows}"
                worksheet.batch_clear([clear_range])
                print(f"✅ 已清空范围: {clear_range}")
        except Exception as e:
            print(f"⚠️ 清空旧数据时出现警告: {e}，继续写入新数据...")
        
        # 6. 写入新数据
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
        
        # 4. 执行 SAP 操作 - 获取工单数据
        print("开始执行 SAP 操作...")
        get_work_order()
        
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
        print("\n按任意键退出...")
        input()
        
    except Exception as e:
        print(f"❌ 程序执行过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        # 确保出错时也关闭 SAP
        close_SAP()
        print("\n按任意键退出...")
        input()

