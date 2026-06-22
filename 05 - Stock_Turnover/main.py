import subprocess
import time
import win32com.client
import sys
import os
import datetime
import pandas as pd
import re
import openpyxl
import gspread
import requests
import urllib3
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import AuthorizedSession

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# --- 核心配置 ---
# 注意：请在使用前替换 USERNAME 和 PASSWORD
SAP_SYSTEM = 'CAP'
SAP_CLIENT = '321'
SAP_USER = 'USERNAME'
SAP_PASSWORD = 'PASSWORD'
SAP_LANGUAGE = 'ZH'

# --- 文件路径配置 ---
OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\StockTurnover"
OUTPUT_FILENAME = "Temporary_StockTurnover.txt"

# --- Google Sheets 配置 ---
SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')
GOOGLE_SHEET_URL = 'https://docs.google.com/spreadsheets/d/1Pa0A6T_qmu5hl2zEPHil-jDDUwc0v_YXGk1PCSgSMB4/edit'
WORKSHEET_NAME = 'Summary'

# -----------------------------------------------


def close_SAP():
    """使用 Windows taskkill 命令强制关闭所有 SAP GUI 进程 (saplogon.exe)"""
    try:
        os.system('taskkill /im saplogon.exe /t /f')
        print("SAP GUI 进程已成功关闭。")
    except Exception as e:
        print(f"关闭 SAP 进程失败: {e}")


def parse_sap_list_report_to_dataframe(file_path):
    """
    解析 SAP 导出的非标准固定宽度文本 (Temporary_StockTurnover.txt)，
    并返回一个干净的 Pandas DataFrame。已修正为 'gbk' 编码。
    """
    try:
        # 关键修正：使用 'gbk' 编码读取 ANSI 文件
        with open(file_path, 'r', encoding='gbk') as f:
            content = f.read()
    except Exception as e:
        print(f"❌ 错误：无法读取 TXT 文件。请确认路径和编码。{e}")
        return None

    # 正则表达式用于匹配数据行 (保留 '总和' 行)
    data_lines = []
    pattern = re.compile(
        r'^\s*(?P<Loc>\S+)\s+(?P<Rate>[\-\d\.\s]+)\s+'
        r'(?P<AvgVal>[\-\d\.\s]+)(RMB)?\s+'
        r'(?P<TotalVal>[\-\d\.\s]+)(RMB)?',
        re.MULTILINE
    )

    for match in pattern.finditer(content):
        data_lines.append([
            match.group('Loc'),
            match.group('Rate'),
            match.group('AvgVal'),
            match.group('TotalVal')
        ])

    df = pd.DataFrame(data_lines, columns=['库存地点', '周转评估-V', '平均估价库存值', '总使用值'])

    # 清理数据：移除空格、逗号，并转换为数字
    numeric_cols = ['周转评估-V', '平均估价库存值', '总使用值']
    for col in numeric_cols:
        df[col] = df[col].astype(str).str.replace(',', '', regex=False)
        df[col] = df[col].astype(str).str.replace(r'\s+', '', regex=True).str.strip()
        df[col] = pd.to_numeric(df[col], errors='coerce')

    return df


def write_to_google_sheet(value, sheet_url, worksheet_name, target_month_key, auth_file):
    """
    连接 Google Sheets API 并将值写入指定位置 (通过查找月份键定位)。
    目标：在 A 列查找 YYYYMM，将值写入 B 列。
    """
    try:
        # 1. 身份验证（禁用 SSL 验证以兼容公司代理）
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        scopes = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = Credentials.from_service_account_file(auth_file, scopes=scopes)
        authed_session = AuthorizedSession(creds)
        authed_session.verify = False
        gc = gspread.Client(auth=creds, session=authed_session)

        # 2. 打开工作簿和工作表
        sh = gc.open_by_url(sheet_url)
        worksheet = sh.worksheet(worksheet_name)

        # 3. 核心查找逻辑：在 'Month' (A) 列中查找匹配项
        print(f"-> 正在 Google Sheet 中查找月份键: {target_month_key}...")

        # gspread.find() 查找整个工作表，会找到 A 列的匹配项
        cell_to_find = worksheet.find(target_month_key)

        if not cell_to_find:
            print(f"❌ 查找失败：在工作表中找不到月份 '{target_month_key}'。请检查 'Month' 列数据。")
            return

        # 4. 确定目标单元格 (写入到找到的行，但固定在 B 列)
        target_row = cell_to_find.row
        target_cell = f"B{target_row}"  # <<< 修正：写入到 B 列

        # 5. 写入数据
        worksheet.update(range_name=target_cell, values=[[value]])

        print(f"✅ 成功将值 '{value}' 写入 Google Sheet: {worksheet_name}!{target_cell}")

    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作表 '{worksheet_name}'。")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作簿。请务必确认已将服务账户邮箱添加为该表格的 '编辑者'。")
    except NameError:
        print("❌ 无法执行 Google Sheets 写入，gspread 库未导入。")
    except Exception as e:
        print(f"❌ Google Sheets 写入失败。错误: {e}")


def run_sap_automation(start_period, end_period, sap_compatible_dir, file_name_only):
    """连接到 SAP 会话并执行 MC.7 操作"""

    # 1. 确保输出目录存在
    if not os.path.exists(OUTPUT_DIR):
        try:
            os.makedirs(OUTPUT_DIR)
            print(f"创建目录: {OUTPUT_DIR}")
        except Exception as e:
            print(f"致命错误: 无法创建输出目录。{e}")
            return

    # 2. 清理旧文件以避免覆盖提示 (解决 Error Saving the List)
    windows_path = os.path.join(OUTPUT_DIR, file_name_only)
    if os.path.exists(windows_path):
        try:
            os.remove(windows_path)
            print(f"清除了旧文件: {windows_path}")
        except Exception as e:
            print(f"错误: 无法删除旧的导出文件。请确保文件未被占用。{e}")
            return  # 无法安全写入，则终止自动化

    # --- 3. SAP 连接 (保持不变) ---
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print(f"错误: 无法连接到 SAP GUI Scripting Engine 或找不到活动会话。请确保 SAP GUI 已登录。{e}")
        return

    print(f"成功连接到 SAP 会话。正在执行 MC.7...")

    # --- 4. MC.7 事务码操作代码 ---
    session.findById("wnd[0]/tbar[0]/okcd").text = "mc.7"
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]").resizeWorkingPane(269, 29, False)
    session.findById("wnd[0]").sendVKey(0)
    session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").text = "CN15"
    session.findById("wnd[0]/usr/ctxtSL_WERKS-LOW").caretPosition = 4
    session.findById("wnd[0]/usr/btn%_SL_MTART_%_APP_%-VALU_PUSH").press()
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "ERSA"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "ZERS"
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
    session.findById(
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 4
    session.findById("wnd[1]/tbar[0]/btn[8]").press()
    session.findById("wnd[0]/usr/ctxtSL_SPMON-LOW").text = start_period
    session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").text = end_period
    session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").setFocus()
    session.findById("wnd[0]/usr/ctxtSL_SPMON-HIGH").caretPosition = 7
    session.findById("wnd[0]/tbar[1]/btn[8]").press()
    session.findById("wnd[0]/tbar[1]/btn[20]").press()

    # --- 5. 文件导出流程 ---
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select()
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    session.findById("wnd[1]/usr/ctxtDY_PATH").text = sap_compatible_dir
    session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name_only
    session.findById("wnd[1]/tbar[0]/btn[0]").press()
    time.sleep(1)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        pass

    session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
    session.findById("wnd[0]").sendVKey(0)
    time.sleep(1)
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
    except Exception:
        pass

    print(f"自动化流程执行完成，文件已保存到: {os.path.join(OUTPUT_DIR, file_name_only)}")


# --- 主执行函数 ---
def main_automation_process():
    """主执行函数：计算期间，启动 SAP，执行操作，并转换文件格式并写入 Google Sheet"""

    # 1. 计算期间
    today = datetime.date.today()
    end_period = today.strftime("%m/%Y")
    start_date = today.replace(year=today.year - 1)
    start_period = start_date.strftime("%m/%Y")

    # 2. 构建文件导出路径
    sap_compatible_dir = OUTPUT_DIR.replace('\\', '/')
    file_name_only = OUTPUT_FILENAME

    print(f"动态期间计算成功: 从 {start_period} 到 {end_period}")

    # 3. 启动 SAP
    try:
        subprocess.Popen([r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe',
                          f'-system={SAP_SYSTEM}', f'-client={SAP_CLIENT}',
                          f'-user={SAP_USER}', f'-pw={SAP_PASSWORD}',
                          f'-language={SAP_LANGUAGE}'])
        print("SAP GUI 启动成功。")
    except Exception as e:
        print(f"错误：SAP GUI 启动失败或路径错误。{e}")
        return
    time.sleep(15)

    # 4. 执行 SAP 自动化操作
    run_sap_automation(start_period, end_period, sap_compatible_dir, file_name_only)

    # ***************************************************************
    # 5. TXT/XLSX 转换、数据提取和 Google Sheets 写入
    # ***************************************************************
    txt_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    xlsx_filename = OUTPUT_FILENAME.replace(".txt", ".xlsx")
    xlsx_file_path = os.path.join(OUTPUT_DIR, xlsx_filename)

    print(f"\n--- 正在处理数据 ---")

    # 步骤 5.1: 读取和解析 TXT 文件
    df = parse_sap_list_report_to_dataframe(txt_file_path)

    if df is not None and not df.empty:
        # 步骤 5.2: 写入 XLSX 文件 (保留，以便您在本地检查)
        try:
            if os.path.exists(xlsx_file_path): os.remove(xlsx_file_path)

            df.to_excel(xlsx_file_path, index=False, engine='openpyxl')
            print(f"✅ 数据成功转换并保存为 XLSX: {xlsx_file_path}")

            # 步骤 5.3: 【关键】提取 B2 单元格的值
            target_value = df.iloc[0]['周转评估-V']
            print(f"✅ 提取的 B2 单元格值 ('总和'的'周转评估-V') 为: {target_value}")

            # 步骤 5.4: 写入 Google Sheets
            # 构造 YYYYMM 格式的查找键
            target_month_key = today.strftime("%Y%m")

            write_to_google_sheet(
                target_value,
                GOOGLE_SHEET_URL,
                WORKSHEET_NAME,
                target_month_key,  # 传入 YYYYMM 查找键
                SERVICE_ACCOUNT_FILE
            )

            # 步骤 5.5: 清理临时 TXT 文件
            os.remove(txt_file_path)
            print(f"✅ 已删除临时 TXT 文件。")

        except Exception as e:
            print(f"❌ 数据处理/写入 XLSX 失败。错误: {e}")
    else:
        print("❌ 无法从 SAP 导出的 TXT 文件中提取数据，跳过后续操作。")

    # 6. 关闭 SAP 进程
    close_SAP()


if __name__ == "__main__":
    close_SAP()
    main_automation_process()
    input("\n程序执行完毕，按 Enter 键退出... / Press Enter to exit...")