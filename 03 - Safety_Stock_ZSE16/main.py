import warnings
warnings.filterwarnings('ignore', category=FutureWarning)

import subprocess
import time
import win32com.client
import sys
import os
import pandas as pd
import openpyxl
import gspread
import requests
import urllib3

def get_resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\Safety_Stock"
OUTPUT_FILENAME = "Temporary_File.XLSX"

GOOGLE_SHEET_ID = '1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs'
WORKSHEET_NAME = '安全库存数据'
SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')

def sap_auto_logo():
    subprocess.check_call(['C:\\Program Files (x86)\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe', '-system=CAP', '-client=321',
                           '-user=USERNAME', '-pw=PASSWORD', '-language=ZH'])
    time.sleep(15)
    print("sap open successfully")


def close_SAP():
    try:
        os.system('taskkill /im saplogon.exe /t /f')
        print("SAP GUI 进程已成功关闭。")
    except Exception as e:
        print(f"关闭 SAP 进程失败: {e}")


def get_safety_stock_zse16():
    output_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    if os.path.exists(output_file_path):
        try:
            os.remove(output_file_path)
            print(f"已删除旧文件: {output_file_path}")
        except Exception as e:
            print(f"警告: 无法删除旧文件 {output_file_path}: {e}")
    
    try:
        SapGuiAuto = win32com.client.GetObject("SAPGUI")
        application = SapGuiAuto.GetScriptingEngine
        connection = application.Children(0)
        session = connection.Children(0)
    except Exception as e:
        print(f"错误: 无法连接到 SAP GUI Scripting Engine 或找不到活动会话。请确保 SAP GUI 已登录。{e}")
        return
    
    print("成功连接到 SAP 会话。正在执行 ZSE16...")
    
    try:
        session.findById("wnd[0]").resizeWorkingPane(232, 29, False)
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "ZSE16"
        session.findById("wnd[0]").sendVKey(0)
        
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").text = "MARC"
        session.findById("wnd[0]/usr/ctxtDATABROWSE-TABLENAME").caretPosition = 4
        
        session.findById("wnd[0]/tbar[1]/btn[7]").press()
        
        session.findById("wnd[0]/usr/btn%_I1_%_APP_%-VALU_PUSH").press()
        
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "E185*"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "Z185*"
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").setFocus()
        session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").caretPosition = 5
        
        session.findById("wnd[1]/tbar[0]/btn[8]").press()
        
        session.findById("wnd[0]/usr/ctxtI2-LOW").text = "CN15"
        
        session.findById("wnd[0]/usr/txtMAX_SEL").text = ""
        session.findById("wnd[0]/usr/txtMAX_SEL").setFocus()
        session.findById("wnd[0]/usr/txtMAX_SEL").caretPosition = 11
        
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)
        
        session.findById("wnd[0]/tbar[1]/btn[43]").press()
        time.sleep(2)
        
        save_wnd = None
        try:
            session.findById("wnd[1]/usr/ctxtDY_PATH")
            save_wnd = "wnd[1]"
        except:
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            time.sleep(3)
            for wnd_id in ["wnd[1]", "wnd[2]"]:
                try:
                    session.findById(f"{wnd_id}/usr/ctxtDY_PATH")
                    save_wnd = wnd_id
                    break
                except:
                    continue
        
        if save_wnd is None:
            raise Exception("无法找到文件保存对话框 (wnd[1] 或 wnd[2])")
        
        session.findById(f"{save_wnd}/usr/ctxtDY_PATH").text = OUTPUT_DIR
        session.findById(f"{save_wnd}/usr/ctxtDY_FILENAME").text = OUTPUT_FILENAME
        session.findById(f"{save_wnd}/usr/ctxtDY_FILENAME").caretPosition = len(OUTPUT_FILENAME)
        
        session.findById(f"{save_wnd}/tbar[0]/btn[0]").press()
        time.sleep(2)
        
        try:
            session.findById(f"{save_wnd}/tbar[0]/btn[0]").press()
            print("已关闭确认窗口")
        except:
            print("确认窗口已自动关闭或不存在，继续执行...")
        
        print(f"✅ ZSE16 安全库存数据导出完成")
        print(f"文件保存路径: {os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)}")

        time.sleep(3)

    except Exception as e:
        print(f"❌ 执行 SAP 操作时发生错误: {e}")
        import traceback
        traceback.print_exc()


def upload_to_google_sheets(excel_file_path, sheet_id, worksheet_name, auth_file):
    try:
        print(f"正在读取 Excel 文件: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"❌ 错误: 找不到文件 {excel_file_path}")
            return False
        
        df = pd.read_excel(excel_file_path, engine='openpyxl')
        
        if df.empty:
            print("⚠️ 警告: Excel 文件中没有数据")
            return False
        
        print(f"读取到 {len(df)} 行数据")
        
        print("正在连接 Google Sheets...")
        urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)
        
        session = requests.Session()
        session.verify = False
        
        from google.auth.transport.requests import AuthorizedSession
        from google.oauth2.service_account import Credentials
        
        scopes = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
        credentials = Credentials.from_service_account_file(auth_file, scopes=scopes)
        authed_session = AuthorizedSession(credentials)
        authed_session.verify = False
        
        gc = gspread.Client(auth=credentials, session=authed_session)
        sh = gc.open_by_key(sheet_id)
        worksheet = sh.worksheet(worksheet_name)
        
        print("正在清空工作表...")
        worksheet.clear()
        
        header = df.columns.tolist()
        data_rows = df.values.tolist()
        
        all_data = [header] + data_rows
        
        print(f"正在上传 {len(all_data)} 行数据（包含表头）到 Google Sheets...")
        worksheet.update(all_data, 'A1', value_input_option='USER_ENTERED')
        
        print(f"✅ Google Sheets 上传完成")
        print(f"   - 表头: 1 行")
        print(f"   - 数据: {len(data_rows)} 行")
        
        return True
        
    except FileNotFoundError:
        print(f"❌ 错误: 找不到文件 {excel_file_path}")
        return False
    except gspread.exceptions.WorksheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作表 '{worksheet_name}'。")
        return False
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"❌ Google Sheets 错误: 找不到工作簿。请确认已将服务账户邮箱添加为该表格的 '编辑者'。")
        return False
    except Exception as e:
        print(f"❌ Google Sheets 上传失败。错误: {e}")
        import traceback
        traceback.print_exc()
        return False


if __name__ == "__main__":
    try:
        print("正在清理可能存在的 SAP 进程...")
        close_SAP()
        time.sleep(2)
        
        print("正在启动 SAP GUI...")
        sap_auto_logo()
        
        print("等待 SAP GUI 完全加载...")
        time.sleep(5)
        
        print("开始执行 SAP 操作...")
        get_safety_stock_zse16()
        
        print("\nSAP 操作完成，正在关闭 SAP...")
        time.sleep(2)
        close_SAP()
        
        print("\n" + "="*50)
        print("开始将 Excel 数据上传到 Google Sheets...")
        print("="*50)
        
        excel_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
        time.sleep(2)
        
        upload_to_google_sheets(
            excel_file_path=excel_file_path,
            sheet_id=GOOGLE_SHEET_ID,
            worksheet_name=WORKSHEET_NAME,
            auth_file=SERVICE_ACCOUNT_FILE
        )
        
        print("\n" + "="*50)
        print("✅ 所有操作已完成！")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 程序执行过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        close_SAP()
    finally:
        print("\n按任意键退出...")
        input()
