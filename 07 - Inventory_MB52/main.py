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

OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\Inventory_MB52"
OUTPUT_FILENAME = "Temporary_Inventory.xlsx"

GOOGLE_SHEET_ID = '1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs'
WORKSHEET_NAME = 'MasterData'
SERVICE_ACCOUNT_FILE = get_resource_path('../pyreadsp-b5b9c1909de6.json')

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


def get_inventory_mb52():
    output_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    if os.path.exists(output_file_path):
        try:
            os.remove(output_file_path)
            print(f"已删除旧文件: {output_file_path}")
        except Exception as e:
            print(f"警告: 无法删除旧文件 {output_file_path}: {e}")
    
    session = None
    for attempt in range(12):
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"[{ts}] 尝试连接 SAP GUI ({attempt + 1}/12)...")
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"[{ts}] ✅ 连接成功")
            break
        except Exception as e:
            ts = datetime.datetime.now().strftime("%H:%M:%S")
            print(f"[{ts}] ❌ 连接失败: {e}")
            if attempt < 11:
                print(f"      等待 5 秒后重试...")
                time.sleep(5)
    if session is None:
        print("错误: 无法连接到 SAP GUI Scripting Engine，已等待60秒，请确保 SAP GUI 已登录。")
        return
    
    print("成功连接到 SAP 会话。正在执行 MB52...")
    
    try:
        session.findById("wnd[0]").maximize()
        
        session.findById("wnd[0]/tbar[0]/okcd").text = "MB52"
        session.findById("wnd[0]").sendVKey(0)
        
        session.findById("wnd[0]/usr/chkNOZERO").selected = True
        
        session.findById("wnd[0]/usr/ctxtWERKS-LOW").text = "CN15"
        
        session.findById("wnd[0]/usr/ctxtLGORT-LOW").text = "MECH"
        
        session.findById("wnd[0]/usr/ctxtP_VARI").text = "/KEL"
        session.findById("wnd[0]/usr/ctxtP_VARI").setFocus()
        session.findById("wnd[0]/usr/ctxtP_VARI").caretPosition = 4
        
        session.findById("wnd[0]/tbar[1]/btn[8]").press()
        time.sleep(2)
        
        session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[1]").select()
        time.sleep(1)
        
        for wait in range(10):
            try:
                session.findById("wnd[1]/usr/ctxtDY_PATH")
                break
            except Exception:
                time.sleep(1)
        
        session.findById("wnd[1]/usr/ctxtDY_PATH").text = OUTPUT_DIR
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = OUTPUT_FILENAME
        session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = len(OUTPUT_FILENAME)
        
        session.findById("wnd[1]/tbar[0]/btn[0]").press()
        
        print(f"✅ MB52 库存数据导出完成")
        print(f"文件保存路径: {os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)}")

        time.sleep(3)

    except Exception as e:
        print(f"❌ 执行 SAP 操作时发生错误: {e}")
        import traceback
        traceback.print_exc()


def determine_process(material_code):
    if len(material_code) >= 5 and material_code[4] == 'N':
        return 'IM'
    return ''


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
        
        today_date = datetime.date.today().strftime("%Y-%m-%d")
        
        processed_data = []
        for _, row in df.iterrows():
            material = str(row.iloc[0]) if pd.notna(row.iloc[0]) else ''
            
            if not material or material.strip() == '':
                continue
            
            material_desc = str(row.iloc[1]) if pd.notna(row.iloc[1]) else ''
            unrestricted_stock = row.iloc[2] if pd.notna(row.iloc[2]) else ''
            unrestricted_value = row.iloc[3] if pd.notna(row.iloc[3]) else ''
            
            process = determine_process(material)
            
            if process != 'IM':
                continue
            
            processed_data.append([
                material,
                material_desc,
                unrestricted_stock,
                unrestricted_value,
                today_date,
                process
            ])
        
        print(f"处理了 {len(processed_data)} 行数据")
        
        print("正在连接 Google Sheets...")
        gc = gspread.service_account(filename=auth_file)
        sh = gc.open_by_key(sheet_id)
        worksheet = sh.worksheet(worksheet_name)
        
        existing_data = worksheet.get_all_values()
        if len(existing_data) > 1:
            existing_records = {}
            for idx, row in enumerate(existing_data[1:], start=2):
                if len(row) >= 5 and row[0]:
                    key = f"{row[0]}_{row[4]}"
                    existing_records[key] = idx
        else:
            existing_records = {}
        
        new_records = []
        update_ranges = []
        updated_count = 0
        added_count = 0
        
        for data_row in processed_data:
            key = f"{data_row[0]}_{data_row[4]}"
            
            if key in existing_records:
                row_index = existing_records[key]
                range_name = f"A{row_index}:F{row_index}"
                update_ranges.append({
                    'range': range_name,
                    'values': [data_row]
                })
                updated_count += 1
            else:
                new_records.append(data_row)
                added_count += 1
        
        if update_ranges:
            print(f"正在批量更新 {len(update_ranges)} 条已存在记录...")
            worksheet.batch_update(update_ranges, value_input_option='USER_ENTERED')
        
        if new_records:
            print(f"正在添加 {len(new_records)} 条新记录到 Google Sheets...")
            worksheet.append_rows(new_records, value_input_option='USER_ENTERED')
        
        print(f"✅ Google Sheets 同步完成")
        print(f"   - 新增记录: {added_count} 条")
        print(f"   - 已存在记录: {updated_count} 条")
        
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
        get_inventory_mb52()
        
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
