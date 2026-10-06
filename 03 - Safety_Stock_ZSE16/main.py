import warnings
warnings.filterwarnings('ignore', category=FutureWarning)

import time
import win32com.client
import sys
import os
import io
import pandas as pd
import gspread
import requests
import urllib3

# 公共部分在仓库根的 sap_common.py，不在本脚本所在目录。
# 先把仓库根加进 sys.path，这样单个脚本仍然可以独立运行。
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from sap_common import get_resource_path, start_sap, wait_for_sap_session, close_sap

OUTPUT_DIR = r"O:\My Drive\071 - SAP 数据\Safety_Stock"
OUTPUT_FILENAME = "Temporary_File.txt"

GOOGLE_SHEET_ID = '1hVHBdnK_EVSMW54meCpx91rooIZ6Y8vICQzG7txVHGs'
WORKSHEET_NAME = '安全库存数据'

# 仅上传以下列，键为 SAP 字段名，值为 Google Sheets 中文表头
COLUMN_MAPPING = {'MATNR': '物料', 'EISBE': '安全库存'}
SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')


def get_safety_stock_zse16(session):
    output_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
    if os.path.exists(output_file_path):
        try:
            os.remove(output_file_path)
            print(f"已删除旧文件: {output_file_path}")
        except Exception as e:
            print(f"警告: 无法删除旧文件 {output_file_path}: {e}")
    
    print("正在执行 ZSE16...")
    
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
        
        session.findById("wnd[0]/mbar/menu[1]/menu[5]").select()
        time.sleep(2)
        
        radio_tab_delimited = "wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]"
        try:
            session.findById(radio_tab_delimited).select()
            session.findById(radio_tab_delimited).setFocus()
            session.findById("wnd[1]/tbar[0]/btn[0]").press()
            print("已选择导出格式: 制表符分隔文本")
            time.sleep(2)
        except Exception:
            print("未出现格式选择窗口，继续执行...")
        
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
        raise


def read_sap_text_export(file_path):
    with open(file_path, 'rb') as f:
        raw = f.read()

    if raw[:2] in (b'\xff\xfe', b'\xfe\xff'):
        encoding = 'utf-16'
    else:
        try:
            raw.decode('utf-8')
            encoding = 'utf-8-sig'
        except UnicodeDecodeError:
            encoding = 'gbk'
    text = raw.decode(encoding, errors='replace')
    print(f"使用编码 {encoding} 解析")

    # SAP 导出前几行为「表：」「显示的字段：」等元数据，表头行以制表符开头
    lines = text.splitlines()
    header_idx = next((i for i, line in enumerate(lines) if line.startswith('\t')), 0)

    df = pd.read_csv(io.StringIO(text), sep='\t', dtype=str, skiprows=header_idx,
                     skip_blank_lines=True, engine='python')

    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, ~df.columns.str.startswith('Unnamed')]
    df = df.apply(lambda col: col.str.strip() if col.dtype == object else col)
    # SAP 文本导出的负数为尾随负号（如 123.45-），转换为标准写法
    df = df.replace(r'^(-?[\d.,]+)-$', r'-\1', regex=True)
    df = df.fillna('')
    return df


def upload_to_google_sheets(excel_file_path, sheet_id, worksheet_name, auth_file):
    try:
        print(f"正在读取导出文件: {excel_file_path}")
        if not os.path.exists(excel_file_path):
            print(f"❌ 错误: 找不到文件 {excel_file_path}")
            return False
        
        df = read_sap_text_export(excel_file_path)
        
        if df.empty:
            print("⚠️ 警告: 导出文件中没有数据，跳过上传（不算失败）")
            return True  # 空数据不是故障，让 run_all 记为成功
        
        print(f"读取到 {len(df)} 行数据")
        
        missing = [col for col in COLUMN_MAPPING if col not in df.columns]
        if missing:
            print(f"❌ 错误: 导出文件缺少所需列 {missing}，实际列: {list(df.columns)}")
            return False
        
        df = df[list(COLUMN_MAPPING)].rename(columns=COLUMN_MAPPING)
        
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
        close_sap()
        time.sleep(2)
        
        print("正在启动 SAP GUI...")
        start_sap()
        
        print("等待 SAP 会话就绪...")
        session = wait_for_sap_session()

        print("开始执行 SAP 操作...")
        get_safety_stock_zse16(session)
        
        print("\nSAP 操作完成，正在关闭 SAP...")
        time.sleep(2)
        close_sap()
        
        print("\n" + "="*50)
        print("开始将导出数据上传到 Google Sheets...")
        print("="*50)
        
        excel_file_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME)
        time.sleep(2)
        
        if not upload_to_google_sheets(
            excel_file_path=excel_file_path,
            sheet_id=GOOGLE_SHEET_ID,
            worksheet_name=WORKSHEET_NAME,
            auth_file=SERVICE_ACCOUNT_FILE
        ):
            close_sap()
            sys.exit(1)
        
        print("\n" + "="*50)
        print("✅ 所有操作已完成！")
        print("="*50)
        
    except Exception as e:
        print(f"❌ 程序执行过程中发生错误: {e}")
        import traceback
        traceback.print_exc()
        close_sap()
        sys.exit(1)
    finally:
        print("\n执行完毕。")
