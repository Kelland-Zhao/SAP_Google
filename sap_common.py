"""SAP GUI 自动化的公共部分。

原先 get_resource_path / sap_auto_logo / close_SAP 以及"取会话"那段代码，
在 9 个 main.py 里各有一份副本。2026-10-06 抽出到这里。

各脚本的用法：

    import os, sys
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from sap_common import get_resource_path, start_sap, wait_for_sap_session, close_sap

    SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')

注意 get_resource_path 的基准目录是**仓库根**，不要再写 '../'。
"""

import os
import subprocess
import sys
import time

import win32com.client

SAPSHCUT = r'C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapshcut.exe'
SAP_SYSTEM = 'LAP'
SAP_CLIENT = '321'
SAP_LANGUAGE = 'ZH'


def get_resource_path(relative_path):
    """解析资源路径。基准目录是仓库根（本文件所在目录）。

    保留 sys._MEIPASS 分支是为了兼容早期 PyInstaller 打包的产物；
    现在直接跑 .py，该分支不会命中。
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


def start_sap():
    """启动 SAP GUI。只负责启动，不等待就绪 —— 就绪判断交给 wait_for_sap_session。

    这里的 -user / -pw 是有意的占位符：公司电脑上的 SAP 走 SSO 登录，
    这两个参数不生效。详见 README「凭据」一节。
    """
    subprocess.check_call([
        SAPSHCUT,
        f'-system={SAP_SYSTEM}', f'-client={SAP_CLIENT}',
        '-user=USERNAME', '-pw=PASSWORD',
        f'-language={SAP_LANGUAGE}',
    ])


def wait_for_sap_session(timeout=60, interval=1.0):
    """轮询等待 SAP GUI Scripting 会话可用，返回 session。

    取代原先"启动后盲等 15 秒"的写法：SAP 实际 6 秒就绪就只等 6 秒，
    需要 20 秒就等到 20 秒 —— 又快又不会等不够。

    超时抛 TimeoutError，由各脚本的 __main__ 捕获后以退出码 1 结束。
    """
    deadline = time.time() + timeout
    attempt = 0
    last_err = None

    while True:
        attempt += 1
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            connection = application.Children(0)
            session = connection.Children(0)
            print(f"  [SAP] 第 {attempt} 次尝试连接成功")
            return session
        except Exception as e:
            last_err = e
            if time.time() >= deadline:
                raise TimeoutError(
                    f"SAP 会话在 {timeout} 秒内未就绪（已尝试 {attempt} 次）。"
                    f"最后一次错误: {last_err}"
                )
            time.sleep(interval)


def close_sap():
    """强制关闭 SAP GUI 进程。

    会关掉机器上**所有** SAP 会话，包括别人正在用的 —— 见 README
    「运行期间的重要注意事项」。
    """
    result = subprocess.run(
        ['taskkill', '/im', 'saplogon.exe', '/t', '/f'],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        print("SAP GUI 进程已关闭。")
    else:
        print("机器上没有正在运行的 SAP 进程，跳过。")
