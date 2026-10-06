"""SAP GUI 自动化的公共部分。

原先 get_resource_path / sap_auto_logo / close_SAP 以及"取会话"那段代码，
在 9 个 main.py 里各有一份副本。2026-10-06 抽出到这里。

各脚本的用法：

    import os, sys
    sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    from sap_common import (get_resource_path, ensure_sap_running,
                            open_session, close_session, close_sap)

    SERVICE_ACCOUNT_FILE = get_resource_path('pyreadsp-b5b9c1909de6.json')

    ...
    we_started_sap = ensure_sap_running()   # SAP 没开就启动，开了就复用
    session = open_session()                # 新开一个自己的窗口（~3 秒）
    try:
        ...干活...
    finally:
        close_session(session)              # 关掉自己的窗口
        if we_started_sap:
            close_sap()                     # 只有自己启动的才关

注意两点：

* get_resource_path 的基准目录是**仓库根**，不要再写 '../'
* 全流程**不关别人的 SAP 会话** —— 只在最后关掉"自己启动的那一个 SAP 进程"
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

    会关掉机器上**所有** SAP 会话，包括别人正在用的。
    只在"SAP 是本脚本启动的"时候才该调用（见 ensure_sap_running）。
    """
    result = subprocess.run(
        ['taskkill', '/im', 'saplogon.exe', '/t', '/f'],
        capture_output=True, text=True,
    )
    if result.returncode == 0:
        print("SAP GUI 进程已关闭。")
    else:
        print("机器上没有正在运行的 SAP 进程，跳过。")


# ── 会话级操作 ────────────────────────────────────────────────────────────────
# 2026-10-07 起，各脚本不再"重启 SAP"，而是在已运行的 SAP 里新开一个会话
# （约 3 秒，对比重启 SAP 的 8~15 秒），干完活再关掉自己那个窗口。
# 这样既不碰别人的会话，脚本之间也天然隔离。


def _iter_sessions(conn):
    """遍历一个连接下的所有会话。

    不依赖 Children.Count —— pywin32 动态派发下该属性未必可用，
    所以用"一直索引到抛异常为止"的退化方案。
    """
    for i in range(30):
        try:
            yield conn.Children(i)
        except Exception:
            return


def _get_connection():
    app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
    return app.Children(0)


def ensure_sap_running():
    """确保 SAP 在运行。没运行就启动它。

    返回 True 表示 SAP 是本次调用启动的 —— 调用方收尾时应当把它关掉，
    把机器还原成原来的样子。返回 False 表示 SAP 本来就在跑，别动它。

    有了这个判断，脚本不需要环境变量就能同时支持两种模式：
      * 被 run_all.ps1 调用 —— SAP 已由编排器启动，各脚本复用
      * 手动单独运行       —— 脚本自己启动 SAP，跑完自己关掉
    """
    try:
        win32com.client.GetObject("SAPGUI")
        return False
    except Exception:
        print("没有检测到正在运行的 SAP，正在启动...")
        start_sap()
        print("等待 SAP 会话就绪...")
        wait_for_sap_session()
        return True


def open_session(timeout=30.0, interval=0.5):
    """在当前 SAP 连接里新开一个会话并返回它。

    CreateSession() 是异步的，而且**不返回**新会话对象，
    所以这里用"记录已有会话 ID → 调用 → 轮询比对"的方式拿到它。
    """
    conn = _get_connection()
    sessions = list(_iter_sessions(conn))
    if not sessions:
        raise RuntimeError("当前 SAP 连接里没有可用会话")

    if len(sessions) > 3:
        print(f"  [SAP] ⚠️ 当前已有 {len(sessions)} 个会话，可能有上次残留的。")
        print(f"  [SAP]    本脚本不会去关别人的窗口，必要时请手动清理。")

    known = {s.Id for s in sessions}
    sessions[-1].CreateSession()

    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(interval)
        for s in _iter_sessions(conn):
            if s.Id not in known:
                print(f"  [SAP] 已在会话 {s.Id} 中打开")
                return s
    raise TimeoutError(f"新会话在 {timeout} 秒内未出现")


def close_session(session):
    """用 /i 关掉自己开的那个会话。

    收尾用途，失败只提示不抛异常 —— 不要因为关不干净而让整个脚本报错。
    """
    if session is None:
        return
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "/i"
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(0.5)
        print("  [SAP] 已关闭本次会话")
    except Exception as e:
        print(f"  [SAP] 关闭会话时出错（可手动关掉那个窗口）: {e}")


if __name__ == "__main__":
    # 供 run_all.ps1 调用，让编排器负责 SAP 的启停：
    #     python sap_common.py start    循环前启动一次
    #     python sap_common.py close    循环后关闭一次
    _cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if _cmd == "start":
        ensure_sap_running()
    elif _cmd == "close":
        close_sap()
    else:
        print("用法: python sap_common.py start|close")
        sys.exit(2)
