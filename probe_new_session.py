"""【一次性探针，跑完即可删除】

验证 SAP GUI Scripting 的「新建会话」能力在当前环境（LAP / 321）是否可用。

回答四个问题：
  A. 调用 session.CreateSession() 能不能开出新会话？多久出现？
  B. 能不能用 /i 干净地关掉自己开的会话？
  C. 备选路径——命令框输入 /o——能不能用？
  D. 当前环境允许同时开几个会话（上限）？

安全说明：
  * 不执行任何业务事务，不改任何数据
  * 只开一个会话、看一眼、关掉
  * 不动已有的会话

用法：直接运行即可。SAP 没开着的话脚本会自己启动它。
"""

import os
import sys
import time

import win32com.client

SEP = "=" * 62


def session_ids(conn):
    """返回当前所有会话的 [(索引, 会话ID, 当前事务), ...]。

    不依赖 Children.Count —— pywin32 动态派发下该属性未必可用，
    所以用「一直索引到抛异常为止」的退化方案。
    """
    out = []
    for i in range(30):
        try:
            s = conn.Children(i)
        except Exception:
            break
        try:
            txn = s.Info.Transaction
        except Exception:
            txn = "?"
        out.append((i, s.Id, txn))
    return out


def dump(conn, label):
    rows = session_ids(conn)
    print(f"{label}（{len(rows)} 个会话）")
    for i, sid, txn in rows:
        print(f"    [{i}] id={sid}  事务={txn}")
    return rows


def wait_for_new(conn, known_ids, timeout):
    """轮询等新会话出现，返回 (会话对象, 耗时秒)；超时返回 (None, 耗时)。"""
    t0 = time.time()
    while time.time() - t0 < timeout:
        time.sleep(0.5)
        for i in range(30):
            try:
                s = conn.Children(i)
            except Exception:
                break
            if s.Id not in known_ids:
                return s, time.time() - t0
    return None, time.time() - t0


def main():
    print(SEP)
    print("SAP 新会话能力探针")
    print(SEP)

    # ── 0. 连接（SAP 没开着就自己启动）────────────────────────────────────
    we_started_it = False
    try:
        app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
        print("检测到 SAP 已在运行，直接使用。")
    except Exception as e:
        print(f"没有检测到正在运行的 SAP（{type(e).__name__}），正在启动...")
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        try:
            from sap_common import start_sap, wait_for_sap_session
            start_sap()
            print("等待 SAP 会话就绪...")
            wait_for_sap_session()
            we_started_it = True
            app = win32com.client.GetObject("SAPGUI").GetScriptingEngine
            print("SAP 已启动。")
        except Exception as e2:
            print(f"❌ 启动 SAP 失败: {type(e2).__name__}: {e2}")
            return

    try:
        print(f"GUI 版本: {app.MajorVersion}.{app.MinorVersion}.{app.PatchLevel}")
    except Exception as e:
        print(f"GUI 版本: (读不到: {type(e).__name__})")

    try:
        conn = app.Children(0)
    except Exception as e:
        print(f"❌ 取不到连接: {type(e).__name__}: {e}")
        return

    try:
        print(f"连接描述: {conn.Description}")
    except Exception:
        pass

    print()
    before = dump(conn, "开始时")
    before_ids = {sid for _, sid, _ in before}
    base = conn.Children(max(0, len(before) - 1))

    results = {}

    # ── A. CreateSession ───────────────────────────────────────────────────
    print(f"\n{SEP}\n【A】session.CreateSession()\n{SEP}")
    call_err = None
    try:
        base.CreateSession()
        print("  调用已发出（异步方法，不返回新会话对象，需轮询）")
    except Exception as e:
        call_err = e
        print(f"  ❌ 调用抛异常: {type(e).__name__}: {e}")

    new_sess, waited = (None, 0.0)
    if call_err is None:
        new_sess, waited = wait_for_new(conn, before_ids, timeout=25)
        if new_sess is None:
            print(f"  ❌ 等了 {waited:.1f} 秒，新会话没出现")
        else:
            print(f"  ✅ 新会话出现，耗时 {waited:.1f} 秒")
    results["A"] = new_sess is not None

    # ── B. 用 /i 关掉它 ────────────────────────────────────────────────────
    print(f"\n{SEP}\n【B】用 /i 关掉自己开的会话\n{SEP}")
    if new_sess is None:
        print("  （跳过：A 没成功）")
        results["B"] = None
    else:
        try:
            new_sess.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            new_sess.findById("wnd[0]").sendVKey(0)
            time.sleep(2.0)
            rows = session_ids(conn)
            if len(rows) == len(before):
                print(f"  ✅ 已关闭，会话数回到 {len(rows)}")
                results["B"] = True
            else:
                print(f"  ⚠️ 会话数 {len(rows)}，期望 {len(before)} —— 可能没关干净")
                results["B"] = False
        except Exception as e:
            print(f"  ❌ 关闭失败: {type(e).__name__}: {e}")
            results["B"] = False

    # ── C. 备选：命令框 /o ─────────────────────────────────────────────────
    print(f"\n{SEP}\n【C】命令框输入 /o（备选方案）\n{SEP}")
    rows_now = session_ids(conn)
    ids_now = {sid for _, sid, _ in rows_now}
    alt = None
    if not rows_now:
        print("  （跳过：没有可用会话）")
    else:
        cur = conn.Children(rows_now[-1][0])
        try:
            cur.findById("wnd[0]/tbar[0]/okcd").text = "/o"
            cur.findById("wnd[0]").sendVKey(0)
            print("  已发送 /o")
            alt, took = wait_for_new(conn, ids_now, timeout=25)
            if alt is None:
                print(f"  ❌ 等了 {took:.1f} 秒，没有新会话")
                print("     （可能 SAP 弹出了「会话列表」对话框挡着——若有，请手动关掉它）")
            else:
                print(f"  ✅ 新会话出现，耗时 {took:.1f} 秒")
        except Exception as e:
            print(f"  ❌ 发送 /o 失败: {type(e).__name__}: {e}")
    results["C"] = alt is not None

    if alt is not None:
        try:
            alt.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            alt.findById("wnd[0]").sendVKey(0)
            time.sleep(2.0)
            print(f"  已关掉，当前会话数: {len(session_ids(conn))}")
        except Exception as e:
            print(f"  ⚠️ 关闭失败: {e}")

    # ── D. 会话上限 ────────────────────────────────────────────────────────
    print(f"\n{SEP}\n【D】试探同时能开几个会话（上限）\n{SEP}")
    print("  逐个开、数数、再全部关掉。若中途报错就说明到上限了。")
    print("  最多额外开 4 个（避免撞上你们的上限后关不掉）。")
    opened = []
    limit_hit = None
    try:
        for n in range(1, 5):
            rows_now = session_ids(conn)
            cur = conn.Children(rows_now[-1][0])
            ids_now = {sid for _, sid, _ in rows_now}
            cur.CreateSession()
            s, took = wait_for_new(conn, ids_now, timeout=15)
            if s is None:
                limit_hit = f"第 {n} 个没开出来"
                break
            opened.append(s)
            print(f"  · 已开 {len(opened)} 个额外会话（本次耗时 {took:.1f}s）")
    except Exception as e:
        limit_hit = f"{type(e).__name__}: {e}"
        print(f"  ⚠️ 第 {len(opened) + 1} 个时报错: {limit_hit}")

    print(f"  → 至少能额外开 {len(opened)} 个" + (f"（受限：{limit_hit}）" if limit_hit else "（未触及上限）"))

    print("  正在关掉刚才开的会话...")
    for s in opened:
        try:
            s.findById("wnd[0]/tbar[0]/okcd").text = "/i"
            s.findById("wnd[0]").sendVKey(0)
            time.sleep(0.8)
        except Exception:
            pass
    time.sleep(1.5)

    # ── 收尾 ───────────────────────────────────────────────────────────────
    print()
    dump(conn, "结束时")

    print(f"\n{SEP}\n结论\n{SEP}")
    print(f"  A. CreateSession 可用 : {'✅ 是' if results['A'] else '❌ 否'}"
          + (f"（耗时 {waited:.1f} 秒）" if results['A'] else ""))
    print(f"  B. /i 能干净关闭      : "
          + {True: "✅ 是", False: "❌ 否", None: "— 未测"}[results["B"]])
    print(f"  C. /o 备选可用        : {'✅ 是' if results['C'] else '❌ 否'}")
    print(f"  D. 额外会话上限       : 至少 {len(opened)} 个")

    if we_started_it:
        print("\n（本次是探针自己启动的 SAP，正在关闭）")
        try:
            from sap_common import close_sap
            close_sap()
        except Exception as e:
            print(f"  关闭失败，可手动关：{e}")

    print(f"\n请把以上全部输出发给 Claude。")


if __name__ == "__main__":
    main()
