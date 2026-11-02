"""09 - 工单自动关闭的 SAP GUI 操作。

只依赖传入的 session 对象，不 import win32com / sap_common，
所以可以在 macOS 上用假 session 跑单元测试。

操作顺序来自两份 SAP 录制脚本：
  * 工时确认IW48.txt
  * PM工单技术完成.txt
"""

import time

OKCODE = "wnd[0]/tbar[0]/okcd"
STATUS_BAR = "wnd[0]/sbar"

# 这些文本出现在状态栏时，即使消息类型是 S/W 也不能算成功。
NEGATIVE_KEYWORDS = (
    "未选择", "未选中", "没有选择", "not selected", "no orders",
)

_ORDER_POPUP = "wnd[0]/usr/btn%_AUFNR_%_APP_%-VALU_PUSH"
_CONFIRMATION_TABLE = "wnd[0]/usr/tblSAPLCORUTABCDEF_3350"
_CONFIRM_CHECKBOXES = (
    f"{_CONFIRMATION_TABLE}/chkAFRD-AUERU[9,0]",
    f"{_CONFIRMATION_TABLE}/chkAFRD-AUSOR[10,0]",
    f"{_CONFIRMATION_TABLE}/chkAFRD-LEKNW[11,0]",
)


def read_status(session):
    """读取 SAP 状态栏，返回 (message_type, text)。"""
    try:
        sbar = session.findById(STATUS_BAR)
    except Exception:
        return "", ""

    text = getattr(sbar, "text", None)
    if text is None:
        text = getattr(sbar, "Text", "")

    message_type = getattr(sbar, "messageType", None)
    if not message_type:
        message_type = getattr(sbar, "MessageType", "")

    return str(message_type or "").upper(), str(text or "")


def is_success(message_type, text):
    """S/W 成功；E/A、空、未知都失败；否定文本也失败。"""
    lowered = str(text or "").lower()
    if any(keyword.lower() in lowered for keyword in NEGATIVE_KEYWORDS):
        return False
    return str(message_type or "").upper() in {"S", "W"}


def wait_for_control(session, control_id, timeout=30.0, interval=0.5):
    """轮询等待控件出现；超时抛 TimeoutError。"""
    deadline = time.time() + timeout
    while True:
        try:
            return session.findById(control_id)
        except Exception:
            if time.time() >= deadline:
                raise TimeoutError(f"等待控件超时: {control_id}")
            time.sleep(interval)


def _press(session, control_id):
    session.findById(control_id).press()


def _set_okcode(session, code):
    session.findById(OKCODE).text = code
    session.findById("wnd[0]").sendVKey(0)


def _resize(session):
    session.findById("wnd[0]").resizeWorkingPane(232, 29, False)


def _clear_multi_select(session, push_button_id):
    _press(session, push_button_id)
    _press(session, "wnd[1]/tbar[0]/btn[16]")
    _press(session, "wnd[1]/tbar[0]/btn[8]")


def copy_to_clipboard(text):
    """把文本放进 Windows 剪贴板；只在 Windows 上调用时才 import。"""
    import win32clipboard

    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32clipboard.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()


def _upload_orders(session, push_button_id, orders):
    """打开订单多选，清空后用剪贴板批量上载。"""
    _press(session, push_button_id)
    _press(session, "wnd[1]/tbar[0]/btn[16]")
    copy_to_clipboard("\r\n".join(orders))
    _press(session, "wnd[1]/tbar[0]/btn[24]")
    _press(session, "wnd[1]/tbar[0]/btn[8]")


def _result(session):
    message_type, text = read_status(session)
    return is_success(message_type, text), text or f"消息类型: {message_type}"


def confirm_times(session, orders):
    """IW48：批量工时确认。"""
    if not orders:
        return True, "没有需要工时确认的工单"

    try:
        _resize(session)
        _set_okcode(session, "/nIW48")
        _press(session, "wnd[0]/tbar[1]/btn[19]")

        _clear_multi_select(session, "wnd[0]/usr/btn%_ARBPL_%_APP_%-VALU_PUSH")
        _clear_multi_select(session, "wnd[0]/usr/btn%_WERKS_%_APP_%-VALU_PUSH")
        _clear_multi_select(session, "wnd[0]/usr/btn%_AUART_%_APP_%-VALU_PUSH")
        _upload_orders(session, _ORDER_POPUP, orders)

        _press(session, "wnd[0]/tbar[1]/btn[8]")
        wait_for_control(session, _CONFIRM_CHECKBOXES[0])

        for control_id in _CONFIRM_CHECKBOXES:
            session.findById(control_id).selected = True

        _press(session, "wnd[0]/tbar[1]/btn[9]")
        _press(session, "wnd[0]/tbar[1]/btn[38]")
        _press(session, "wnd[0]/tbar[0]/btn[11]")

        return _result(session)
    except Exception as error:
        return False, f"{type(error).__name__}: {error}"


def complete_technically(session, orders):
    """IW38：批量技术完成。"""
    if not orders:
        return True, "没有需要技术完成的工单"

    try:
        _resize(session)
        _set_okcode(session, "/nIW38")
        _upload_orders(session, _ORDER_POPUP, orders)

        session.findById("wnd[0]/usr/ctxtVARIANT").text = "/KEL"

        _press(session, "wnd[0]/tbar[1]/btn[8]")
        time.sleep(3)
        _press(session, "wnd[0]/tbar[1]/btn[5]")
        _press(session, "wnd[0]/tbar[1]/btn[45]")
        _press(session, "wnd[1]/usr/btnSPOP-VAROPTION1")

        return _result(session)
    except Exception as error:
        return False, f"{type(error).__name__}: {error}"
