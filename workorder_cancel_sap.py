"""10 - 工单取消技术完成的 SAP GUI 操作。

只依赖传入的 session 对象，不 import win32com / sap_common，
所以可以在 macOS 上用假 session 跑单元测试。

操作顺序来自 SAP 录制脚本：
  取消技术完成IW32.txt
"""

from workorder_close_sap import (
    NEGATIVE_KEYWORDS, is_success, read_status, wait_for_control,
)
from workorder_cancel_utils import SessionLostError

OKCODE = "wnd[0]/tbar[0]/okcd"
WINDOW = "wnd[0]"
ORDER_FIELD = "wnd[0]/usr/ctxtCAUFVD-AUFNR"
CANCEL_MENU = "wnd[0]/mbar/menu[0]/menu[8]/menu[2]/menu[1]"
SAVE_BUTTON = "wnd[0]/tbar[0]/btn[11]"
POPUP = "wnd[1]"
POPUP_STATUS = "wnd[1]/sbar"

EXTRA_NEGATIVE_KEYWORDS = (
    "未技术完成", "没有技术完成", "无法取消", "未执行",
    "not technically completed", "not possible", "no change",
)
SAVED_KEYWORDS = ("保存", "saved")

_ALL_NEGATIVE = tuple(NEGATIVE_KEYWORDS) + EXTRA_NEGATIVE_KEYWORDS


def _has_negative(text):
    lowered = str(text or "").lower()
    return any(keyword.lower() in lowered for keyword in _ALL_NEGATIVE)


def success_is_proven(message_type, text):
    """S/W、非空、无否定词、且含「保存」/saved 才算成功。"""
    if not str(text or "").strip():
        return False
    if _has_negative(text):
        return False
    if not is_success(message_type, text):
        return False
    lowered = str(text or "").lower()
    return any(keyword.lower() in lowered for keyword in SAVED_KEYWORDS)


def prepare_session(session):
    """把主窗口最大化；纯外观操作，失败不影响后续。"""
    try:
        session.findById(WINDOW).maximize()
    except Exception:
        pass


def _set_okcode(session, code):
    try:
        session.findById(OKCODE).text = code
        session.findById(WINDOW).sendVKey(0)
    except Exception as error:
        raise SessionLostError(
            f"SAP 会话不可用: {type(error).__name__}: {error}"
        ) from error


def _popup_state(session):
    """返回 (是否存在, 消息类型, 文本)；没有状态栏时类型/文本为空。"""
    try:
        session.findById(POPUP)
    except Exception:
        return False, "", ""

    try:
        sbar = session.findById(POPUP_STATUS)
    except Exception:
        return True, "", ""

    message_type = getattr(sbar, "messageType", None)
    if message_type is None:
        message_type = getattr(sbar, "MessageType", "")
    text = getattr(sbar, "text", None)
    if text is None:
        text = getattr(sbar, "Text", "")
    return True, str(message_type or "").upper(), str(text or "")


def _popup_is_error(message_type, text):
    if str(message_type).upper() in {"E", "A"}:
        return True
    return _has_negative(text)


def _dismiss_popup(session):
    try:
        session.findById(POPUP).sendVKey(0)
    except Exception:
        pass


def _dismiss_stray_popup(session):
    """进单前清掉上一单可能留下的弹窗，最多两次。"""
    for _ in range(2):
        exists, _, _ = _popup_state(session)
        if not exists:
            return
        _dismiss_popup(session)


def cancel_technical_completion(session, order):
    """IW32：取消单笔工单的技术完成。"""
    try:
        _dismiss_stray_popup(session)

        _set_okcode(session, "/nIW32")
        try:
            wait_for_control(session, ORDER_FIELD, timeout=30.0)
        except TimeoutError as error:
            raise SessionLostError(
                f"SAP 会话疑似不可用: {error}"
            ) from error

        session.findById(ORDER_FIELD).text = order
        session.findById(WINDOW).sendVKey(0)

        exists, message_type, text = _popup_state(session)
        if exists:
            if _popup_is_error(message_type, text):
                _dismiss_popup(session)
                return False, text or "SAP 弹窗报错"
            _dismiss_popup(session)

        try:
            menu = wait_for_control(session, CANCEL_MENU, timeout=30.0)
        except TimeoutError as error:
            _, status_text = read_status(session)
            if status_text.strip():
                return False, status_text
            return False, f"等待订单屏幕超时: {error}"

        menu.select()

        popup_type, popup_text = "", ""
        exists, message_type, text = _popup_state(session)
        if exists:
            if _popup_is_error(message_type, text):
                _dismiss_popup(session)
                return False, text or "SAP 弹窗报错"
            popup_type, popup_text = message_type, text
            _dismiss_popup(session)

        session.findById(SAVE_BUTTON).press()

        exists, message_type, text = _popup_state(session)
        if exists:
            if _popup_is_error(message_type, text):
                _dismiss_popup(session)
                return False, text or "SAP 弹窗报错"
            popup_type, popup_text = message_type, text
            _dismiss_popup(session)

        message_type, text = read_status(session)
        if not text.strip():
            message_type, text = popup_type, popup_text

        if success_is_proven(message_type, text):
            return True, text
        return False, text or f"消息类型: {message_type}"
    except SessionLostError:
        raise
    except Exception as error:
        return False, f"{type(error).__name__}: {error}"
