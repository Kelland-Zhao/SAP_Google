import unittest
from unittest.mock import patch

import workorder_cancel_sap as sap
from workorder_cancel_utils import SessionLostError


class FakeControl:
    def __init__(self, cid, session):
        self.id = cid
        self.session = session
        self.text = ""
        self.selected = False

    def press(self):
        self.session.pressed.append(self.id)
        self.session.maybe_popup("save")

    def select(self):
        self.session.selected.append(self.id)
        self.session.maybe_popup("menu")

    def sendVKey(self, key):
        self.session.sent.append((self.id, key))
        if self.id == sap.WINDOW and key == 0:
            field = self.session.controls.get(sap.ORDER_FIELD)
            if field is not None and field.text:
                self.session.maybe_popup("enter")

    def maximize(self):
        self.session.maximized.append(self.id)


class FakeStatusBar:
    def __init__(self, message_type="", text=""):
        self.messageType = message_type
        self.text = text


class FakePopup:
    def __init__(self, message_type="", text=""):
        self.messageType = message_type
        self.text = text

    def sendVKey(self, key):
        self.session.popup = None


class FakeSession:
    def __init__(self, message_type="S", text="订单 5650370 已保存",
                 missing=(), popup_after=None, popup_type="", popup_text=""):
        self.pressed = []
        self.selected = []
        self.sent = []
        self.maximized = []
        self.message_type = message_type
        self.text = text
        self.missing = set(missing)
        self.popup_after = popup_after
        self.popup_type = popup_type
        self.popup_text = popup_text
        self.popup = None
        self.controls = {}

    def maybe_popup(self, moment):
        if self.popup_after == moment:
            self.popup = FakePopup(self.popup_type, self.popup_text)
            self.popup.session = self

    def findById(self, cid):
        if cid in self.missing:
            raise RuntimeError(f"not found: {cid}")
        if cid == "wnd[0]/sbar":
            return FakeStatusBar(self.message_type, self.text)
        if cid == sap.POPUP:
            if self.popup is None:
                raise RuntimeError("not found: wnd[1]")
            return self.popup
        if cid == sap.POPUP_STATUS:
            if self.popup is None:
                raise RuntimeError("not found: wnd[1]/sbar")
            return FakeStatusBar(self.popup.messageType, self.popup.text)
        return self.controls.setdefault(cid, FakeControl(cid, self))


class CancelTests(unittest.TestCase):
    def test_presses_recorded_sequence(self):
        session = FakeSession()
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertTrue(ok, message)
        self.assertEqual(session.controls[sap.OKCODE].text, "/nIW32")
        self.assertEqual(session.controls[sap.ORDER_FIELD].text, "5650370")
        self.assertEqual(session.selected, [sap.CANCEL_MENU])
        self.assertEqual(session.pressed, [sap.SAVE_BUTTON])
        self.assertEqual(session.sent, [(sap.WINDOW, 0), (sap.WINDOW, 0)])

    def test_error_message_fails(self):
        session = FakeSession(message_type="E", text="订单已锁定")
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertIn("订单已锁定", message)

    def test_success_without_saved_keyword_fails(self):
        session = FakeSession(message_type="S", text="处理完成")
        ok, _ = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)

    def test_extra_negative_keyword_fails(self):
        session = FakeSession(message_type="S", text="订单 5650370 未技术完成")
        ok, _ = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)

    def test_popup_error_skips_save(self):
        session = FakeSession(
            popup_after="enter", popup_type="E", popup_text="订单已锁定",
        )
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertEqual(session.pressed, [])
        self.assertIn("订单已锁定", message)

    def test_confirmation_popup_is_dismissed_then_save(self):
        session = FakeSession(popup_after="menu")
        ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertTrue(ok, message)
        self.assertEqual(session.pressed, [sap.SAVE_BUTTON])

    def test_missing_okcode_raises_session_lost(self):
        session = FakeSession(missing={sap.OKCODE})
        with self.assertRaises(SessionLostError):
            sap.cancel_technical_completion(session, "5650370")

    def test_order_field_timeout_raises_session_lost(self):
        session = FakeSession(missing={sap.ORDER_FIELD})
        with patch(
            "workorder_cancel_sap.wait_for_control",
            side_effect=TimeoutError("等待控件超时"),
        ):
            with self.assertRaises(SessionLostError):
                sap.cancel_technical_completion(session, "5650370")

    def test_menu_timeout_uses_status_bar_text(self):
        session = FakeSession(
            message_type="E", text="订单不存在", missing={sap.CANCEL_MENU},
        )

        def fake_wait(session, control_id, timeout=30.0, interval=0.5):
            if control_id == sap.ORDER_FIELD:
                return session.findById(control_id)
            raise TimeoutError("等待控件超时: " + control_id)

        with patch("workorder_cancel_sap.wait_for_control", side_effect=fake_wait):
            ok, message = sap.cancel_technical_completion(session, "5650370")
        self.assertFalse(ok)
        self.assertEqual(message, "订单不存在")


class SuccessIsProvenTests(unittest.TestCase):
    def test_requires_saved_keyword(self):
        self.assertTrue(sap.success_is_proven("S", "订单 5650370 已保存"))
        self.assertTrue(sap.success_is_proven("W", "已保存，有警告"))
        self.assertFalse(sap.success_is_proven("S", "处理完成"))
        self.assertFalse(sap.success_is_proven("S", "订单 未技术完成"))
        self.assertFalse(sap.success_is_proven("E", "已保存"))
        self.assertFalse(sap.success_is_proven("S", "not changed"))
