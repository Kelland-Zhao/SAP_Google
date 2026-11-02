import unittest
from unittest.mock import patch

import workorder_close_sap as sap


class FakeControl:
    def __init__(self, cid, session):
        self.id = cid
        self.session = session
        self.text = ""
        self.selected = False

    def press(self):
        self.session.pressed.append(self.id)

    def sendVKey(self, key):
        self.session.sent.append((self.id, key))

    def setFocus(self):
        pass

    def resizeWorkingPane(self, *args):
        self.session.resized.append(self.id)


class FakeStatusBar:
    def __init__(self, message_type, text):
        self.messageType = message_type
        self.text = text


class FakeSession:
    def __init__(self, message_type="S", text="已保存", missing=()):
        self.pressed = []
        self.sent = []
        self.resized = []
        self.controls = {}
        self.message_type = message_type
        self.text = text
        self.missing = set(missing)

    def findById(self, cid):
        if cid in self.missing:
            raise RuntimeError(f"not found: {cid}")
        if cid == "wnd[0]/sbar":
            return FakeStatusBar(self.message_type, self.text)
        return self.controls.setdefault(cid, FakeControl(cid, self))


class ConfirmTimesTests(unittest.TestCase):
    def test_uploads_orders_and_presses_recorded_sequence(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard") as clipboard:
            ok, message = sap.confirm_times(session, ["5609259", "5611111"])
        self.assertTrue(ok)
        clipboard.assert_called_once_with("5609259\r\n5611111")
        for cid in (
            "wnd[0]/tbar[1]/btn[19]",
            "wnd[1]/tbar[0]/btn[24]",
            "wnd[0]/tbar[1]/btn[8]",
            "wnd[0]/tbar[1]/btn[9]",
            "wnd[0]/tbar[1]/btn[38]",
            "wnd[0]/tbar[0]/btn[11]",
        ):
            self.assertIn(cid, session.pressed)
        for cid in (
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-AUERU[9,0]",
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-AUSOR[10,0]",
            "wnd[0]/usr/tblSAPLCORUTABCDEF_3350/chkAFRD-LEKNW[11,0]",
        ):
            self.assertTrue(session.controls[cid].selected)

    def test_no_orders_does_not_touch_sap(self):
        session = FakeSession()
        ok, message = sap.confirm_times(session, [])
        self.assertTrue(ok)
        self.assertEqual(session.pressed, [])

    def test_error_message_returns_false(self):
        session = FakeSession(message_type="E", text="订单已锁定")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("订单已锁定", message)

    def test_empty_message_returns_false(self):
        session = FakeSession(message_type="", text="")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, _ = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)

    def test_missing_confirmation_table_returns_false(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"), \
             patch("workorder_close_sap.wait_for_control",
                   side_effect=TimeoutError("等待控件超时")):
            ok, message = sap.confirm_times(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("TimeoutError", message)


class CompleteTechnicallyTests(unittest.TestCase):
    def test_uploads_orders_and_presses_recorded_sequence(self):
        session = FakeSession()
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard") as clipboard:
            ok, message = sap.complete_technically(session, ["5609259", "5611111"])
        self.assertTrue(ok)
        clipboard.assert_called_once_with("5609259\r\n5611111")
        for cid in (
            "wnd[1]/tbar[0]/btn[24]",
            "wnd[1]/tbar[0]/btn[8]",
            "wnd[0]/tbar[1]/btn[8]",
            "wnd[0]/tbar[1]/btn[5]",
            "wnd[0]/tbar[1]/btn[45]",
            "wnd[1]/usr/btnSPOP-VAROPTION1",
        ):
            self.assertIn(cid, session.pressed)
        self.assertEqual(session.controls["wnd[0]/usr/ctxtVARIANT"].text, "/KEL")

    def test_no_orders_does_not_touch_sap(self):
        session = FakeSession()
        ok, message = sap.complete_technically(session, [])
        self.assertTrue(ok)
        self.assertEqual(session.pressed, [])

    def test_error_message_returns_false(self):
        session = FakeSession(message_type="E", text="未选择订单")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.complete_technically(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("未选择订单", message)

    def test_warning_that_says_no_orders_selected_is_failure(self):
        session = FakeSession(message_type="W", text="未选择订单")
        with patch("workorder_close_sap.time.sleep"), \
             patch("workorder_close_sap.copy_to_clipboard"):
            ok, message = sap.complete_technically(session, ["5609259"])
        self.assertFalse(ok)
        self.assertIn("未选择订单", message)


class IsSuccessTests(unittest.TestCase):
    def test_success_and_warning_are_success(self):
        self.assertTrue(sap.is_success("S", "已保存"))
        self.assertTrue(sap.is_success("W", "警告"))

    def test_error_abort_empty_and_unknown_are_failure(self):
        for message_type in ("E", "A", "", "X"):
            self.assertFalse(sap.is_success(message_type, "消息"))

    def test_warning_with_negative_text_is_failure(self):
        self.assertFalse(sap.is_success("W", "未选择订单"))
        self.assertFalse(sap.is_success("S", "not selected"))
