from contextlib import redirect_stdout
import importlib.util
import io
from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import MagicMock, patch

import openpyxl
import win32con


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPT_PATH = (
    PROJECT_ROOT
    / "行情交易"
    / "下单"
    / "自动化下单"
    / "4.快速下单_自动化下单_Excel驱动版.py"
)
MODULE_NAME = "quick_order_script"
SPEC = importlib.util.spec_from_file_location(MODULE_NAME, SCRIPT_PATH)
assert SPEC is not None and SPEC.loader is not None
quick_order = importlib.util.module_from_spec(SPEC)
sys.modules[MODULE_NAME] = quick_order
SPEC.loader.exec_module(quick_order)

ACTION_TITLES = quick_order.ACTION_TITLES
MAX_DIALOGS = quick_order.MAX_DIALOGS
QUOTE_TYPES = quick_order.QUOTE_TYPES
REQUIRED_COLUMNS = quick_order.REQUIRED_COLUMNS
SUBMIT_DIALOG_DELAY = quick_order.SUBMIT_DIALOG_DELAY
SUBMIT_DIALOG_TIMEOUT = quick_order.SUBMIT_DIALOG_TIMEOUT
QuickOrder = quick_order.QuickOrder
_confirm_new_dialogs = quick_order._confirm_new_dialogs
_set_confirmed_toggle = quick_order._set_confirmed_toggle
_set_edit_verified = quick_order._set_edit_verified
execute_order = quick_order.execute_order
load_orders = quick_order.load_orders


class QuickOrderExcelTests(unittest.TestCase):
    def _write_workbook(self, headers, rows):
        temp_dir = tempfile.TemporaryDirectory()
        self.addCleanup(temp_dir.cleanup)
        path = Path(temp_dir.name) / "quick_order.xlsx"
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(list(headers))
        for row in rows:
            sheet.append(list(row))
        workbook.save(path)
        workbook.close()
        return str(path)

    def test_loads_linked_order_and_ignores_extra_unique_column(self):
        headers = (*REQUIRED_COLUMNS, "备注")
        path = self._write_workbook(
            headers,
            [
                (
                    "快速下单",
                    10011851,
                    "对手价",
                    2,
                    "买入",
                    "开",
                    None,
                    None,
                    True,
                    "测试数据",
                )
            ],
        )

        orders = load_orders(path)

        self.assertEqual(len(orders), 1)
        order = orders[0]
        self.assertEqual(order.contract_code, "10011851")
        self.assertEqual(order.action_title, "买入开仓")
        self.assertEqual(order.quantity, "2")
        self.assertTrue(order.linked)

    def test_loads_unlinked_order_and_all_quantity(self):
        path = self._write_workbook(
            REQUIRED_COLUMNS,
            [
                (
                    "快速下单",
                    "10011852",
                    "限价",
                    "全",
                    "卖出",
                    "平",
                    False,
                    True,
                    False,
                )
            ],
        )

        order = load_orders(path)[0]

        self.assertEqual(order.action_title, "卖出平仓")
        self.assertEqual(order.quantity, "全")
        self.assertTrue(order.fok)

    def test_blank_linked_defaults_to_false(self):
        path = self._write_workbook(
            REQUIRED_COLUMNS,
            [
                (
                    "快速下单",
                    "10011852",
                    "限价",
                    1,
                    "卖出",
                    "平",
                    False,
                    False,
                    None,
                )
            ],
        )

        order = load_orders(path)[0]

        self.assertFalse(order.linked)

    def test_rejects_missing_required_column_before_gui_work(self):
        headers = [header for header in REQUIRED_COLUMNS if header != "开平"]
        path = self._write_workbook(headers, [])

        with self.assertRaisesRegex(ValueError, "缺少必填列.*开平"):
            load_orders(path)

    def test_rejects_duplicate_headers(self):
        path = self._write_workbook((*REQUIRED_COLUMNS, "委托数量"), [])

        with self.assertRaisesRegex(ValueError, "表头重复.*委托数量"):
            load_orders(path)

    def test_reports_excel_row_for_invalid_value(self):
        path = self._write_workbook(
            REQUIRED_COLUMNS,
            [
                (
                    "快速下单",
                    "10011853",
                    "对手价",
                    1,
                    "购买",
                    "开",
                    False,
                    False,
                    True,
                )
            ],
        )

        with self.assertRaisesRegex(ValueError, "第 2 行.*动作"):
            load_orders(path)

    def test_live_quote_order_and_action_combinations_are_fixed(self):
        self.assertEqual(
            QUOTE_TYPES,
            (
                "对手价",
                "挂盘价",
                "涨停价",
                "跌停价",
                "限价",
                "超价",
                "市价转限",
                "市价FAK",
                "市价FOK",
            ),
        )
        self.assertEqual(
            set(ACTION_TITLES.values()),
            {"买入开仓", "卖出开仓", "买入平仓", "卖出平仓"},
        )

    def test_owner_drawn_edit_uses_message_result_and_text_length(self):
        def send_message(_hwnd, message, _wparam, _lparam):
            if message == win32con.WM_SETTEXT:
                return 1
            if message == win32con.WM_GETTEXTLENGTH:
                return 8
            raise AssertionError(f"未预期的消息: {message}")

        with (
            patch("quick_order_script._find_control", return_value=123),
            patch("quick_order_script.win32gui.SendMessage", side_effect=send_message),
            patch("quick_order_script.win32gui.GetWindowText") as get_window_text,
            patch("quick_order_script._mouse_click") as mouse_click,
            patch("quick_order_script.send_keys") as keyboard_input,
        ):
            _set_edit_verified(100, 18005, "10011857")

        get_window_text.assert_not_called()
        mouse_click.assert_not_called()
        keyboard_input.assert_not_called()

    def test_linked_toggle_confirms_new_dialog_before_state_check(self):
        with (
            patch("quick_order_script._find_control", return_value=5024),
            patch(
                "quick_order_script.win32gui.SendMessage",
                side_effect=[0, 1],
            ),
            patch("quick_order_script.win32gui.IsWindowEnabled", return_value=True),
            patch("quick_order_script._process_id", return_value=99),
            patch(
                "quick_order_script._visible_process_windows",
                return_value={700},
            ),
            patch("quick_order_script._mouse_click") as mouse_click,
            patch(
                "quick_order_script._confirm_new_dialogs",
                return_value=1,
            ) as confirm,
        ):
            _set_confirmed_toggle(100, 5024, True, "联动")

        mouse_click.assert_called_once_with(5024)
        confirm.assert_called_once_with(
            100,
            {700},
            first_timeout=2.0,
            context="切换‘联动’",
        )

    def test_linked_toggle_fails_if_confirmation_dialog_does_not_appear(self):
        with (
            patch("quick_order_script._find_control", return_value=5024),
            patch("quick_order_script.win32gui.SendMessage", return_value=0),
            patch("quick_order_script.win32gui.IsWindowEnabled", return_value=True),
            patch("quick_order_script._process_id", return_value=99),
            patch("quick_order_script._visible_process_windows", return_value=set()),
            patch("quick_order_script._mouse_click"),
            patch("quick_order_script._confirm_new_dialogs", return_value=0),
        ):
            with self.assertRaisesRegex(RuntimeError, "未出现确认提示框"):
                _set_confirmed_toggle(100, 5024, True, "联动")

    def test_execute_order_waits_then_confirms_delayed_dialog_chain(self):
        order = QuickOrder(
            excel_row=7,
            contract_code="10011857",
            direction="买入",
            offset="开仓",
            quote_type="对手价",
            quantity="1",
            covered=False,
            fok=False,
            linked=False,
        )
        output = io.StringIO()
        with (
            patch("quick_order_script._set_edit_verified"),
            patch("quick_order_script._set_radio"),
            patch("quick_order_script._select_quote_type"),
            patch("quick_order_script._set_toggle"),
            patch("quick_order_script._set_confirmed_toggle"),
            patch("quick_order_script._set_quantity"),
            patch("quick_order_script._find_control", return_value=18001),
            patch("quick_order_script.win32gui.IsWindowEnabled", return_value=True),
            patch("quick_order_script.win32gui.GetWindowText", return_value="买入开仓"),
            patch("quick_order_script._process_id", return_value=99),
            patch("quick_order_script._visible_process_windows", return_value={700}),
            patch("quick_order_script._mouse_click"),
            patch("quick_order_script.time.sleep") as sleep,
            patch("quick_order_script._confirm_new_dialogs", return_value=2) as confirm,
            redirect_stdout(output),
        ):
            execute_order(100, order)

        confirm.assert_called_once_with(
            100,
            {700},
            max_dialogs=MAX_DIALOGS,
            first_timeout=SUBMIT_DIALOG_TIMEOUT,
            next_timeout=SUBMIT_DIALOG_TIMEOUT,
            context="本次下单",
        )
        sleep.assert_any_call(SUBMIT_DIALOG_DELAY)
        log = output.getvalue()
        self.assertIn("合约代码已填写并校验: 10011857", log)
        self.assertIn("下单动作已触发: 买入开仓", log)
        self.assertIn("确认弹窗=2 个", log)

    def test_confirm_new_dialogs_keeps_waiting_for_followup_dialog(self):
        visible_results = iter(({10}, {20}))

        def visible_windows(_pid, _main_hwnd):
            return next(visible_results, set())

        clock = [0.0]

        def monotonic():
            clock[0] += 1.0
            return clock[0]

        wrapper = MagicMock()
        application = MagicMock()
        application.return_value.connect.return_value.window.return_value = wrapper
        output = io.StringIO()
        with (
            patch("quick_order_script._process_id", return_value=99),
            patch(
                "quick_order_script._visible_process_windows",
                side_effect=visible_windows,
            ),
            patch("quick_order_script.time.monotonic", side_effect=monotonic),
            patch("quick_order_script.time.sleep"),
            patch("quick_order_script.win32gui.GetWindowText", return_value="快速下单"),
            patch("quick_order_script._wait_until", return_value=True),
            patch("quick_order_script.Application", application),
            redirect_stdout(output),
        ):
            confirmed = _confirm_new_dialogs(
                100,
                set(),
                first_timeout=3.0,
                next_timeout=3.0,
                context="本次下单",
            )

        self.assertEqual(confirmed, 2)
        self.assertEqual(wrapper.type_keys.call_count, 2)
        self.assertIn("已确认第 2 个弹窗", output.getvalue())


if __name__ == "__main__":
    unittest.main()
