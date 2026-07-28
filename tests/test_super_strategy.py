# -*- coding: utf-8 -*-

import unittest
from contextlib import nullcontext
from unittest import mock

from core import super_strategy
from core import tactics_panel


class PromptClassificationTests(unittest.TestCase):
    def test_login_prompt_is_recognized_after_whitespace_normalization(self):
        self.assertTrue(super_strategy._is_login_prompt("请先 登录 交易账号"))

    def test_trading_time_prompt_matches_real_client_message(self):
        text = (
            "errcode=11205, errmsg=交易时间错误"
            "[系统停市期间禁止委托(STKBD:上海股票期权(15))]"
        )
        self.assertTrue(super_strategy._is_trading_time_prompt(text))

    def test_order_confirmation_requires_explicit_question(self):
        self.assertTrue(super_strategy._is_order_confirmation("您确定要下单吗？"))
        self.assertFalse(super_strategy._is_order_confirmation("下单失败，请重试"))

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1, 2})
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="提示\n系统停市期间禁止委托",
    )
    def test_wait_after_open_raises_trading_time_error(self, *_):
        with self.assertRaises(super_strategy.TradingTimeBlockedError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1, 2})
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="请先登录后再进行委托",
    )
    def test_wait_after_open_raises_login_error(self, *_):
        with self.assertRaises(super_strategy.LoginRequiredError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._confirm_order_dialog")
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1, 3}],
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        side_effect=["下单\n确定\n取消", "提示\n系统停市期间禁止委托"],
    )
    @mock.patch(
        "core.super_strategy._ocr_window_text",
        return_value="您确定要下单吗？",
    )
    def test_wait_after_open_confirms_order_then_classifies_result(
        self, _, __, ___, ____, _____, confirm, ______
    ):
        with self.assertRaises(super_strategy.TradingTimeBlockedError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1)
        confirm.assert_called_once()
        self.assertEqual(confirm.call_args.args[0], 2)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1})
    @mock.patch("core.super_strategy._ocr_window_text", return_value="")
    def test_wait_after_open_rejects_missing_feedback(self, *_):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "未检测到批量下单结果"
        ):
            super_strategy.wait_after_open(99, {1}, timeout=0)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1, 2})
    @mock.patch("core.super_strategy._window_text_dump", return_value="陌生错误")
    @mock.patch("core.super_strategy._ocr_window_text", return_value="无法分类的提示")
    def test_wait_after_open_closes_unknown_dialog_before_error(
        self, _, __, ___, ____, dismiss, _____
    ):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "未识别弹窗.*弹窗已关闭"
        ):
            super_strategy.wait_after_open(99, {1}, timeout=0.1)
        dismiss.assert_called_once_with(2)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1, 2})
    @mock.patch(
        "core.super_strategy._confirm_success_result_dialog",
        side_effect=lambda hwnd, text: super_strategy._handle_batch_result(text),
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="期权合约批量下单处理完毕：成功2笔，失败0笔。",
    )
    def test_wait_after_open_confirms_successful_batch_result(self, *_):
        result = super_strategy.wait_after_open(99, {1}, timeout=0.1)
        self.assertTrue(result["open_confirmed"])
        self.assertEqual(result["success_count"], 2)
        self.assertEqual(result["failure_count"], 0)

    def test_failed_batch_result_is_not_success(self):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "成功0笔，失败2笔"
        ):
            super_strategy._handle_batch_result(
                "期权合约批量下单处理完毕：成功0笔，失败2笔。"
            )

    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    def test_failed_batch_result_closes_known_dialog_before_error(self, dismiss):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "成功0笔，失败2笔"
        ):
            super_strategy._handle_batch_result(
                "期权合约批量下单处理完毕：成功0笔，失败2笔。", 22
            )
        dismiss.assert_called_once_with(22)


class StrategyStartupTests(unittest.TestCase):
    @mock.patch.dict(
        "os.environ",
        {"GUI_CLIENT_ID": "guotai_haitong", "GUI_COUNTDOWN": "3"},
        clear=False,
    )
    @mock.patch("core.super_strategy.click_tactics_item", return_value={"text": "牛市认沽"})
    @mock.patch("core.super_strategy.ensure_workspace")
    @mock.patch("core.super_strategy._activate_main_window")
    @mock.patch("core.super_strategy.find_window", return_value=11)
    @mock.patch(
        "core.super_strategy.get_client",
        return_value={"window_key": "国泰海通证券期权宝"},
    )
    @mock.patch("core.super_strategy.countdown")
    def test_strategy_counts_down_and_activates_before_navigation(
        self, countdown, _, find_window, activate, ensure, click_menu
    ):
        manager = mock.Mock()
        manager.attach_mock(countdown, "countdown")
        manager.attach_mock(find_window, "find_window")
        manager.attach_mock(activate, "activate")
        manager.attach_mock(ensure, "ensure")
        manager.attach_mock(click_menu, "click_menu")

        result = super_strategy.run_strategy(
            "牛市认沽", add_underlying=False, execute_open=False
        )

        self.assertFalse(result["open_triggered"])
        self.assertEqual(
            [call[0] for call in manager.mock_calls],
            ["countdown", "find_window", "activate", "ensure", "click_menu"],
        )
        countdown.assert_called_once_with(3)
        activate.assert_called_once_with(11)


class RealMouseClickTests(unittest.TestCase):
    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy.win32gui.IsWindowVisible", return_value=False)
    @mock.patch("core.super_strategy.win32gui.IsWindow", return_value=False)
    @mock.patch("core.super_strategy._real_mouse_click")
    @mock.patch("core.super_strategy.dpi_unaware", return_value=nullcontext())
    @mock.patch(
        "core.super_strategy.win32gui.GetWindowRect",
        return_value=(700, 500, 800, 540),
    )
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[22])
    def test_known_or_unknown_prompt_prefers_unique_confirm_button(
        self, _, __, ___, real_click, ____, _____, ______
    ):
        super_strategy._dismiss_prompt_dialog(11)
        real_click.assert_called_once_with(11, 22, 750.0, 520.0, delay=0.3)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy.win32gui.IsWindowVisible", return_value=False)
    @mock.patch("core.super_strategy.win32gui.IsWindow", return_value=False)
    @mock.patch("core.super_strategy.win32gui.PostMessage")
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[])
    def test_prompt_without_unique_confirm_uses_close_message(
        self, _, post_message, __, ___, ____
    ):
        super_strategy._dismiss_prompt_dialog(11)
        post_message.assert_called_once_with(11, super_strategy.win32con.WM_CLOSE, 0, 0)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy.win32gui.IsWindowVisible", return_value=False)
    @mock.patch("core.super_strategy.win32gui.IsWindow", return_value=False)
    @mock.patch("core.super_strategy._real_mouse_click")
    @mock.patch("core.super_strategy.dpi_unaware", return_value=nullcontext())
    @mock.patch(
        "core.super_strategy.win32gui.GetWindowRect",
        return_value=(700, 500, 800, 540),
    )
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[22])
    @mock.patch("core.super_strategy.win32gui.GetWindowText", return_value="提示")
    def test_success_result_clicks_unique_confirm_and_waits_for_close(
        self, _, __, ___, ____, real_click, _____, ______, _______
    ):
        result = super_strategy._confirm_success_result_dialog(
            11, "期权合约批量下单处理完毕：成功2笔，失败0笔。"
        )
        real_click.assert_called_once_with(11, 22, 750.0, 520.0, delay=0.3)
        self.assertEqual(result["success_count"], 2)

    @mock.patch("core.super_strategy._real_mouse_click")
    @mock.patch(
        "core.super_strategy._find_action_text",
        return_value={"cx": 1050.4, "cy": 640.6, "score": 0.99, "ocr_elapsed": 0.2},
    )
    @mock.patch("core.super_strategy.get_action_panel", return_value=22)
    def test_action_uses_real_mouse_at_ocr_screen_point(self, _, __, real_click):
        super_strategy.click_action(11, super_strategy.OPEN_POSITION_TEXT, delay=0.3)
        real_click.assert_called_once_with(
            11, 22, 1050.4, 640.6, delay=0.3
        )

    @mock.patch("core.super_strategy.dpi_unaware", return_value=nullcontext())
    @mock.patch("core.super_strategy._activate_main_window")
    @mock.patch("core.super_strategy.win32api.mouse_event")
    @mock.patch("core.super_strategy.win32api.SetCursorPos")
    @mock.patch("core.super_strategy.win32gui.IsChild", return_value=False)
    @mock.patch("core.super_strategy.win32gui.WindowFromPoint", return_value=22)
    @mock.patch(
        "core.super_strategy.win32gui.GetWindowRect",
        return_value=(100, 200, 1200, 700),
    )
    def test_real_mouse_click_checks_target_then_sends_hardware_events(
        self, _, window_from_point, __, set_cursor, mouse_event, activate, ___
    ):
        super_strategy._real_mouse_click(11, 22, 1050.4, 640.6, delay=0)
        activate.assert_called_once_with(11)
        window_from_point.assert_called_once_with((1050, 641))
        set_cursor.assert_called_once_with((1050, 641))
        self.assertEqual(mouse_event.call_count, 2)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._set_foreground_with_attached_input")
    @mock.patch("core.super_strategy.win32gui.GetAncestor", side_effect=[44, 11])
    @mock.patch("core.super_strategy.win32gui.GetForegroundWindow", side_effect=[44, 11])
    @mock.patch(
        "core.super_strategy.win32gui.SetForegroundWindow",
        side_effect=RuntimeError("foreground locked"),
    )
    @mock.patch("core.super_strategy.win32gui.SetWindowPos")
    @mock.patch("core.super_strategy.win32gui.IsIconic", return_value=False)
    def test_activate_falls_back_to_attached_input_without_extra_click(
        self, _, set_pos, __, ___, ____, attached_input, _____
    ):
        super_strategy._activate_main_window(11)
        attached_input.assert_called_once_with(11)
        self.assertEqual(set_pos.call_count, 2)


class TacticsOcrTests(unittest.TestCase):
    def test_four_formal_targets_have_distinct_fast_ocr_cells(self):
        self.assertEqual(set(tactics_panel.FORMAL_TARGET_CELLS), tactics_panel.SUPER_TARGETS)
        self.assertEqual(
            len(set(tactics_panel.FORMAL_TARGET_CELLS.values())),
            len(tactics_panel.SUPER_TARGETS),
        )

    def test_exact_target_match_wins(self):
        hit = tactics_panel._match_target(
            [
                {"text": "牛市认购", "score": 0.99},
                {"text": "牛市认沽", "score": 0.99},
            ],
            "牛市认购",
            fuzzy_threshold=0.8,
        )
        self.assertEqual(hit["text"], "牛市认购")


if __name__ == "__main__":
    unittest.main()
