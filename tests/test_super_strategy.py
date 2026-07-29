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
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1}],
    )
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="提示\n系统停市期间禁止委托",
    )
    def test_wait_after_open_raises_trading_time_error(self, *_):
        with self.assertRaises(super_strategy.TradingTimeBlockedError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1, quiet_period=0)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1}],
    )
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="请先登录后再进行委托",
    )
    def test_wait_after_open_raises_login_error(self, *_):
        with self.assertRaises(super_strategy.LoginRequiredError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1, quiet_period=0)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._confirm_order_dialog")
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1, 3}, {1}],
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        side_effect=[
            "下单\n您确定要下单吗？\n确定\n取消",
            "提示\n系统停市期间禁止委托",
        ],
    )
    def test_wait_after_open_confirms_order_then_classifies_result(
        self, _, __, ___, ____, confirm, _____
    ):
        with self.assertRaises(super_strategy.TradingTimeBlockedError):
            super_strategy.wait_after_open(99, {1}, timeout=0.1, quiet_period=0)
        confirm.assert_called_once()
        self.assertEqual(confirm.call_args.args[0], 2)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1})
    def test_wait_after_open_rejects_missing_feedback(self, *_):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "未检测到原生 Win32 弹窗"
        ):
            super_strategy.wait_after_open(99, {1}, timeout=0)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch("core.super_strategy._visible_process_surfaces", return_value={1, 2})
    @mock.patch("core.super_strategy._window_text_dump", return_value="陌生错误")
    def test_wait_after_open_closes_unknown_dialog_before_error(
        self, _, __, ___, dismiss, ____
    ):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError, "未识别弹窗.*弹窗已关闭"
        ):
            super_strategy.wait_after_open(99, {1}, timeout=0.1)
        dismiss.assert_called_once_with(2)

    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1}],
    )
    @mock.patch(
        "core.super_strategy._confirm_success_result_dialog",
        side_effect=lambda hwnd, text: super_strategy._handle_batch_result(text),
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="期权合约批量下单处理完毕：成功2笔，失败0笔。",
    )
    def test_wait_after_open_confirms_successful_batch_result(self, *_):
        result = super_strategy.wait_after_open(
            99, {1}, timeout=0.1, quiet_period=0
        )
        self.assertTrue(result["open_confirmed"])
        self.assertEqual(result["success_count"], 2)
        self.assertEqual(result["failure_count"], 0)

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1, 3}, {1}],
    )
    @mock.patch(
        "core.super_strategy._confirm_success_result_dialog",
        side_effect=[
            {
                "open_confirmed": True,
                "success_count": 2,
                "failure_count": 0,
            },
            {
                "open_confirmed": True,
                "success_count": 1,
                "failure_count": 0,
            },
        ],
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="期权合约批量下单处理完毕：成功2笔，失败0笔。",
    )
    def test_wait_after_open_records_open_then_delayed_add_result(
        self, _, confirm, __, ___, ____
    ):
        result = super_strategy.wait_after_open(
            99,
            {1},
            timeout=0.1,
            quiet_period=0,
            expect_add_result=True,
        )
        self.assertTrue(result["open_confirmed"])
        self.assertEqual(result["success_count"], 2)
        self.assertEqual(result["add_result"]["success_count"], 1)
        self.assertEqual(result["result_dialog_count"], 2)
        self.assertEqual(confirm.call_count, 2)
        self.assertEqual(
            [call.args[0] for call in confirm.call_args_list], [2, 3]
        )

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._visible_process_surfaces",
        side_effect=[{1, 2}, {1, 3}, {1}],
    )
    @mock.patch("core.super_strategy._dismiss_prompt_dialog")
    @mock.patch(
        "core.super_strategy._window_text_dump",
        side_effect=[
            "提示\n期权合约批量下单处理完毕：成功0笔，失败1笔。",
            (
                "提示\n期权合约批量下单处理完毕：成功0笔，失败2笔。\n"
                "失败信息如下：\n当前市场状态不允许交易\n当前市场状态不允许交易"
            ),
        ],
    )
    def test_wait_after_open_closes_both_delayed_failure_results(
        self, _, dismiss, __, ___, ____
    ):
        with self.assertRaisesRegex(
            super_strategy.SuperStrategyError,
            "一键开仓结果：成功0笔，失败1笔.*加入标的结果：成功0笔，失败2笔"
        ) as raised:
            super_strategy.wait_after_open(
                99,
                {1},
                timeout=0.1,
                quiet_period=0,
                expect_add_result=True,
            )
        self.assertIn("所有已出现的客户端弹窗均已关闭", str(raised.exception))
        self.assertEqual(
            [call.args[0] for call in dismiss.call_args_list], [2, 3]
        )

    @mock.patch("core.super_strategy.time.sleep")
    @mock.patch("core.super_strategy._process_id", return_value=7)
    @mock.patch(
        "core.super_strategy._confirm_success_result_dialog",
        return_value={
            "open_confirmed": True,
            "success_count": 2,
            "failure_count": 0,
        },
    )
    @mock.patch(
        "core.super_strategy._window_text_dump",
        return_value="期权合约批量下单处理完毕：成功2笔，失败0笔。",
    )
    def test_wait_after_open_reports_missing_add_result(
        self, _, __, ___, ____
    ):
        calls = 0

        def surfaces(*_):
            nonlocal calls
            calls += 1
            return {1, 2} if calls == 1 else {1}

        with mock.patch(
            "core.super_strategy._visible_process_surfaces",
            side_effect=surfaces,
        ):
            with self.assertRaisesRegex(
                super_strategy.SuperStrategyError,
                "期望至少2个，实际1个",
            ):
                super_strategy.wait_after_open(
                    99,
                    {1},
                    timeout=0.01,
                    quiet_period=0,
                    expect_add_result=True,
                )

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

    @mock.patch.dict(
        "os.environ",
        {"GUI_CLIENT_ID": "qianlong", "GUI_COUNTDOWN": "0"},
        clear=False,
    )
    def test_add_result_is_expected_after_open_not_before_it(self):
        open_result = {
            "open_confirmed": True,
            "success_count": 2,
            "failure_count": 0,
            "add_result": {
                "open_confirmed": True,
                "success_count": 1,
                "failure_count": 0,
            },
        }
        with (
            mock.patch("core.super_strategy.countdown"),
            mock.patch(
                "core.super_strategy.get_client",
                return_value={"window_key": "钱龙模拟期权宝"},
            ),
            mock.patch("core.super_strategy.find_window", return_value=11),
            mock.patch("core.super_strategy._activate_main_window"),
            mock.patch("core.super_strategy.ensure_workspace"),
            mock.patch(
                "core.super_strategy.click_tactics_item",
                return_value={"text": "牛市认沽"},
            ),
            mock.patch("core.super_strategy._process_id", return_value=7),
            mock.patch(
                "core.super_strategy._visible_process_surfaces",
                return_value={1},
            ),
            mock.patch("core.super_strategy.get_action_panel", return_value=22),
            mock.patch(
                "core.super_strategy._find_action_texts",
                return_value={
                    super_strategy.ADD_UNDERLYING_TEXT: {
                        "text": super_strategy.ADD_UNDERLYING_TEXT,
                    },
                    super_strategy.OPEN_POSITION_TEXT: {
                        "text": super_strategy.OPEN_POSITION_TEXT,
                    },
                },
            ),
            mock.patch("core.super_strategy.click_action") as click_action,
            mock.patch(
                "core.super_strategy.wait_after_open",
                return_value=open_result,
            ) as wait_after_open,
        ):
            result = super_strategy.run_strategy(
                "牛市认沽", add_underlying=True, execute_open=True
            )

        self.assertEqual(
            [call.args[1] for call in click_action.call_args_list],
            [super_strategy.ADD_UNDERLYING_TEXT, super_strategy.OPEN_POSITION_TEXT],
        )
        wait_after_open.assert_called_once_with(
            11, {1}, expect_add_result=True
        )
        self.assertEqual(result["success_count"], 2)
        self.assertEqual(result["add_result"]["success_count"], 1)


class NativePopupTests(unittest.TestCase):
    @mock.patch("core.super_strategy._wait_dialog_transition", return_value=True)
    @mock.patch("core.super_strategy._dialog_signature", return_value=("提示", ()))
    @mock.patch("core.super_strategy._dialog_is_visible", return_value=True)
    @mock.patch("core.super_strategy.win32gui.PostMessage")
    def test_dialog_button_uses_bm_click_and_verifies_transition(
        self, post_message, _, __, wait_transition
    ):
        super_strategy._click_dialog_button(
            11, 22, attempts=1, timeout=0.4, require_closed=True
        )
        post_message.assert_called_once_with(
            22, super_strategy.win32con.BM_CLICK, 0, 0
        )
        wait_transition.assert_called_once_with(
            11, ("提示", ()), timeout=0.4, require_closed=True
        )

    @mock.patch("core.super_strategy.win32gui.SendMessage")
    def test_native_window_text_uses_wm_gettext(self, send_message):
        def send(_, message, size, buffer):
            if message == super_strategy.win32con.WM_GETTEXTLENGTH:
                return 2
            self.assertEqual(message, super_strategy.win32con.WM_GETTEXT)
            self.assertEqual(size, 3)
            buffer.value = "提示"
            return 2

        send_message.side_effect = send
        self.assertEqual(super_strategy._native_window_text(11), "提示")

    @mock.patch("core.super_strategy._dialog_is_visible", return_value=False)
    @mock.patch("core.super_strategy._click_dialog_button")
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[22])
    def test_known_or_unknown_prompt_prefers_unique_confirm_button(self, _, click, __):
        super_strategy._dismiss_prompt_dialog(11)
        click.assert_called_once_with(
            11, 22, attempts=2, timeout=0.8, require_closed=True
        )

    @mock.patch("core.super_strategy._dialog_is_visible", return_value=False)
    @mock.patch("core.super_strategy.win32gui.PostMessage")
    @mock.patch("core.super_strategy._button_by_id", return_value=None)
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[])
    def test_prompt_without_unique_confirm_uses_close_message(
        self, _, __, post_message, ___
    ):
        super_strategy._dismiss_prompt_dialog(11)
        post_message.assert_called_once_with(11, super_strategy.win32con.WM_CLOSE, 0, 0)

    @mock.patch("core.super_strategy._click_dialog_button")
    @mock.patch("core.super_strategy._exact_child_buttons", return_value=[22])
    @mock.patch("core.super_strategy._native_window_text", return_value="提示")
    def test_success_result_clicks_unique_confirm_and_waits_for_close(
        self, _, __, click
    ):
        result = super_strategy._confirm_success_result_dialog(
            11, "期权合约批量下单处理完毕：成功2笔，失败0笔。"
        )
        click.assert_called_once_with(11, 22, require_closed=True)
        self.assertEqual(result["success_count"], 2)

    @mock.patch("core.super_strategy._click_dialog_button")
    @mock.patch(
        "core.super_strategy._exact_child_buttons",
        side_effect=[[22], [23]],
    )
    @mock.patch("core.super_strategy._native_window_text", return_value="下单")
    def test_order_confirmation_uses_win32_button_message(self, _, __, click):
        super_strategy._confirm_order_dialog(11, "您确定要下单吗？")
        click.assert_called_once_with(11, 22)


class RealMouseClickTests(unittest.TestCase):

    @mock.patch("core.super_strategy._real_mouse_click")
    @mock.patch(
        "core.super_strategy._find_action_text",
        return_value={
            "text": super_strategy.OPEN_POSITION_TEXT,
            "cx": 1050.4,
            "cy": 640.6,
            "score": 0.99,
            "ocr_elapsed": 0.2,
        },
    )
    @mock.patch("core.super_strategy.get_action_panel", return_value=22)
    def test_action_uses_real_mouse_at_ocr_screen_point(self, _, __, real_click):
        super_strategy.click_action(11, super_strategy.OPEN_POSITION_TEXT, delay=0.3)
        real_click.assert_called_once_with(
            11, 22, 1050.4, 640.6, delay=0.3
        )

    @mock.patch("core.super_strategy._cache_detected_action_hits")
    @mock.patch("core.super_strategy._cached_action_hits", return_value=None)
    @mock.patch("core.super_strategy._action_cache_key", return_value="layout")
    @mock.patch("core.super_strategy._action_strip")
    @mock.patch("core.super_strategy.ocr_image_items")
    def test_two_actions_share_one_full_ocr_pass(
        self, ocr_items, action_strip, _, __, cache_hits
    ):
        image = mock.Mock(width=900, height=48)
        action_strip.return_value = (image, (100, 200))
        add_hit = {
            "text": super_strategy.ADD_UNDERLYING_TEXT,
            "left": 150,
            "top": 210,
            "right": 210,
            "bottom": 230,
        }
        open_hit = {
            "text": super_strategy.OPEN_POSITION_TEXT,
            "left": 230,
            "top": 210,
            "right": 290,
            "bottom": 230,
        }
        ocr_items.return_value = [add_hit, open_hit]

        result = super_strategy._find_action_texts(
            11,
            22,
            (
                super_strategy.ADD_UNDERLYING_TEXT,
                super_strategy.OPEN_POSITION_TEXT,
            ),
        )

        self.assertEqual(ocr_items.call_count, 1)
        self.assertFalse(ocr_items.call_args.kwargs["use_cls"])
        self.assertIs(result[super_strategy.ADD_UNDERLYING_TEXT], add_hit)
        self.assertIs(result[super_strategy.OPEN_POSITION_TEXT], open_hit)
        cache_hits.assert_called_once()

    @mock.patch("core.super_strategy._cache_detected_action_hits")
    @mock.patch("core.super_strategy.ocr_image_items")
    @mock.patch("core.super_strategy._action_cache_key", return_value="layout")
    @mock.patch("core.super_strategy._action_strip")
    @mock.patch("core.super_strategy._cached_action_hits")
    def test_verified_action_cache_skips_full_text_detection(
        self, cached_hits, action_strip, _, ocr_items, cache_detected
    ):
        image = mock.Mock(width=900, height=48)
        action_strip.return_value = (image, (100, 200))
        cached = {
            super_strategy.OPEN_POSITION_TEXT: {
                "text": super_strategy.OPEN_POSITION_TEXT,
                "score": 0.99,
                "ocr_elapsed": 0.1,
            }
        }
        cached_hits.return_value = cached

        result = super_strategy._find_action_texts(
            11, 22, (super_strategy.OPEN_POSITION_TEXT,)
        )

        self.assertIs(result, cached)
        ocr_items.assert_not_called()
        cache_detected.assert_not_called()

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
    def test_six_formal_targets_have_distinct_fast_ocr_cells(self):
        self.assertEqual(set(tactics_panel.FORMAL_TARGET_CELLS), tactics_panel.SUPER_TARGETS)
        self.assertEqual(
            len(set(tactics_panel.FORMAL_TARGET_CELLS.values())),
            len(tactics_panel.SUPER_TARGETS),
        )

    def test_shared_driver_and_menu_support_the_same_targets(self):
        self.assertEqual(
            set(super_strategy.SUPER_STRATEGY_TARGETS),
            tactics_panel.SUPER_TARGETS,
        )

    def test_neutral_strategy_targets_use_verified_top_page_cells(self):
        self.assertEqual(
            tactics_panel.FORMAL_TARGET_CELLS["卖出跨式"],
            (5, 540, 80, 582),
        )
        self.assertEqual(
            tactics_panel.FORMAL_TARGET_CELLS["卖宽跨式"],
            (82, 540, 158, 582),
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
