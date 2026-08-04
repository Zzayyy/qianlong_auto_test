# -*- coding: utf-8 -*-
"""组合申报（超级策略填表）合约派生与解析逻辑的单元测试。

这些测试不依赖真实交易客户端，仅验证：
  - parse_contract 对期权合约代码的解析
  - _pair_matches 对策略两腿配对的判定
  - derive_contract_pair 从下拉候选项按持仓派生配对
"""

import unittest

from core.super_strategy import (
    COMBO_MARKETS,
    COMBO_STRATEGIES,
    derive_contract_pair,
    parse_contract,
    _pair_matches,
)


class ParseContractTests(unittest.TestCase):
    def test_parse_call_contract(self):
        p = parse_contract("50ETF购8月2850")
        self.assertIsNotNone(p)
        self.assertEqual(p["under"], "50ETF")
        self.assertEqual(p["type"], "购")
        self.assertEqual(p["month"], "8月")
        self.assertEqual(p["strike"], 2850)

    def test_parse_put_contract(self):
        p = parse_contract("300ETF沽12月4300")
        self.assertIsNotNone(p)
        self.assertEqual(p["under"], "300ETF")
        self.assertEqual(p["type"], "沽")
        self.assertEqual(p["month"], "12月")
        self.assertEqual(p["strike"], 4300)

    def test_parse_with_chinese_underlying(self):
        p = parse_contract("华夏科创50ETF购9月1200")
        self.assertIsNotNone(p)
        self.assertEqual(p["under"], "华夏科创50ETF")
        self.assertEqual(p["type"], "购")
        self.assertEqual(p["strike"], 1200)

    def test_invalid_contract_returns_none(self):
        self.assertIsNone(parse_contract(""))
        self.assertIsNone(parse_contract("50ETF"))
        self.assertIsNone(parse_contract("不是期权代码"))
        self.assertIsNone(parse_contract("50ETF购8月"))  # 缺行权价


class PairMatchesTests(unittest.TestCase):
    def test_bull_call_spread(self):
        # 同类型、不同行权价
        self.assertTrue(
            _pair_matches("认购牛市价差策略", "50ETF购8月2850", "50ETF购8月3100")
        )
        # 同行权价不算价差
        self.assertFalse(
            _pair_matches("认购牛市价差策略", "50ETF购8月2850", "50ETF购8月2850")
        )
        # 不同类型不算
        self.assertFalse(
            _pair_matches("认购牛市价差策略", "50ETF购8月2850", "50ETF沽8月3100")
        )

    def test_bear_put_spread(self):
        self.assertTrue(
            _pair_matches("认沽熊市价差策略", "50ETF沽8月2950", "50ETF沽8月2850")
        )
        self.assertFalse(
            _pair_matches("认沽熊市价差策略", "50ETF购8月2950", "50ETF购8月2850")
        )

    def test_straddle_short(self):
        # 一购一沽、同行权价
        self.assertTrue(
            _pair_matches("跨式空头策略", "50ETF购8月3100", "50ETF沽8月3100")
        )
        # 不同行权价不算跨式
        self.assertFalse(
            _pair_matches("跨式空头策略", "50ETF购8月3000", "50ETF沽8月3300")
        )

    def test_strangle_short(self):
        # 一购一沽、不同行权价
        self.assertTrue(
            _pair_matches("宽跨式空头策略", "50ETF购8月3000", "50ETF沽8月3300")
        )
        # 同行权价不算宽跨式
        self.assertFalse(
            _pair_matches("宽跨式空头策略", "50ETF购8月3100", "50ETF沽8月3100")
        )

    def test_different_underlying_or_month(self):
        self.assertFalse(
            _pair_matches("认购牛市价差策略", "50ETF购8月2850", "300ETF购8月3100")
        )
        self.assertFalse(
            _pair_matches("认购牛市价差策略", "50ETF购8月2850", "50ETF购9月3100")
        )


class DeriveContractPairTests(unittest.TestCase):
    def test_bull_call_spread_derives_low_then_high(self):
        c1 = ["50ETF购8月2850", "50ETF购8月3100", "50ETF购8月3400"]
        c2 = ["50ETF购8月2850", "50ETF购8月3100", "50ETF购8月3400"]
        pair = derive_contract_pair("认购牛市价差策略", c1, c2)
        self.assertIsNotNone(pair)
        self.assertEqual(pair, ("50ETF购8月2850", "50ETF购8月3100"))

    def test_bull_put_spread_derives(self):
        c1 = ["50ETF沽8月2850", "50ETF沽8月2950"]
        c2 = ["50ETF沽8月2850", "50ETF沽8月2950"]
        pair = derive_contract_pair("认沽牛市价差策略", c1, c2)
        self.assertIsNotNone(pair)
        self.assertEqual(pair, ("50ETF沽8月2850", "50ETF沽8月2950"))

    def test_straddle_short_derives_call_and_put_same_strike(self):
        c1 = ["50ETF购8月3100", "50ETF购8月3000"]
        c2 = ["50ETF沽8月3100", "50ETF沽8月3300"]
        pair = derive_contract_pair("跨式空头策略", c1, c2)
        self.assertIsNotNone(pair)
        # 跨式以同行权价的一购一沽配对
        self.assertEqual(pair, ("50ETF购8月3100", "50ETF沽8月3100"))

    def test_strangle_short_derives_call_and_put_diff_strike(self):
        c1 = ["50ETF购8月3000", "50ETF购8月3200"]
        c2 = ["50ETF沽8月3300", "50ETF沽8月3400"]
        pair = derive_contract_pair("宽跨式空头策略", c1, c2)
        self.assertIsNotNone(pair)
        self.assertEqual(pair, ("50ETF购8月3000", "50ETF沽8月3300"))

    def test_anchor_on_contract2_then_pair(self):
        # 合约二下拉里只出现一个合约，应从它反推合约一
        c1 = ["50ETF购8月2850", "50ETF购8月3100", "50ETF购8月3400"]
        c2 = ["50ETF购8月3100"]
        pair = derive_contract_pair("认购牛市价差策略", c1, c2)
        self.assertIsNotNone(pair)
        self.assertEqual(pair, ("50ETF购8月2850", "50ETF购8月3100"))

    def test_no_valid_pair_raises(self):
        from core.super_strategy import SuperStrategyError

        c1 = ["50ETF购8月2850"]
        c2 = ["50ETF沽8月3300"]
        # 购/沽不同且无跨式/宽跨式所需同行权价，牛市价差无法配对 -> 抛错
        with self.assertRaises(SuperStrategyError):
            derive_contract_pair("认购牛市价差策略", c1, c2)


class ComboConstantsTests(unittest.TestCase):
    def test_markets_and_strategies_are_fixed_enums(self):
        self.assertEqual(COMBO_MARKETS, ("上证", "深证"))
        self.assertEqual(
            COMBO_STRATEGIES,
            (
                "认购牛市价差策略",
                "认购熊市价差策略",
                "宽跨式空头策略",
                "跨式空头策略",
                "认沽牛市价差策略",
                "认沽熊市价差策略",
            ),
        )


class BatchDeclareTests(unittest.TestCase):
    def test_batch_iterates_market_times_strategy(self):
        """run_combination_declare_all 对 (市场 × 策略) 笛卡尔积逐个组合。"""
        from unittest.mock import patch

        from core import super_strategy as ss

        captured = []
        fake_result = {"main_hwnd": 999, "submitted": True, "closed": False}

        def fake_declare(*, market, strategy, qty=1, execute=True,
                         cleanup_leftover=True, close_after=False,
                         do_setup=True, main_hwnd=None, combo_dlg=None):
            captured.append((market, strategy, do_setup, main_hwnd))
            return dict(fake_result, market=market, strategy=strategy)

        with patch.object(ss, "run_combination_declare", side_effect=fake_declare):
            results = ss.run_combination_declare_all(
                markets=["上证", "深证"],
                strategies=["认购牛市价差策略", "跨式空头策略"],
                qty=3,
                execute=True,
            )

        # 2 市场 × 2 策略 = 4 个组合
        self.assertEqual(len(results), 4)
        self.assertEqual(len(captured), 4)
        # 首个 do_setup=True，后续 do_setup=False 且复用主窗口句柄
        self.assertTrue(captured[0][2])
        self.assertIsNone(captured[0][3])
        for m, s, do_setup, mh in captured[1:]:
            self.assertFalse(do_setup)
            self.assertEqual(mh, 999)
        # 组合覆盖全部笛卡尔积
        pairs = {(m, s) for m, s, _, _ in captured}
        self.assertEqual(
            pairs,
            {
                ("上证", "认购牛市价差策略"),
                ("上证", "跨式空头策略"),
                ("深证", "认购牛市价差策略"),
                ("深证", "跨式空头策略"),
            },
        )

    def test_batch_defaults_to_first_when_empty(self):
        """未选择任何市场/策略时，回退到各首项，避免空跑。"""
        from unittest.mock import patch

        from core import super_strategy as ss

        captured = []
        with patch.object(
            ss, "run_combination_declare",
            side_effect=lambda **kw: captured.append(kw) or {"main_hwnd": 1},
        ):
            ss.run_combination_declare_all(markets=[], strategies=[], qty=1)
        self.assertEqual(len(captured), 1)
        self.assertEqual(captured[0]["market"], COMBO_MARKETS[0])
        self.assertEqual(captured[0]["strategy"], COMBO_STRATEGIES[0])

    def test_is_combo_dialog_alive_checks_handle_and_key_control(self):
        """对话框复用前必须校验：句柄有效+可见+含市场下拉框(9059)。

        客户端提交成功后可能自动关闭申报对话框（华宝实测），旧句柄失效；
        Windows 句柄也可能被新窗口重用，故必须校验关键控件存在。
        """
        from unittest.mock import patch

        from core import super_strategy as ss

        with patch.object(ss.win32gui, "IsWindow", return_value=True), \
             patch.object(ss.win32gui, "IsWindowVisible", return_value=True), \
             patch.object(ss.win32gui, "GetDlgItem", return_value=123):
            self.assertTrue(ss._is_combo_dialog_alive(999))
        # 不可见 → 失效（提交后客户端自动关闭的场景）
        with patch.object(ss.win32gui, "IsWindow", return_value=True), \
             patch.object(ss.win32gui, "IsWindowVisible", return_value=False), \
             patch.object(ss.win32gui, "GetDlgItem", return_value=123):
            self.assertFalse(ss._is_combo_dialog_alive(999))
        # 缺关键控件（句柄被新窗口重用）→ 失效
        with patch.object(ss.win32gui, "IsWindow", return_value=True), \
             patch.object(ss.win32gui, "IsWindowVisible", return_value=True), \
             patch.object(ss.win32gui, "GetDlgItem", return_value=0):
            self.assertFalse(ss._is_combo_dialog_alive(999))
        self.assertFalse(ss._is_combo_dialog_alive(0))
        self.assertFalse(ss._is_combo_dialog_alive(None))

    def test_declare_reuses_live_dialog(self):
        """对话框仍有效时走复用路径，不重新点击打开。"""
        from unittest.mock import patch

        from core import super_strategy as ss

        with patch.object(ss, "_is_combo_dialog_alive", return_value=True), \
             patch.object(ss, "_activate_main_window"), \
             patch.object(ss, "click_action") as click, \
             patch.object(ss, "_fill_combo_dialog",
                          return_value={"dialog_hwnd": 789}) as fill:
            res = ss.run_combination_declare(
                market="深证", strategy="认沽牛市价差策略", qty=1,
                execute=True, do_setup=False, main_hwnd=111, combo_dlg=789,
            )
        click.assert_not_called()
        self.assertEqual(res["dialog_hwnd"], 789)
        self.assertEqual(fill.call_args.args[0], 789)

    def test_declare_reopens_dialog_when_reused_handle_dead(self):
        """提交后客户端自动关闭对话框（华宝实测）时，自动重新打开再填表。"""
        from unittest.mock import patch

        from core import super_strategy as ss

        with patch.object(ss, "_is_combo_dialog_alive", return_value=False), \
             patch.object(ss, "_close_combo_leftover_dialogs"), \
             patch.object(ss, "get_action_panel", return_value=123), \
             patch.object(ss, "click_action") as click, \
             patch.object(ss, "_find_combination_dialog", return_value=456) as find_dlg, \
             patch.object(ss, "_fill_combo_dialog",
                          return_value={"dialog_hwnd": 456}) as fill:
            res = ss.run_combination_declare(
                market="深证", strategy="认沽牛市价差策略", qty=1,
                execute=True, cleanup_leftover=False, close_after=False,
                do_setup=False, main_hwnd=111, combo_dlg=789,
            )
        click.assert_called_once()
        find_dlg.assert_called_once()
        self.assertEqual(res["dialog_hwnd"], 456)
        self.assertEqual(fill.call_args.args[0], 456)


if __name__ == "__main__":
    unittest.main(verbosity=2)
