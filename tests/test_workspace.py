# -*- coding: utf-8 -*-

import unittest
from unittest import mock

from core import workspace


class WorkspaceRoutingTests(unittest.TestCase):
    def test_market_categories_route_to_market_workspace(self):
        for category in ("查询", "通知查询", "结算单", "下单", "组合申报"):
            with self.subTest(category=category):
                self.assertEqual(
                    workspace.workspace_for_category(category),
                    workspace.WORKSPACE_MARKET,
                )

    def test_super_categories_route_to_super_workspace(self):
        for category in (
            "超级策略", "牛市认购", "牛市认沽", "熊市认购", "熊市认沽"
        ):
            with self.subTest(category=category):
                self.assertEqual(
                    workspace.workspace_for_category(category),
                    workspace.WORKSPACE_SUPER,
                )

    def test_settings_does_not_force_workspace_switch(self):
        self.assertIsNone(workspace.workspace_for_category("交易系统设置"))

    @mock.patch("core.workspace.win32gui.SendMessage")
    @mock.patch("core.workspace._single_control", return_value=456)
    @mock.patch("core.workspace.is_workspace_ready", side_effect=[False, True])
    def test_ensure_workspace_clicks_profile_button_and_verifies_target(
        self, ready, single, send
    ):
        result = workspace.ensure_workspace(
            123,
            workspace.WORKSPACE_SUPER,
            {"workspace": {"super_button_id": 9001}},
            timeout=0.2,
        )
        single.assert_called_once_with(123, class_name="Button", control_id=9001)
        send.assert_called_once()
        self.assertTrue(result["changed"])
        self.assertGreaterEqual(ready.call_count, 2)


if __name__ == "__main__":
    unittest.main()
