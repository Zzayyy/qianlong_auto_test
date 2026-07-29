import importlib.util
import json
from pathlib import Path
import unittest


PROJECT_ROOT = Path(__file__).resolve().parents[1]


class QianlongClientProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.clients = json.loads(
            (PROJECT_ROOT / "clients.json").read_text(encoding="utf-8")
        )["clients"]
        cls.qianlong = next(
            client for client in cls.clients if client["id"] == "qianlong"
        )

    def test_win11_tree_fingerprint_matches_captured_topology(self):
        profile = self.qianlong["native_tree_profile"]
        topology = profile["expected_root_child_counts"]
        self.assertEqual(profile["expected_node_count"], 52)
        self.assertEqual(len(topology), 22)
        self.assertEqual(22 + sum(topology), 52)
        self.assertEqual(topology[18:], [6, 20, 4, 0])

    def test_known_qianlong_paths_have_expected_positions(self):
        positions = self.qianlong["native_tree_profile"]["positions"]
        self.assertEqual(positions[r"\期权下单(新)"], [0])
        self.assertEqual(positions[r"\三键下单"], [1])
        self.assertEqual(positions[r"\四键下单"], [2])
        self.assertEqual(positions[r"\撤单"], [17])
        self.assertEqual(positions[r"\查询\资金查询"], [19, 9])
        self.assertEqual(
            positions[r"\查询\历史行权负债信息"], [19, 19]
        )

    def test_qianlong_declares_one_click_settings_unsupported(self):
        self.assertIn(
            r"\交易系统设置\一键炒单设置",
            self.qianlong["unsupported"],
        )

    def test_gui_hides_one_click_settings_for_qianlong(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        names = [
            script["name"]
            for script in module.get_scripts_config("qianlong")["交易系统设置"]
        ]
        self.assertFalse(any("一键炒单设置" in name for name in names))

    def test_gui_shows_three_key_panel_for_qianlong(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_orders_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        names = [
            script["name"]
            for script in module.get_scripts_config("qianlong")["下单"]
        ]
        self.assertTrue(any("三键下单" in name for name in names))

    def test_gui_registers_quick_order_script(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_quick_order_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        scripts = module.get_scripts_config("qianlong")["下单"]
        quick_order = next(
            script for script in scripts if "快速下单_自动化下单" in script["name"]
        )
        self.assertTrue(Path(quick_order["path"]).is_file())

        guotai_names = [
            script["name"]
            for script in module.get_scripts_config("guotai_haitong")["下单"]
        ]
        self.assertTrue(any("快速下单_自动化下单" in name for name in guotai_names))

    def test_gui_exposes_three_top_level_modules(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_modules_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        self.assertEqual(
            list(module.MODULE_GROUPS),
            ["行情交易", "超级策略", "交易系统设置"],
        )
        self.assertEqual(
            tuple(module.MODULE_GROUPS["超级策略"]),
            ("超级策略",),
        )
        self.assertEqual(
            module.SUPER_STRATEGY_CATEGORIES,
            frozenset(("超级策略",)),
        )
        scripts = module.get_scripts_config("guotai_haitong")
        self.assertEqual(
            [script["name"] for script in scripts["超级策略"]],
            ["牛市认沽", "牛市认购", "熊市认沽", "熊市认购"],
        )
        for script in scripts["超级策略"]:
            self.assertTrue(script["path"].endswith(
                f"超级策略\\{script['name']}_一键开仓.py"
            ))

    def test_both_clients_declare_workspace_fingerprints(self):
        for client in self.clients:
            with self.subTest(client=client["id"]):
                profile = client["workspace"]
                self.assertEqual(profile["market_button_id"], 1019)
                self.assertEqual(profile["super_button_id"], 1023)
                self.assertEqual(profile["tactics_panel_id"], 103)


if __name__ == "__main__":
    unittest.main()
