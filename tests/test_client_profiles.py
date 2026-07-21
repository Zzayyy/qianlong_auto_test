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

    def test_gui_hides_missing_three_key_panel_for_qianlong(self):
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
        self.assertFalse(any("三键下单" in name for name in names))


if __name__ == "__main__":
    unittest.main()
