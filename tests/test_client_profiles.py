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

    def test_gui_registers_normal_order_script_for_zhongtai_only(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_normal_order_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        # 中泰：普通下单 = 钱龙/国泰的快速下单界面，复用快速下单脚本
        zhongtai_scripts = module.get_scripts_config("zhongtai")["下单"]
        normal_order = next(
            script
            for script in zhongtai_scripts
            if "普通下单_自动化下单" in script["name"]
        )
        self.assertTrue(Path(normal_order["path"]).is_file())
        self.assertTrue(
            normal_order["path"].endswith(
                "4.快速下单_自动化下单_Excel驱动版.py"
            )
        )

        # 钱龙/国泰没有“普通下单”菜单，应被过滤
        for client_id in ("qianlong", "guotai_haitong"):
            with self.subTest(client=client_id):
                names = [
                    script["name"]
                    for script in module.get_scripts_config(client_id)["下单"]
                ]
                self.assertFalse(
                    any("普通下单" in name for name in names),
                    f"{client_id} 不应显示普通下单条目",
                )

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
            [
                "牛市认沽",
                "牛市认购",
                "熊市认沽",
                "熊市认购",
                "卖出跨式",
                "卖宽跨式",
                "组合申报",
            ],
        )
        for script in scripts["超级策略"]:
            if script["name"] == "组合申报":
                expected_suffix = f"超级策略\\{script['name']}.py"
            else:
                expected_suffix = f"超级策略\\{script['name']}_一键开仓.py"
            self.assertTrue(script["path"].endswith(expected_suffix))
            self.assertTrue(Path(script["path"]).is_file())

    def test_both_clients_declare_workspace_fingerprints(self):
        for client in self.clients:
            with self.subTest(client=client["id"]):
                profile = client["workspace"]
                self.assertEqual(profile["market_button_id"], 1019)
                self.assertEqual(profile["super_button_id"], 1023)
                self.assertEqual(profile["tactics_panel_id"], 103)


class HuabaoClientProfileTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.clients = json.loads(
            (PROJECT_ROOT / "clients.json").read_text(encoding="utf-8")
        )["clients"]
        cls.huabao = next(
            client for client in cls.clients if client["id"] == "huabao"
        )

    def test_win11_tree_fingerprint_matches_captured_topology(self):
        profile = self.huabao["native_tree_profile"]
        topology = profile["expected_root_child_counts"]
        # 华宝与钱龙同为 52 节点，拓扑必须不同，否则位置定位会错位
        self.assertEqual(profile["expected_node_count"], 52)
        self.assertEqual(len(topology), 24)
        self.assertEqual(24 + sum(topology), 52)
        self.assertEqual(topology[18:], [8, 16, 4, 0, 0, 0])

    def test_known_huabao_paths_have_expected_positions(self):
        positions = self.huabao["native_tree_profile"]["positions"]
        self.assertEqual(positions[r"\期权下单"], [0])
        self.assertEqual(positions[r"\三键下单"], [1])
        # 华宝止盈止损位于第5位（与钱龙/国泰不同）
        self.assertEqual(positions[r"\止盈止损"], [4])
        self.assertEqual(positions[r"\快速下单"], [5])
        self.assertEqual(positions[r"\协议行权"], [16])
        self.assertEqual(positions[r"\撤单"], [17])
        self.assertEqual(positions[r"\组合申报\组合历史成交"], [18, 7])
        self.assertEqual(positions[r"\查询\资金持仓"], [19, 0])
        self.assertEqual(positions[r"\查询\账号查询"], [19, 15])
        self.assertEqual(positions[r"\通知查询\客户账单(结算清单)"], [20, 3])

    def test_huabao_menu_map_redirects_combo_query_names(self):
        menu_map = self.huabao["menu_map"]
        self.assertEqual(
            menu_map[r"\组合申报\组合策略持仓查询"],
            r"\组合申报\组合策略持仓",
        )
        self.assertEqual(
            menu_map[r"\组合申报\组合策略信息查询"],
            r"\组合申报\组合策略信息",
        )
        self.assertEqual(
            menu_map[r"\组合申报\组合委托流水查询"],
            r"\组合申报\组合当日委托",
        )
        self.assertEqual(
            menu_map[r"\组合申报\历史组合委托流水"],
            r"\组合申报\组合历史委托",
        )

    def test_huabao_unsupported_covers_missing_menus(self):
        unsupported = self.huabao["unsupported"]
        # 华宝无结算单/通知查询（合约变更等）/行权类查询/普通下单等菜单
        self.assertIn(r"\结算单\结算单", unsupported)
        self.assertIn(r"\通知查询\合约信息变更数量", unsupported)
        self.assertIn(r"\查询\当日行权被指派查询", unsupported)
        self.assertIn(r"\查询\客户限仓信息查询", unsupported)
        self.assertIn(r"\查询\限购额度查询", unsupported)
        self.assertIn(r"\查询\可用备兑股份", unsupported)
        self.assertIn(r"\查询\备兑股份不足", unsupported)
        self.assertIn(r"\组合申报\历史组合策略持仓查询", unsupported)
        self.assertIn(r"\下单\普通下单_自动化下单", unsupported)
        self.assertIn(r"\交易系统设置\一键炒单设置", unsupported)
        # 华宝有账号查询菜单，不得误过滤
        self.assertNotIn(r"\查询\账号查询", unsupported)

    def test_huabao_menu_alias_option_order_new(self):
        self.assertEqual(self.huabao["menu_aliases"]["期权下单(新)"], "期权下单")

    def test_gui_hides_unsupported_categories_for_huabao(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_huabao_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        scripts = module.get_scripts_config("huabao")

        # 结算单在华宝下应全部被过滤为空；通知查询保留华宝专属 3 项
        # （客户账单(结算清单) 输出逻辑与查询驱动不通用，已剔除）
        self.assertEqual(scripts["结算单"], [])
        notify_names = [script["name"] for script in scripts["通知查询"]]
        self.assertEqual(len(notify_names), 3)
        for item in ("行权待交收证券缺口", "备兑证券缺口查询", "行权待交收资金缺口"):
            self.assertTrue(any(item in n for n in notify_names))
        self.assertFalse(any("客户账单" in n for n in notify_names))
        # 国泰/中泰的通知查询项（合约变更/风险通知）应被过滤
        for item in ("合约信息变更", "风险通知"):
            self.assertFalse(any(item in n for n in notify_names))

        # 查询类目保留华宝有的 16 项（与华宝树查询子菜单一一对应）
        query_names = [script["name"] for script in scripts["查询"]]
        self.assertEqual(len(query_names), 16)
        bare_query = {n.split(". ", 1)[-1].strip() for n in query_names}
        self.assertTrue(any("资金持仓" in name for name in query_names))
        self.assertTrue(any("账号查询" in name for name in query_names))
        # 华宝专属 5 项：备兑股份/历史损益/待交收查询/额度查询/持仓限制查询
        for item in ("备兑股份", "历史损益", "待交收查询", "额度查询", "持仓限制查询"):
            self.assertIn(item, bare_query)
        # 华宝树没有的菜单项必须被过滤
        for item in ("当日行权被指派", "历史对账单查询", "可用备兑股份", "备兑股份不足",
                     "限购额度查询", "客户限仓信息查询", "对账单资金流水", "对账单资金资产"):
            self.assertNotIn(item, bare_query)

        # 下单类目保留华宝有的项、过滤普通下单/一键导出
        order_names = [script["name"] for script in scripts["下单"]]
        self.assertTrue(any("三键下单" in name for name in order_names))
        self.assertTrue(any("四键下单" in name for name in order_names))
        self.assertFalse(any("普通下单" in name for name in order_names))
        self.assertFalse(any("一键导出" in name for name in order_names))

        # 组合申报类目：映射项保留（显示名为华宝真实菜单名）+ 华宝专属两项
        combo_names = [script["name"] for script in scripts["组合申报"]]
        self.assertTrue(any("组合策略持仓" in name for name in combo_names))
        self.assertTrue(any("组合策略信息" in name for name in combo_names))
        self.assertTrue(any("组合当日委托" in name for name in combo_names))
        self.assertTrue(any("组合历史委托" in name for name in combo_names))
        self.assertTrue(any("组合当日成交" in name for name in combo_names))
        self.assertTrue(any("组合历史成交" in name for name in combo_names))
        self.assertEqual(len(combo_names), 8)
        self.assertFalse(any("历史组合策略持仓" in name for name in combo_names))

    def test_huabao_combo_orders_filtered_for_other_clients(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_combo_filter_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        for client in self.clients:
            if client["id"] == "huabao":
                continue
            with self.subTest(client=client["id"]):
                scripts = module.get_scripts_config(client["id"])
                names = [s["name"] for s in scripts.get("组合申报", [])]
                self.assertFalse(any("组合当日成交" in n for n in names))
                self.assertFalse(any("组合历史成交" in n for n in names))

    def test_huabao_query_orders_filtered_for_other_clients(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_query_filter_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        for client in self.clients:
            if client["id"] == "huabao":
                continue
            with self.subTest(client=client["id"]):
                names = [s["name"] for s in module.get_scripts_config(client["id"]).get("查询", [])]
                bare = {n.split(". ", 1)[-1].strip() for n in names}
                for item in ("备兑股份", "历史损益", "待交收查询", "额度查询", "持仓限制查询"):
                    self.assertNotIn(item, bare)

    def test_huabao_notify_orders_filtered_for_other_clients(self):
        config_path = PROJECT_ROOT / "GUI自动化工具2" / "config.py"
        spec = importlib.util.spec_from_file_location(
            "gui_automation_config_notify_filter_for_test", config_path
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)
        for client in self.clients:
            if client["id"] == "huabao":
                continue
            with self.subTest(client=client["id"]):
                names = [s["name"] for s in module.get_scripts_config(client["id"]).get("通知查询", [])]
                for item in ("行权待交收证券缺口", "备兑证券缺口查询", "行权待交收资金缺口"):
                    self.assertFalse(any(item in n for n in names))

    def test_huabao_query_export_flows(self):
        """华宝资金持仓→数据输出弹窗、期权合约→系统级另存为（与其他客户端不同）。"""
        import json as _json

        table = _json.loads(
            (PROJECT_ROOT / "行情交易" / "查询" / "queries.json").read_text(encoding="utf-8")
        )
        spec = importlib.util.spec_from_file_location(
            "run_query_huabao_export_for_test",
            PROJECT_ROOT / "行情交易" / "查询" / "run_query.py",
        )
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(module)

        # 华宝资金持仓：export_type 被删除 → run_export_dialog（数据输出弹窗）
        entry = module.apply_client_override(table[r"\查询\资金持仓"], "huabao")
        self.assertNotIn("export_type", entry)

        # 华宝期权合约：export_type=xls_only → run_save_as（系统级另存为）
        entry = module.apply_client_override(table[r"\查询\期权合约"], "huabao")
        self.assertEqual(entry["export_type"], "xls_only")

        # 华宝其他查询项默认走数据输出弹窗
        entry = module.apply_client_override(table[r"\查询\当日成交"], "huabao")
        self.assertNotIn("export_type", entry)


if __name__ == "__main__":
    unittest.main()
