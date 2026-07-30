# -*- coding: utf-8 -*-

from pathlib import Path
import sys
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
GUI_ROOT = PROJECT_ROOT / "GUI自动化工具2"
if str(GUI_ROOT) not in sys.path:
    sys.path.insert(0, str(GUI_ROOT))

from gui.main_window import AutomationGUI
from gui.task_center import TaskCenter
from config import SUPER_STRATEGY_UNDERLYINGS


class _Value:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value


class _SourceTree:
    def __init__(self, iid):
        self.iid = iid

    def identify_row(self, _y):
        return self.iid


class _TargetTree:
    def winfo_rootx(self):
        return 100

    def winfo_rooty(self):
        return 200

    def winfo_width(self):
        return 300

    def winfo_height(self):
        return 240


class SuperStrategyParameterUiTests(unittest.TestCase):
    def test_parameter_panel_uses_readonly_underlying_selector(self):
        gui = AutomationGUI.__new__(AutomationGUI)
        gui.params_frame = Mock()
        gui.super_strategy_underlying = _Value(SUPER_STRATEGY_UNDERLYINGS[0])
        gui.super_add_underlying = _Value(False)
        with (
            patch("gui.main_window.ttk.Label") as label,
            patch("gui.main_window.ttk.Combobox") as combobox,
            patch("gui.main_window.ttk.Checkbutton") as checkbutton,
        ):
            AutomationGUI._build_super_strategy_params(gui)

        combobox.assert_called_once_with(
            gui.params_frame,
            textvariable=gui.super_strategy_underlying,
            values=SUPER_STRATEGY_UNDERLYINGS,
            state="readonly",
            width=28,
        )
        label.assert_called()
        checkbutton.assert_called_once()
        gui.params_frame.columnconfigure.assert_called_once_with(1, weight=1)

    def test_combination_declare_builds_checkbox_panel(self):
        """组合申报（超级策略分类）参数面板用复选框多选市场/策略 + 组合数量输入框。"""
        from config import SUPER_STRATEGY_COMBO_MARKETS, SUPER_STRATEGY_COMBO_STRATEGIES

        gui = AutomationGUI.__new__(AutomationGUI)
        gui.params_frame = Mock()
        gui.super_combo_market_vars = {
            m: _Value(m == SUPER_STRATEGY_COMBO_MARKETS[0])
            for m in SUPER_STRATEGY_COMBO_MARKETS
        }
        gui.super_combo_strategy_vars = {
            s: _Value(s == SUPER_STRATEGY_COMBO_STRATEGIES[0])
            for s in SUPER_STRATEGY_COMBO_STRATEGIES
        }
        gui.super_combo_qty = _Value(1)
        gui.script_tree = Mock()
        gui.script_tree.selection.return_value = ["script::组合申报"]
        gui.tree_script_map = {
            "script::组合申报": {"script": {"name": "组合申报"}},
        }
        with (
            patch("gui.main_window.ttk.Label") as label,
            patch("gui.main_window.ttk.Frame") as frame,
            patch("gui.main_window.ttk.Combobox") as combobox,
            patch("gui.main_window.ttk.Checkbutton") as checkbutton,
            patch("gui.main_window.ttk.Entry") as entry,
        ):
            AutomationGUI._build_super_strategy_params(gui)

        # 不应构造下拉框（改用复选框）
        combobox.assert_not_called()
        # 每个市场/策略各一个复选框，无“加入标的”勾选框
        self.assertEqual(
            checkbutton.call_count,
            len(SUPER_STRATEGY_COMBO_MARKETS) + len(SUPER_STRATEGY_COMBO_STRATEGIES),
        )
        # 应构造组合数量 Entry
        entry.assert_called_once()
        # 不应出现“无需参数配置”提示
        texts = [
            c.kwargs.get("text")
            for c in label.call_args_list
            if c.kwargs.get("text")
        ]
        self.assertFalse(
            any("无需参数配置" in (t or "") for t in texts),
            msg=f"不应出现无需参数配置提示, labels={texts}",
        )


class MainWindowDirectoryDragTests(unittest.TestCase):
    def _gui(self, iid):
        gui = AutomationGUI.__new__(AutomationGUI)
        gui.task_center = SimpleNamespace(is_running=False)
        gui.script_tree = _SourceTree(iid)
        gui.tree_script_map = {}
        gui._drag_script = {"stale": True}
        gui._drag_category = "stale"
        gui._drag_module = "stale"
        gui._drag_active = True
        gui._drag_occurred = True
        gui._drag_start_y = 0
        return gui

    def test_top_level_module_becomes_drag_payload(self):
        gui = self._gui("module::行情交易")

        gui._on_list_drag_start(SimpleNamespace(y=12))

        self.assertEqual("行情交易", gui._drag_module)
        self.assertIsNone(gui._drag_category)
        self.assertIsNone(gui._drag_script)
        self.assertFalse(gui._drag_active)
        self.assertFalse(gui._drag_occurred)

    def test_unknown_tree_node_clears_stale_drag_state(self):
        gui = self._gui("unknown::node")

        gui._on_list_drag_start(SimpleNamespace(y=12))

        self.assertIsNone(gui._drag_module)
        self.assertIsNone(gui._drag_category)
        self.assertIsNone(gui._drag_script)
        self.assertFalse(gui._drag_active)

    def test_global_drop_dispatches_module_to_task_center(self):
        calls = []
        task_center = SimpleNamespace(
            is_running=False,
            tree=_TargetTree(),
            _hide_drop_indicator=lambda: None,
            add_module_from_drop=lambda module, y: calls.append((module, y)),
        )
        gui = AutomationGUI.__new__(AutomationGUI)
        gui.task_center = task_center
        gui._drag_script = None
        gui._drag_category = None
        gui._drag_module = "超级策略"
        gui._drag_active = True

        gui._on_global_drop(SimpleNamespace(x_root=150, y_root=250))

        self.assertEqual([("超级策略", 50)], calls)
        self.assertIsNone(gui._drag_module)
        self.assertFalse(gui._drag_active)


class TaskCenterBatchDragTests(unittest.TestCase):
    def _task_center(self, xlsx_file=""):
        logs = []
        controller = SimpleNamespace(
            client_id="test",
            xlsx_file=_Value(xlsx_file),
            make_task_params=lambda script, category: {
                "source": script["name"],
                "category": category,
            },
        )
        center = TaskCenter.__new__(TaskCenter)
        center.is_running = False
        center.controller = controller
        center.gui = SimpleNamespace(_log=logs.append)
        center.tasks = [
            {
                "category": "已有",
                "script_name": "已有1",
                "script_path": "existing-1.py",
                "params": {},
                "status": center.ST_PENDING,
            },
            {
                "category": "已有",
                "script_name": "已有2",
                "script_path": "existing-2.py",
                "params": {},
                "status": center.ST_PENDING,
            },
        ]
        center._save = Mock()
        center._refresh = Mock()
        center._update_group_hint = Mock()
        center._dirty = False
        return center, logs

    def test_module_adds_all_categories_in_configured_order_once(self):
        center, logs = self._task_center()
        scripts = {
            "分类A": [
                {"name": "A1", "path": "a1.py"},
                {"name": "A工具", "path": "tool.py", "exclude_from_batch": True},
            ],
            "分类B": [
                {"name": "B缺失", "path": "missing.py"},
                {"name": "B1", "path": "b1.py"},
            ],
        }

        with (
            patch("gui.task_center.get_scripts_config", return_value=scripts),
            patch.dict(
                "gui.task_center.MODULE_GROUPS",
                {"测试模块": ("分类A", "分类B")},
                clear=True,
            ),
            patch(
                "gui.task_center.os.path.exists",
                side_effect=lambda path: path != "missing.py",
            ),
        ):
            added = center.add_module("测试模块", index=1)

        self.assertEqual(2, added)
        self.assertEqual(
            ["已有1", "A1", "B1", "已有2"],
            [item["script_name"] for item in center.tasks],
        )
        center._save.assert_called_once_with()
        center._refresh.assert_called_once_with()
        center._update_group_hint.assert_called_once_with()
        self.assertTrue(center._dirty)
        self.assertIn("模块「测试模块」已批量加入 2 个脚本", logs[-1])
        self.assertIn("跳过 2 个", logs[-1])
        self.assertIn("批量排除 1 个", logs[-1])
        self.assertIn("脚本文件不存在 1 个", logs[-1])

    def test_category_reports_excel_dependent_scripts_as_skipped(self):
        center, logs = self._task_center(xlsx_file="")
        scripts = {
            "下单": [
                {"name": "1.普通下单", "path": "order.py"},
                {"name": "2.全选撤单", "path": "cancel.py"},
            ]
        }

        with (
            patch("gui.task_center.get_scripts_config", return_value=scripts),
            patch("gui.task_center.os.path.exists", return_value=True),
        ):
            added = center.add_category("下单")

        self.assertEqual(1, added)
        self.assertEqual("全选撤单", center.tasks[-1]["script_name"].split(".", 1)[-1])
        self.assertIn("未选择 Excel 配置文件 1 个", logs[-1])

    def test_empty_queue_displays_drop_indicator(self):
        placed = []
        center = TaskCenter.__new__(TaskCenter)
        center.tree = SimpleNamespace(
            get_children=lambda: (),
            winfo_height=lambda: 120,
        )
        center.drop_indicator = SimpleNamespace(
            place=lambda **kwargs: placed.append(kwargs),
            place_forget=lambda: None,
        )

        center._update_drop_indicator(60)

        self.assertEqual(
            [{"x": 0, "y": 60, "relwidth": 1.0, "height": 2}],
            placed,
        )

    def test_saved_task_path_is_resolved_after_script_move(self):
        center, _logs = self._task_center()
        current_path = str(PROJECT_ROOT / "超级策略" / "牛市认购_一键开仓.py")
        row = {
            "category": "超级策略",
            "script_name": "牛市认购",
            "script_path": str(
                PROJECT_ROOT / "超级策略" / "牛市认购" / "牛市认购_一键开仓.py"
            ),
        }
        scripts = {
            "超级策略": [
                {"name": "牛市认购", "path": current_path},
            ]
        }

        with patch("gui.task_center.get_scripts_config", return_value=scripts):
            resolved = center._resolve_script_path(row)

        self.assertEqual(current_path, resolved)


if __name__ == "__main__":
    unittest.main()
