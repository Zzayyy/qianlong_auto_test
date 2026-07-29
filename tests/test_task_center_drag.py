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


if __name__ == "__main__":
    unittest.main()
