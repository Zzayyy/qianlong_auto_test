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
from gui.task_center import TaskCenter, GROUP_PLACEHOLDER
from gui.scheduler_view import AddSchedDialog
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


class TaskCenterReorderTests(unittest.TestCase):
    """任务中心内部拖拽重排（_reorder）的边界场景。"""

    def _center(self, names, same_path=True):
        center = TaskCenter.__new__(TaskCenter)
        center.is_running = False
        center.tasks = [
            {
                "category": "查询",
                "script_name": name,
                # 查询类脚本共用同一个驱动文件（run_query.py），path 完全相同
                "script_path": (r"X:\行情交易\查询\run_query.py" if same_path
                                else f"X:\\{name}.py"),
                "query_key": f"\\查询\\{name}",
                "params": {},
                "status": center.ST_PENDING,
            }
            for name in names
        ]
        center._save = Mock()
        center._refresh = Mock()
        center._update_group_hint = Mock()
        center._dirty = False
        center.tree = Mock()
        return center

    def test_reorder_works_when_all_tasks_share_same_script_path(self):
        """共用同一驱动文件的队列（如全查询脚本），拖动也必须换位。

        回归：此前用 script_path 列表判断"是否无变化"，同路径脚本重排前后
        path 序列不变，被误判为无变化而 return，导致拖拽不生效。
        """
        center = self._center(["委托查询", "持仓查询", "资金查询", "委托流水查询"])
        center._reorder(0, 1, True)  # 委托查询 拖到 持仓查询 之后
        self.assertEqual(
            ["持仓查询", "委托查询", "资金查询", "委托流水查询"],
            [t["script_name"] for t in center.tasks],
        )
        center._save.assert_called_once_with()
        center._refresh.assert_called_once_with()
        center._update_group_hint.assert_called_once_with()
        self.assertTrue(center._dirty)

    def test_reorder_same_position_is_noop(self):
        """拖回原位置不应触发保存/刷新。"""
        center = self._center(["A", "B", "C"])
        center._reorder(0, 0, False)
        center._save.assert_not_called()
        center._refresh.assert_not_called()
        self.assertFalse(center._dirty)

    def test_reorder_up_and_down_with_distinct_paths(self):
        """不同路径脚本的常规上移/下移仍正常。"""
        center = self._center(["A", "B", "C", "D"], same_path=False)
        center._reorder(3, 0, False)  # D 拖到 A 之前
        self.assertEqual(
            ["D", "A", "B", "C"], [t["script_name"] for t in center.tasks]
        )
        center._reorder(1, 2, True)  # A 拖到 B 之后（当前第 2 位是 B）
        self.assertEqual(
            ["D", "B", "A", "C"], [t["script_name"] for t in center.tasks]
        )


class TaskCenterGroupClientFilterTests(unittest.TestCase):
    """编队按客户端过滤 + 保存时记录 client_id。"""

    def _center(self, client_id="qianlong"):
        center = TaskCenter.__new__(TaskCenter)
        center.is_running = False
        center.parent = None
        center.show_all_var = _Value(False)
        center.controller = SimpleNamespace(client_id=client_id)
        center.gui = SimpleNamespace(
            _log=lambda *a, **k: None, logger=Mock(),
        )
        center.groups = []
        center.tasks = []
        center._dirty = False
        return center

    def _groups(self):
        return [
            {"name": "通用编队", "client_id": ""},
            {"name": "钱龙开盘", "client_id": "qianlong"},
            {"name": "中泰开盘", "client_id": "zhongtai"},
            {"name": "旧版编队"},  # 无 client_id 字段的旧数据视为通用
        ]

    def test_visible_groups_filters_by_current_client(self):
        center = self._center("qianlong")
        center.groups = self._groups()
        names = [g["name"] for g in center._visible_groups()]
        self.assertEqual(["通用编队", "钱龙开盘", "旧版编队"], names)

    def test_visible_groups_show_all_returns_everything(self):
        center = self._center("qianlong")
        center.groups = self._groups()
        center.show_all_var = _Value(True)
        self.assertEqual(4, len(center._visible_groups()))

    def test_visible_groups_other_client(self):
        center = self._center("zhongtai")
        center.groups = self._groups()
        names = [g["name"] for g in center._visible_groups()]
        self.assertEqual(["通用编队", "中泰开盘", "旧版编队"], names)

    def test_save_group_records_current_client_id(self):
        center = self._center("qianlong")
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "query_key": r"\查询\资金持仓",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with patch("gui.task_center.SaveGroupDialog") as dlg_cls:
            dlg_cls.return_value.result_name = "新编队"
            dlg_cls.return_value.result_universal = False
            center.save_group()
        self.assertEqual(1, len(center.groups))
        self.assertEqual("新编队", center.groups[0]["name"])
        self.assertEqual("qianlong", center.groups[0]["client_id"])
        self.assertEqual(r"\查询\资金持仓", center.groups[0]["tasks"][0]["query_key"])

    def test_save_group_universal_stores_empty_client_id(self):
        center = self._center("qianlong")
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with patch("gui.task_center.SaveGroupDialog") as dlg_cls:
            dlg_cls.return_value.result_name = "通用晨检"
            dlg_cls.return_value.result_universal = True
            center.save_group()
        self.assertEqual("", center.groups[0]["client_id"])
        # 通用编队切换到其它客户端仍可见
        center.controller.client_id = "zhongtai"
        names = [g["name"] for g in center._visible_groups()]
        self.assertIn("通用晨检", names)

    def test_save_group_cancel_does_nothing(self):
        center = self._center("qianlong")
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with patch("gui.task_center.SaveGroupDialog") as dlg_cls:
            dlg_cls.return_value.result_name = None  # 取消
            center.save_group()
        self.assertEqual([], center.groups)
        center._save_groups.assert_not_called()

    def test_save_group_overwrite_keeps_original_client_id(self):
        center = self._center("zhongtai")
        center.groups = [{"name": "钱龙开盘", "client_id": "qianlong", "tasks": []}]
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with (
            patch("gui.task_center.SaveGroupDialog") as dlg_cls,
            patch("gui.task_center.messagebox.askyesno", return_value=True),
        ):
            dlg_cls.return_value.result_name = "钱龙开盘"
            dlg_cls.return_value.result_universal = False
            center.save_group()
        self.assertEqual("qianlong", center.groups[0]["client_id"])

    def test_save_group_overwrite_can_make_universal(self):
        center = self._center("zhongtai")
        center.groups = [{"name": "钱龙开盘", "client_id": "qianlong", "tasks": []}]
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with (
            patch("gui.task_center.SaveGroupDialog") as dlg_cls,
            patch("gui.task_center.messagebox.askyesno", return_value=True),
        ):
            dlg_cls.return_value.result_name = "钱龙开盘"
            dlg_cls.return_value.result_universal = True
            center.save_group()
        self.assertEqual("", center.groups[0]["client_id"])

    def test_group_picker_placeholder_name_rejected(self):
        center = self._center("qianlong")
        center.tasks = [
            {
                "category": "查询",
                "script_name": "资金持仓",
                "script_path": r"X:\run_query.py",
                "params": {},
                "status": center.ST_PENDING,
            }
        ]
        center._save_groups = Mock()
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        with (
            patch("gui.task_center.SaveGroupDialog") as dlg_cls,
            patch("gui.task_center.messagebox.showwarning") as warn,
        ):
            dlg_cls.return_value.result_name = GROUP_PLACEHOLDER
            dlg_cls.return_value.result_universal = False
            center.save_group()
        self.assertEqual([], center.groups)
        warn.assert_called_once()


class TaskCenterClientMismatchPromptTests(unittest.TestCase):
    """「开始顺序执行」前客户端不一致提示（_confirm_client_mismatch）。"""

    def _center(self, client_id="qianlong"):
        center = TaskCenter.__new__(TaskCenter)
        center.controller = SimpleNamespace(client_id=client_id)
        center.tasks = []
        return center

    def test_mismatch_returns_false_when_user_cancels(self):
        center = self._center("zhongtai")
        center.tasks = [
            {"params": {"client_id": "qianlong"}, "status": ""},
        ]
        with patch("gui.task_center.messagebox.askyesno", return_value=False) as ask:
            result = center._confirm_client_mismatch()
        self.assertFalse(result)
        ask.assert_called_once()

    def test_mismatch_returns_true_when_user_confirms(self):
        center = self._center("zhongtai")
        center.tasks = [
            {"params": {"client_id": "qianlong"}, "status": ""},
        ]
        with patch("gui.task_center.messagebox.askyesno", return_value=True) as ask:
            result = center._confirm_client_mismatch()
        self.assertTrue(result)
        ask.assert_called_once()

    def test_mixed_client_queue_prompts_once(self):
        center = self._center("qianlong")
        center.tasks = [
            {"params": {"client_id": "qianlong"}, "status": ""},
            {"params": {"client_id": "zhongtai"}, "status": ""},
        ]
        with patch("gui.task_center.messagebox.askyesno", return_value=True) as ask:
            center._confirm_client_mismatch()
        ask.assert_called_once()

    def test_all_tasks_match_current_client_no_prompt(self):
        center = self._center("qianlong")
        center.tasks = [
            {"params": {"client_id": "qianlong"}, "status": ""},
            {"params": {"client_id": "qianlong"}, "status": ""},
        ]
        with patch("gui.task_center.messagebox.askyesno") as ask:
            result = center._confirm_client_mismatch()
        self.assertTrue(result)
        ask.assert_not_called()

    def test_no_client_id_snapshot_no_prompt(self):
        center = self._center("qianlong")
        center.tasks = [{"params": {}, "status": ""}]
        with patch("gui.task_center.messagebox.askyesno") as ask:
            result = center._confirm_client_mismatch()
        self.assertTrue(result)
        ask.assert_not_called()


class TaskCenterRenameGroupTests(unittest.TestCase):
    """编队重命名：保留 client_id/tasks，并同步定时任务引用。"""

    def _center(self, client_id="qianlong"):
        center = TaskCenter.__new__(TaskCenter)
        center.is_running = False
        center.parent = None
        center.controller = SimpleNamespace(client_id=client_id)
        center.gui = SimpleNamespace(_log=lambda *a, **k: None)
        center.groups = [
            {"name": "钱龙开盘", "client_id": "qianlong", "tasks": [{"a": 1}]},
        ]
        center._current_group = "钱龙开盘"
        center._dirty = True
        center._refresh_group_combo = Mock()
        center._refresh_group_controls = Mock()
        center._save_groups = Mock()
        center.group_combo = SimpleNamespace(get=lambda: "钱龙开盘")
        return center

    def test_rename_keeps_client_and_tasks(self):
        center = self._center()
        with (
            patch("gui.task_center.ask_string", return_value="钱龙晨检"),
            patch.object(center, "_rename_group_in_scheduler", return_value=0) as rn,
        ):
            center.rename_group()
        self.assertEqual("钱龙晨检", center.groups[0]["name"])
        self.assertEqual("qianlong", center.groups[0]["client_id"])
        self.assertEqual([{"a": 1}], center.groups[0]["tasks"])
        center._save_groups.assert_called_once_with()
        center._refresh_group_combo.assert_called_once_with(select="钱龙晨检")
        self.assertEqual("钱龙晨检", center._current_group)
        self.assertFalse(center._dirty)
        rn.assert_called_once_with("钱龙开盘", "钱龙晨检")

    def test_rename_duplicate_name_rejected(self):
        center = self._center()
        center.groups.append({"name": "钱龙晨检", "client_id": "qianlong", "tasks": []})
        with (
            patch("gui.task_center.ask_string", return_value="钱龙晨检"),
            patch("gui.task_center.messagebox.showwarning") as warn,
        ):
            center.rename_group()
        self.assertEqual("钱龙开盘", center.groups[0]["name"])
        warn.assert_called_once()

    def test_rename_same_name_noop(self):
        center = self._center()
        with patch("gui.task_center.ask_string", return_value="钱龙开盘"):
            center.rename_group()
        center._save_groups.assert_not_called()
        self.assertEqual("钱龙开盘", center.groups[0]["name"])

    def test_rename_syncs_scheduler_references(self):
        center = self._center()
        scheduler = Mock()
        scheduler.tasks = [
            {"id": 1, "target_type": "group", "group_name": "钱龙开盘"},
            {"id": 2, "target_type": "script", "group_name": "钱龙开盘"},
            {"id": 3, "target_type": "group", "group_name": "别的"},
        ]
        center.controller.scheduler = scheduler
        updated = center._rename_group_in_scheduler("钱龙开盘", "钱龙晨检")
        self.assertEqual(1, updated)
        scheduler.update_task.assert_called_once_with(1, {"group_name": "钱龙晨检"})

    def test_rename_sync_without_scheduler_returns_0(self):
        center = self._center()
        center.controller = SimpleNamespace(client_id="qianlong")
        self.assertEqual(0, center._rename_group_in_scheduler("a", "b"))


class AddSchedDialogGroupClientFilterTests(unittest.TestCase):
    """定时任务对话框的编队按客户端过滤。"""

    def _dialog(self, client_id="qianlong", show_all=False):
        dlg = AddSchedDialog.__new__(AddSchedDialog)
        dlg.client_id = client_id
        dlg.group_show_all = _Value(show_all)
        dlg.groups = [
            {"name": "通用编队", "client_id": ""},
            {"name": "钱龙开盘", "client_id": "qianlong"},
            {"name": "中泰开盘", "client_id": "zhongtai"},
        ]
        return dlg

    def test_visible_groups_filters_by_current_client(self):
        dlg = self._dialog("qianlong")
        names = [g["name"] for g in dlg._visible_groups()]
        self.assertEqual(["通用编队", "钱龙开盘"], names)

    def test_visible_groups_show_all(self):
        dlg = self._dialog("qianlong", show_all=True)
        self.assertEqual(3, len(dlg._visible_groups()))


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
