# -*- coding: utf-8 -*-

from pathlib import Path
import sys
import unittest
from unittest.mock import Mock, patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
GUI_ROOT = PROJECT_ROOT / "GUI自动化工具2"
if str(GUI_ROOT) not in sys.path:
    sys.path.insert(0, str(GUI_ROOT))

from gui.main_window import (
    AutomationGUI,
    RightTabsDialog,
    normalize_right_tabs,
    RIGHT_PANEL_DEFS,
    DEFAULT_RIGHT_TABS,
)

ALL_KEYS = [key for key, _name, _visible in RIGHT_PANEL_DEFS]


class NormalizeRightTabsTests(unittest.TestCase):
    def test_missing_config_falls_back_to_default(self):
        self.assertEqual(normalize_right_tabs(None), DEFAULT_RIGHT_TABS)

    def test_invalid_type_falls_back_to_default(self):
        self.assertEqual(normalize_right_tabs("oops"), DEFAULT_RIGHT_TABS)

    def test_unknown_keys_are_filtered(self):
        result = normalize_right_tabs(["不存在的页", "运行日志"])
        self.assertEqual(
            result,
            ["运行日志"] + [k for k in DEFAULT_RIGHT_TABS if k != "运行日志"],
        )

    def test_duplicated_keys_dedupe_and_keep_first(self):
        result = normalize_right_tabs(["任务中心", "运行日志", "任务中心"])
        self.assertEqual(result[0], "任务中心")
        self.assertEqual(result.count("任务中心"), 1)

    def test_new_panels_are_appended(self):
        # 模拟旧配置缺了后面新增的面板：缺失面板自动追加到末尾
        partial = ["运行日志", "任务历史"]
        result = normalize_right_tabs(partial)
        self.assertEqual(result[:2], partial)
        self.assertEqual(result[2:], [k for k in DEFAULT_RIGHT_TABS if k not in partial])

    def test_empty_list_rebuilds_full_default(self):
        self.assertEqual(normalize_right_tabs([]), DEFAULT_RIGHT_TABS)


class ApplyRightTabsTests(unittest.TestCase):
    def _make_gui(self, vis_order):
        gui = AutomationGUI.__new__(AutomationGUI)
        gui._right_tabs = list(vis_order)
        gui._right_panels = {k: {"frame": f"frame-{k}", "name": k} for k in ALL_KEYS}
        notebook = Mock()
        notebook.tabs.return_value = [f"frame-{k}" for k in reversed(vis_order)]
        gui.right_notebook = notebook
        return gui

    def test_apply_mounts_visible_tabs_in_configured_order(self):
        gui = self._make_gui(["报告中心", "运行日志", "任务历史"])
        AutomationGUI._apply_right_tabs(gui)
        gui.right_notebook.forget.assert_called()
        mounted = [c.args[0] for c in gui.right_notebook.add.call_args_list]
        self.assertEqual(mounted, ["frame-报告中心", "frame-运行日志", "frame-任务历史"])

    def test_apply_hides_unlisted_tabs(self):
        gui = self._make_gui(["运行日志"])
        AutomationGUI._apply_right_tabs(gui)
        mounted = [c.args[0] for c in gui.right_notebook.add.call_args_list]
        self.assertEqual(mounted, ["frame-运行日志"])

    def test_apply_without_notebook_is_noop(self):
        gui = AutomationGUI.__new__(AutomationGUI)
        gui._right_tabs = ["运行日志"]
        gui._right_panels = {k: {"frame": f"frame-{k}", "name": k} for k in ALL_KEYS}
        AutomationGUI._apply_right_tabs(gui)  # 不应抛异常


class RightTabsDialogLogicTests(unittest.TestCase):
    def _make_dialog(self, vis_order):
        dlg = RightTabsDialog.__new__(RightTabsDialog)
        dlg.vis_order = list(vis_order)
        dlg._hidden = [k for k in ALL_KEYS if k not in dlg.vis_order]
        dlg._name_map = {key: name for key, name, _v in RIGHT_PANEL_DEFS}
        dlg._entries = dlg.vis_order + dlg._hidden
        dlg.listbox = Mock()
        dlg.listbox.curselection.return_value = ()
        return dlg

    def test_toggle_checked_moves_key_to_hidden(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心"])
        # entries: 运行日志(0) 任务历史(1) 任务中心(2) 定时任务(3) 结果比对(4) 报告中心(5)
        dlg.listbox.curselection.return_value = (1,)  # 取消勾选 任务历史
        dlg._toggle_selected()
        self.assertEqual(dlg.vis_order, ["运行日志", "任务中心"])
        self.assertEqual(
            dlg._hidden,
            ["定时任务", "结果比对", "报告中心", "任务历史"],
        )

    def test_toggle_unchecked_appends_to_vis_order(self):
        dlg = self._make_dialog(["运行日志"])
        # entries: 运行日志(0) 任务历史(1) 任务中心(2) 定时任务(3) 结果比对(4) 报告中心(5)
        dlg.listbox.curselection.return_value = (3,)  # 勾选 定时任务
        dlg._toggle_selected()
        self.assertEqual(dlg.vis_order, ["运行日志", "定时任务"])
        self.assertNotIn("定时任务", dlg._hidden)

    def test_toggle_empty_selection_is_noop(self):
        dlg = self._make_dialog(["运行日志", "任务历史"])
        dlg._toggle_selected()
        self.assertEqual(dlg.vis_order, ["运行日志", "任务历史"])

    def test_move_up_single(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心", "定时任务"])
        dlg.listbox.curselection.return_value = (2,)  # 任务中心
        dlg._move(-1)
        self.assertEqual(dlg.vis_order, ["运行日志", "任务中心", "任务历史", "定时任务"])

    def test_move_down_single(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心", "定时任务"])
        dlg.listbox.curselection.return_value = (1,)  # 任务历史
        dlg._move(1)
        self.assertEqual(dlg.vis_order, ["运行日志", "任务中心", "任务历史", "定时任务"])

    def test_move_ignores_hidden_items(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心"])
        # entries: 显示区 3 项(0-2)，隐藏区 3 项(3-5)；选中隐藏区「定时任务」
        dlg.listbox.curselection.return_value = (3,)
        dlg._move(-1)
        self.assertEqual(dlg.vis_order, ["运行日志", "任务历史", "任务中心"])

    def test_move_blocked_at_edges(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心"])
        dlg.listbox.curselection.return_value = (0,)
        dlg._move(-1)
        self.assertEqual(dlg.vis_order, ["运行日志", "任务历史", "任务中心"])
        dlg.listbox.curselection.return_value = (2,)
        dlg._move(1)
        self.assertEqual(dlg.vis_order, ["运行日志", "任务历史", "任务中心"])

    def test_move_multi_block_down(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心", "定时任务"])
        dlg.listbox.curselection.return_value = (0, 1)  # 运行日志+任务历史
        dlg._move(1)
        self.assertEqual(
            dlg.vis_order,
            ["任务中心", "运行日志", "任务历史", "定时任务"],
        )

    def test_move_multi_block_up(self):
        dlg = self._make_dialog(["运行日志", "任务历史", "任务中心", "定时任务"])
        dlg.listbox.curselection.return_value = (2, 3)  # 任务中心+定时任务
        dlg._move(-1)
        self.assertEqual(
            dlg.vis_order,
            ["运行日志", "任务中心", "定时任务", "任务历史"],
        )

    def test_show_all_moves_hidden_to_tail(self):
        dlg = self._make_dialog(["运行日志", "任务历史"])
        dlg._show_all()
        self.assertEqual(
            dlg.vis_order,
            ["运行日志", "任务历史", "任务中心", "定时任务", "结果比对", "报告中心"],
        )
        self.assertEqual(dlg._hidden, [])

    def test_hide_all_moves_all_to_hidden(self):
        dlg = self._make_dialog(["运行日志", "任务历史"])
        dlg._hide_all()
        self.assertEqual(dlg.vis_order, [])
        self.assertEqual(set(dlg._hidden), set(ALL_KEYS))

    def test_reset_default(self):
        dlg = self._make_dialog(["任务中心"])
        dlg._reset_default()
        self.assertEqual(dlg.vis_order, DEFAULT_RIGHT_TABS)
        self.assertEqual(dlg._hidden, [])

    def test_validate_rejects_all_hidden(self):
        dlg = self._make_dialog([])
        with patch("gui.main_window.messagebox.showwarning"):
            self.assertFalse(dlg.validate())
        dlg = self._make_dialog(["运行日志"])
        self.assertTrue(dlg.validate())


class ShowReportCenterTests(unittest.TestCase):
    def _make_gui(self, vis_order):
        gui = AutomationGUI.__new__(AutomationGUI)
        gui._right_tabs = list(vis_order)
        gui._report_frame = "frame-报告中心"
        gui.right_notebook = Mock()
        gui._log = Mock()
        return gui

    def test_select_report_when_visible(self):
        gui = self._make_gui(["报告中心", "运行日志"])
        AutomationGUI.show_report_center(gui)
        gui.right_notebook.select.assert_called_once_with("frame-报告中心")
        gui._log.assert_not_called()

    def test_skip_select_when_hidden(self):
        gui = self._make_gui(["运行日志"])
        AutomationGUI.show_report_center(gui)
        gui.right_notebook.select.assert_not_called()
        gui._log.assert_called_once()


if __name__ == "__main__":
    unittest.main()
