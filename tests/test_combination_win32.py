import importlib.util
from pathlib import Path
import ctypes
from ctypes import wintypes
import os
import struct
import sys
import types
import unittest
from unittest.mock import Mock, call, patch


PROJECT_ROOT = Path(__file__).resolve().parents[1]
MODULE_PATH = PROJECT_ROOT / "core" / "combination_order.py"
SPEC = importlib.util.spec_from_file_location("combination_win32", MODULE_PATH)
combination = importlib.util.module_from_spec(SPEC)
assert SPEC.loader is not None
SPEC.loader.exec_module(combination)


class HybridListSelectionTests(unittest.TestCase):
    def test_lvitem_state_offsets_are_shared_by_32_and_64_bit(self):
        for bits in (32, 64):
            packed = combination._pack_lvitem_state(
                bits, 7, combination.LVIS_SELECTED, combination.LVIS_SELECTED
            )
            self.assertEqual(len(packed), 24)
            self.assertEqual(struct.unpack_from("<i", packed, 4)[0], 7)
            self.assertEqual(
                struct.unpack_from("<I", packed, 12)[0],
                combination.LVIS_SELECTED,
            )
            self.assertEqual(
                struct.unpack_from("<I", packed, 16)[0],
                combination.LVIS_SELECTED,
            )

    def test_uses_native_selection_when_remote_memory_is_available(self):
        with (
            patch.object(combination, "_list_select_first_native") as native,
            patch.object(combination, "_list_select_first_keyboard") as keyboard,
            patch.object(combination, "_list_select_first_uia") as uia,
        ):
            result = combination.list_select_first(100, 200)

        self.assertTrue(result["ok"])
        self.assertEqual(result["method"], "Win32远程内存")
        native.assert_called_once_with(200)
        keyboard.assert_not_called()
        uia.assert_not_called()

    def test_falls_back_to_keyboard_when_remote_memory_is_denied(self):
        with (
            patch.object(
                combination,
                "_list_select_first_native",
                side_effect=PermissionError(5, "拒绝访问"),
            ),
            patch.object(combination, "_list_select_first_keyboard") as keyboard,
            patch.object(combination, "_list_select_first_uia") as uia,
        ):
            result = combination.list_select_first(100, 200)

        self.assertTrue(result["ok"])
        self.assertEqual(result["method"], "Win32键盘消息")
        keyboard.assert_called_once_with(100, 200)
        uia.assert_not_called()

    def test_falls_back_to_uia_when_keyboard_messages_fail(self):
        with (
            patch.object(
                combination,
                "_list_select_first_native",
                side_effect=PermissionError(5, "拒绝访问"),
            ),
            patch.object(
                combination,
                "_list_select_first_keyboard",
                side_effect=RuntimeError("键盘消息未选中"),
            ),
            patch.object(combination, "_list_select_first_uia") as uia,
        ):
            result = combination.list_select_first(100, 200)

        self.assertTrue(result["ok"])
        self.assertEqual(result["method"], "UIA")
        uia.assert_called_once_with(100, 200)

    def test_fails_closed_when_all_selection_methods_fail(self):
        with (
            patch.object(
                combination,
                "_list_select_first_native",
                side_effect=PermissionError(5, "拒绝访问"),
            ),
            patch.object(
                combination,
                "_list_select_first_keyboard",
                side_effect=RuntimeError("键盘消息未选中"),
            ),
            patch.object(
                combination,
                "_list_select_first_uia",
                side_effect=RuntimeError("UIA未暴露列表行"),
            ),
        ):
            result = combination.list_select_first(100, 200)

        self.assertFalse(result["ok"])
        self.assertIn("拒绝访问", result["error"])
        self.assertIn("键盘消息未选中", result["error"])
        self.assertIn("UIA未暴露列表行", result["error"])


class StableListTests(unittest.TestCase):
    def test_requires_three_equal_row_counts(self):
        counts = iter([2, 3, 3, 3])
        with (
            patch.object(combination, "find_visible_child", return_value=200),
            patch.object(
                combination, "list_count", side_effect=lambda _: next(counts)
            ),
            patch.object(combination.time, "sleep"),
        ):
            self.assertTrue(
                combination.check_list_has_data(100, wait=0, timeout=1)
            )

    def test_split_accepts_one_stable_row(self):
        spec = importlib.util.spec_from_file_location(
            "combination_win32_split_one_row", MODULE_PATH
        )
        split = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        with patch.dict(os.environ, {"GUI_COMBINATION_ACTION": "split"}):
            spec.loader.exec_module(split)
        with (
            patch.object(split, "find_visible_child", return_value=200),
            patch.object(split, "list_count", return_value=1),
            patch.object(split.time, "sleep"),
        ):
            self.assertEqual(split.MIN_DATA_ROWS, 1)
            self.assertTrue(split.check_list_has_data(100, wait=0, timeout=1))


class NativeListIntegrationTests(unittest.TestCase):
    def test_real_sys_listview_remote_selection_round_trip(self):
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        comctl32 = ctypes.WinDLL("comctl32", use_last_error=True)

        class InitCommonControlsEx(ctypes.Structure):
            _fields_ = [("dwSize", wintypes.DWORD), ("dwICC", wintypes.DWORD)]

        class LvItemW(ctypes.Structure):
            _fields_ = [
                ("mask", wintypes.UINT),
                ("iItem", ctypes.c_int),
                ("iSubItem", ctypes.c_int),
                ("state", wintypes.UINT),
                ("stateMask", wintypes.UINT),
                ("pszText", wintypes.LPWSTR),
                ("cchTextMax", ctypes.c_int),
                ("iImage", ctypes.c_int),
                ("lParam", ctypes.c_ssize_t),
            ]

        init = InitCommonControlsEx(ctypes.sizeof(InitCommonControlsEx), 0x0001)
        self.assertTrue(comctl32.InitCommonControlsEx(ctypes.byref(init)))
        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, ctypes.c_void_p, wintypes.HINSTANCE, ctypes.c_void_p,
        ]
        user32.CreateWindowExW.restype = wintypes.HWND
        user32.SendMessageW.argtypes = [
            wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t
        ]
        user32.SendMessageW.restype = ctypes.c_ssize_t
        instance = kernel32.GetModuleHandleW(None)
        parent = user32.CreateWindowExW(
            0, "STATIC", "NativeListTest", 0x00CF0000,
            0, 0, 400, 300, None, None, instance, None,
        )
        list_hwnd = user32.CreateWindowExW(
            0, "SysListView32", "", 0x50000001,
            0, 0, 300, 200, parent, ctypes.c_void_p(1229), instance, None,
        )
        self.assertTrue(parent)
        self.assertTrue(list_hwnd)
        text_buffer = ctypes.create_unicode_buffer("first row")
        item = LvItemW()
        item.mask = 0x0001  # LVIF_TEXT
        item.iItem = 0
        item.pszText = ctypes.cast(text_buffer, wintypes.LPWSTR)
        try:
            inserted = user32.SendMessageW(
                list_hwnd, 0x104D, 0, ctypes.addressof(item)
            )  # LVM_INSERTITEMW
            self.assertEqual(inserted, 0)
            self.assertEqual(combination.list_count(list_hwnd), 1)
            combination._list_select_first_native(list_hwnd)
            self.assertEqual(combination._selected_list_index(list_hwnd), 0)
        finally:
            user32.DestroyWindow(parent)


class WorkflowModeTests(unittest.TestCase):
    def _load_candidate(self, action):
        name = f"combination_win32_{action}"
        spec = importlib.util.spec_from_file_location(name, MODULE_PATH)
        module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        with patch.dict(
            os.environ,
            {"GUI_COMBINATION_ACTION": action},
        ):
            spec.loader.exec_module(module)
        return module

    def test_split_mode_uses_split_panel_without_strategy_combo(self):
        split = self._load_candidate("split")
        self.assertFalse(split.IS_COMBINE)
        self.assertEqual(split.ACTION_NAME, "拆分")
        self.assertEqual(split.DIALOG_TITLE, "拆分申报")
        self.assertEqual(split.PANEL_PATH, r"\组合申报\拆分申报")

    def test_formal_entry_scripts_select_fixed_modes(self):
        entries = (
            ("2.组合申报_全自动.py", "combine"),
            ("2.拆分申报_全自动.py", "split"),
        )
        for filename, expected_action in entries:
            path = PROJECT_ROOT / "组合申报" / filename
            spec = importlib.util.spec_from_file_location(
                f"entry_{expected_action}", path
            )
            module = importlib.util.module_from_spec(spec)
            assert spec.loader is not None
            fake_driver = types.ModuleType("core.combination_order")
            fake_driver.main = Mock()
            with (
                patch.dict(os.environ, {}, clear=False),
                patch.dict(
                    sys.modules, {"core.combination_order": fake_driver}
                ),
            ):
                spec.loader.exec_module(module)
                self.assertEqual(
                    os.environ["GUI_COMBINATION_ACTION"], expected_action
                )
                self.assertIs(module.main, fake_driver.main)
                fake_driver.main.assert_not_called()


class GuiTaskEnvironmentTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        task_path = PROJECT_ROOT / "GUI自动化工具2" / "engine" / "task.py"
        spec = importlib.util.spec_from_file_location("gui_task", task_path)
        cls.task_module = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        spec.loader.exec_module(cls.task_module)

    def test_gui_tasks_do_not_set_dry_run(self):
        task = self.task_module.Task(
            {"name": "组合申报", "path": "unused.py"}, "组合申报"
        )
        self.assertNotIn("GUI_DRY_RUN", task.build_env(str(PROJECT_ROOT)))


class DialogStateMachineTests(unittest.TestCase):
    def test_owner_drawn_idok_is_accepted_as_affirmative(self):
        def get_item(_dialog, ctrl_id):
            return 901 if ctrl_id == 1 else 0

        with (
            patch.object(combination.win32gui, "GetDlgItem", side_effect=get_item),
            patch.object(combination, "_usable_button", return_value=True),
        ):
            self.assertEqual(combination._find_dialog_button(100), 901)

    def test_transition_ignores_dialogs_that_existed_before_click(self):
        with (
            patch.object(combination.win32gui, "IsWindow", return_value=True),
            patch.object(
                combination.win32gui, "IsWindowVisible", return_value=True
            ),
            patch.object(
                combination,
                "_relevant_dialogs",
                side_effect=[[10, 20], [10, 20, 30]],
            ),
            patch.object(combination.time, "sleep"),
        ):
            self.assertTrue(
                combination._wait_dialog_transition(
                    100, 10, known_dialogs={10, 20}, timeout=1
                )
            )

    def test_confirm_retries_until_dialog_state_changes(self):
        with (
            patch.object(combination.win32gui, "IsWindow", return_value=True),
            patch.object(
                combination.win32gui, "IsWindowVisible", return_value=True
            ),
            patch.object(combination.win32gui, "GetWindowText", return_value="警告"),
            patch.object(combination, "_wait_for_dialog_button", return_value=901),
            patch.object(combination, "_relevant_dialogs", return_value=[10]),
            patch.object(
                combination,
                "_wait_dialog_transition",
                side_effect=[False, True],
            ),
            patch.object(combination.win32gui, "PostMessage") as post,
        ):
            self.assertTrue(combination.confirm_dialog(100, 10, attempts=3))
        self.assertEqual(post.call_count, 2)

    def test_abort_closes_nested_dialogs_until_none_remain(self):
        with (
            patch.object(
                combination,
                "_relevant_dialogs",
                side_effect=[[20], [10], []],
            ),
            patch.object(combination, "has_control", return_value=False),
            patch.object(combination.win32gui, "GetWindowText", return_value="警告"),
            patch.object(combination, "cancel_dialog", return_value=True) as cancel,
            patch.object(combination.time, "sleep"),
        ):
            self.assertTrue(combination.abort_transaction_dialogs(100))
        self.assertEqual(cancel.call_count, 2)

    def test_followup_warning_is_confirmed_then_waits_for_quiet(self):
        with (
            patch.object(
                combination,
                "_relevant_dialogs",
                side_effect=[[20], [], []],
            ),
            patch.object(combination, "has_control", return_value=False),
            patch.object(combination, "confirm_dialog", return_value=True) as confirm,
            patch.object(
                combination.time,
                "time",
                side_effect=[0, 0, 0, 0, 0.5, 1.3],
            ),
            patch.object(combination.time, "sleep"),
        ):
            self.assertTrue(
                combination.handle_dialogs(100, timeout=8, quiet_period=1.2)
            )
        confirm.assert_called_once_with(100, 20, attempts=3)

    def test_return_to_quantity_dialog_is_not_reported_as_success(self):
        with (
            patch.object(combination, "_relevant_dialogs", return_value=[10]),
            patch.object(combination, "has_control", return_value=True),
            patch.object(combination.time, "time", side_effect=[0, 0]),
        ):
            self.assertFalse(combination.handle_dialogs(100))


class FullTraversalTests(unittest.TestCase):
    def _run_main_with_mocks(self, module):
        combo_handles = {9059: 501, 9040: 502}
        module._combo_sources.clear()
        module._combo_sources.update({501: "测试", 502: "测试"})
        with (
            patch.object(module, "countdown"),
            patch.object(module, "find_main_window", return_value=100),
            patch.object(module, "close_leftover_qty_dialogs"),
            patch.object(module, "activate_main_window", return_value=object()),
            patch.object(module, "switch_panel"),
            patch.object(
                module,
                "find_visible_child",
                side_effect=lambda _hwnd, cid: combo_handles.get(cid),
            ),
            patch.object(module, "build_combo_map", return_value={"测试": 0}),
            patch.object(module, "select_exchange", return_value=True),
            patch.object(module, "process_item", return_value=False) as process,
            patch.object(module.time, "sleep"),
        ):
            module.main()
        return process

    def test_combination_traverses_all_twelve_exchange_strategy_pairs(self):
        process = self._run_main_with_mocks(combination)
        self.assertEqual(process.call_count, 12)
        expected = [
            call(100, exchange, strategy)
            for exchange in combination.EXCHANGES
            for strategy in combination.STRATEGIES
        ]
        self.assertEqual(process.call_args_list, expected)

    def test_split_traverses_both_exchanges(self):
        spec = importlib.util.spec_from_file_location(
            "combination_win32_split_full", MODULE_PATH
        )
        split = importlib.util.module_from_spec(spec)
        assert spec.loader is not None
        with patch.dict(os.environ, {"GUI_COMBINATION_ACTION": "split"}):
            spec.loader.exec_module(split)
        process = self._run_main_with_mocks(split)
        self.assertEqual(
            process.call_args_list,
            [call(100, exchange, None) for exchange in split.EXCHANGES],
        )


if __name__ == "__main__":
    unittest.main()
