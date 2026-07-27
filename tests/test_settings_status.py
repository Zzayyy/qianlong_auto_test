# -*- coding: utf-8 -*-

import ast
from pathlib import Path
import unittest

from core.settings import (
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_DIFFERENCE,
    STATUS_DISABLED,
    STATUS_EXECUTION_FAILED,
    STATUS_NOT_APPLICABLE,
    STATUS_PASS,
    STATUS_UNVERIFIED,
    count_statuses,
    is_difference_status,
    normalize_status,
)


class SettingsStatusTests(unittest.TestCase):
    def test_canonical_statuses_remain_unchanged(self):
        statuses = (
            STATUS_PASS,
            STATUS_DIFFERENCE,
            STATUS_DISABLED,
            STATUS_UNVERIFIED,
            STATUS_ADDED,
            STATUS_CONFLICT,
            STATUS_EXECUTION_FAILED,
            STATUS_NOT_APPLICABLE,
        )
        self.assertEqual([normalize_status(value) for value in statuses], list(statuses))

    def test_legacy_report_labels_are_normalized(self):
        self.assertEqual(normalize_status("✓"), STATUS_PASS)
        self.assertEqual(normalize_status("✗ 差异"), STATUS_DIFFERENCE)
        self.assertEqual(normalize_status("○ 未启用"), STATUS_DISABLED)

    def test_english_and_module_statuses_are_normalized(self):
        self.assertEqual(normalize_status("unverified"), STATUS_UNVERIFIED)
        self.assertEqual(normalize_status("FAILED"), STATUS_EXECUTION_FAILED)
        self.assertEqual(normalize_status("skipped"), STATUS_NOT_APPLICABLE)

    def test_unknown_status_is_rejected(self):
        with self.assertRaises(ValueError):
            normalize_status("待确认的新状态")

    def test_only_business_differences_enter_difference_report(self):
        for status in (STATUS_DIFFERENCE, STATUS_ADDED, STATUS_CONFLICT):
            self.assertTrue(is_difference_status(status))
        for status in (
            STATUS_PASS,
            STATUS_DISABLED,
            STATUS_UNVERIFIED,
            STATUS_EXECUTION_FAILED,
            STATUS_NOT_APPLICABLE,
        ):
            self.assertFalse(is_difference_status(status))

    def test_counts_use_the_canonical_status_field(self):
        rows = [
            {"状态": "✓"},
            {"状态": STATUS_PASS},
            {"状态": "✗ 差异"},
        ]
        counts = count_statuses(rows)
        self.assertEqual(counts[STATUS_PASS], 2)
        self.assertEqual(counts[STATUS_DIFFERENCE], 1)

    def test_all_settings_scripts_use_the_shared_result_model(self):
        settings_dir = Path(__file__).resolve().parents[1] / "交易系统设置"
        scripts = sorted(settings_dir.glob("*.py"))
        legacy_window_helpers = {
            "open_settings_dialog",
            "_find_existing_settings_dlg",
            "_click_settings_button",
            "_click_button_by_auto_id",
            "_click_context_menu_item",
            "_try_keyboard_select",
            "switch_to_settings_panel",
            "_toggle_combobox",
        }
        self.assertEqual(len(scripts), 7)
        for script in scripts:
            with self.subTest(script=script.name):
                tree = ast.parse(script.read_text(encoding="utf-8-sig"))
                imports_shared_result = any(
                    isinstance(node, ast.ImportFrom)
                    and node.module == "core.settings"
                    for node in ast.walk(tree)
                )
                self.assertTrue(imports_shared_result)
                imports_shared_window = any(
                    isinstance(node, ast.ImportFrom)
                    and node.module == "core.settings_window"
                    for node in ast.walk(tree)
                )
                self.assertTrue(imports_shared_window)
                local_function_names = {
                    node.name
                    for node in tree.body
                    if isinstance(node, ast.FunctionDef)
                }
                self.assertFalse(local_function_names & legacy_window_helpers)
                local_result_classes = [
                    node
                    for node in ast.walk(tree)
                    if isinstance(node, ast.ClassDef)
                    and node.name == "SettingsTestResult"
                ]
                self.assertFalse(local_result_classes)
                legacy_keys = [
                    node.value
                    for node in ast.walk(tree)
                    if isinstance(node, ast.Constant) and node.value == "是否一致"
                ]
                self.assertFalse(legacy_keys)


if __name__ == "__main__":
    unittest.main()
