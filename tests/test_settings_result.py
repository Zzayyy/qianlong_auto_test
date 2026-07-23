# -*- coding: utf-8 -*-

import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from core.settings_result import SettingsTestResult
from core.settings_status import (
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_DIFFERENCE,
    STATUS_DISABLED,
    STATUS_PASS,
    STATUS_UNVERIFIED,
)


class SettingsResultTests(unittest.TestCase):
    def test_all_rows_share_one_schema_and_status_mapping(self):
        result = SettingsTestResult("测试设置")
        result.add_result("通过项", 1, 1)
        result.add_result("差异项", 2, 1, "值不一致")
        result.add_not_enabled("未启用项")
        result.add_unverified("未验证项", "期望", "无法读取")
        result.add_status("新增项", "新值", "不存在", STATUS_ADDED)
        result.add_status("冲突项", "A、B", "无冲突", STATUS_CONFLICT)

        expected_keys = {"名称", "期望值", "实际值", "状态", "说明"}
        self.assertTrue(all(set(row) == expected_keys for row in result.results))
        self.assertEqual(
            [row["状态"] for row in result.results],
            [
                STATUS_PASS,
                STATUS_DIFFERENCE,
                STATUS_DISABLED,
                STATUS_UNVERIFIED,
                STATUS_ADDED,
                STATUS_CONFLICT,
            ],
        )
        self.assertEqual(len(result.differences), 3)
        self.assertEqual(len(result.unverified), 1)
        self.assertEqual(result.not_enabled, 1)

    def test_optional_normalizer_controls_value_comparison(self):
        result = SettingsTestResult(
            "测试设置", normalizer=lambda value: str(value).strip().lower()
        )
        result.add_result("规范化项", " Value ", "value")
        self.assertEqual(result.results[0]["状态"], STATUS_PASS)

    def test_summary_uses_canonical_counts(self):
        result = SettingsTestResult("测试设置")
        result.add_result("通过项", True, True)
        result.add_result("差异项", False, True)
        result.add_unverified("未验证项", 1, "读取失败")
        result.add_observation("采集项", "值", "仅记录")

        summary = result.summary()
        self.assertEqual(summary["总项目数"], 3)
        self.assertEqual(summary[STATUS_PASS], 1)
        self.assertEqual(summary[STATUS_DIFFERENCE], 1)
        self.assertEqual(summary[STATUS_UNVERIFIED], 1)
        self.assertEqual(summary["差异合计"], 1)
        self.assertEqual(summary["采集项"], 1)

    def test_text_and_json_reports_are_written_together(self):
        result = SettingsTestResult("测试设置")
        result.add_result("通过项", "A", "A")
        result.add_result("差异项", "B", "A", "测试说明")
        result.add_observation("来源", "测试", "只读")

        with tempfile.TemporaryDirectory() as temp_dir:
            report_path = Path(temp_dir) / "测试设置测试报告.txt"
            with mock.patch.dict(
                os.environ,
                {
                    "GUI_CLIENT_ID": "test_client",
                    "GUI_SETTINGS_RUN_ID": "test_run",
                },
            ):
                json_path = Path(result.to_file(str(report_path)))

            self.assertTrue(report_path.exists())
            self.assertTrue(json_path.exists())
            text = report_path.read_text(encoding="utf-8")
            self.assertIn("[通过] 通过项", text)
            self.assertIn("[差异] 差异项", text)
            self.assertIn("未验证: 0", text)

            payload = json.loads(json_path.read_text(encoding="utf-8"))
            self.assertEqual(payload["schema_version"], 1)
            self.assertEqual(payload["run_id"], "test_run")
            self.assertEqual(payload["module"], "测试设置")
            self.assertEqual(payload["client_id"], "test_client")
            self.assertEqual(payload["execution_status"], STATUS_PASS)
            self.assertEqual(payload["summary"]["差异合计"], 1)
            self.assertEqual(payload["items"][1]["status"], STATUS_DIFFERENCE)
            self.assertEqual(payload["items"][1]["detail"], "测试说明")
            self.assertEqual(payload["observations"][0]["name"], "来源")


if __name__ == "__main__":
    unittest.main()
