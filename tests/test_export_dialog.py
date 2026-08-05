# -*- coding: utf-8 -*-
"""导出流程"提示"弹窗三态处理测试。

覆盖：遇到"提示"弹窗（数据未就绪）时 handle_* 返回 "no_data"，
runner 不再重试点击"输出"按钮；其余成功/失败路径行为不变。
"""

import unittest
from unittest.mock import patch

from core import export_dialog
from core import runner
from core import save_as_dialog


def _base_cfg(**overrides):
    cfg = {
        "panel_path": r"\查询\策略委托",
        "name": "策略委托",
        "default_txt": "out.txt",
        "default_xls": "out.xlsx",
        "export_format": "xls",
        "auto_open": True,
        "countdown_sec": 0,
        "settle_delay": 0,
        "window_key": "测试窗口",
        "use_title": False,
        "output_path": "out.xlsx",
        "client_id": "qianlong",
    }
    cfg.update(overrides)
    return cfg


class HandleDialogTriStateTests(unittest.TestCase):
    """handle_export_dialog / handle_save_as_dialog 的三态返回。"""

    def test_export_dialog_no_data_returns_no_data_sentinel(self):
        with patch.object(export_dialog, "_find_dialog", return_value=(None, True)):
            result = export_dialog.handle_export_dialog(timeout=0.1)
        self.assertEqual(result, "no_data")

    def test_export_dialog_timeout_returns_false(self):
        with patch.object(export_dialog, "_find_dialog", return_value=(None, False)):
            result = export_dialog.handle_export_dialog(timeout=0.1)
        self.assertFalse(result)

    def test_saveas_dialog_no_data_returns_no_data_sentinel(self):
        with patch.object(
            save_as_dialog, "_find_saveas_window", return_value=(None, True)
        ):
            result = save_as_dialog.handle_save_as_dialog(
                save_dir=".", filename="x.xlsx", timeout=0.1
            )
        self.assertEqual(result, "no_data")

    def test_saveas_dialog_timeout_returns_false(self):
        with patch.object(
            save_as_dialog, "_find_saveas_window", return_value=(None, False)
        ):
            result = save_as_dialog.handle_save_as_dialog(
                save_dir=".", filename="x.xlsx", timeout=0.1
            )
        self.assertFalse(result)


class RunnerNoDataRetryTests(unittest.TestCase):
    """runner 遇到 no_data 时不再重试点击"输出"按钮。"""

    def _run_export(self, result, retry_click=True):
        cfg = _base_cfg()
        with (
            patch.object(runner, "parse_env_config", return_value=cfg),
            patch.object(runner, "countdown"),
            patch.object(runner, "find_window", return_value=123),
            patch.object(runner, "activate_window", return_value=object()),
            patch.object(runner, "switch_panel"),
            patch.object(runner, "click_output_button",
                         return_value=retry_click) as mock_click,
            patch.object(runner, "handle_export_dialog",
                         return_value=result) as mock_handle,
        ):
            runner.run_export_dialog(cfg)
        return mock_click, mock_handle

    def test_export_no_data_does_not_retry_click(self):
        mock_click, mock_handle = self._run_export("no_data")
        mock_click.assert_called_once()
        mock_handle.assert_called_once()

    def test_export_success_does_not_retry_click(self):
        mock_click, mock_handle = self._run_export(True)
        mock_click.assert_called_once()
        mock_handle.assert_called_once()

    def test_export_timeout_retries_once(self):
        mock_click, mock_handle = self._run_export(False)
        self.assertEqual(mock_click.call_count, 2)
        self.assertEqual(mock_handle.call_count, 2)

    def _run_save_as(self, result):
        cfg = _base_cfg(output_path=r"outdir\x.xlsx")
        with (
            patch.object(runner, "parse_env_config", return_value=cfg),
            patch.object(runner, "countdown"),
            patch.object(runner, "find_window", return_value=123),
            patch.object(runner, "activate_window", return_value=object()),
            patch.object(runner, "switch_panel"),
            patch.object(runner, "click_output_button",
                         return_value=True) as mock_click,
            patch.object(runner, "handle_save_as_dialog",
                         return_value=result) as mock_handle,
        ):
            runner.run_save_as(cfg)
        return mock_click, mock_handle

    def test_saveas_no_data_does_not_retry_click(self):
        mock_click, mock_handle = self._run_save_as("no_data")
        mock_click.assert_called_once()
        mock_handle.assert_called_once()

    def test_saveas_success_does_not_retry_click(self):
        mock_click, mock_handle = self._run_save_as(True)
        mock_click.assert_called_once()

    def test_saveas_timeout_retries_once(self):
        mock_click, mock_handle = self._run_save_as(False)
        self.assertEqual(mock_click.call_count, 2)
        self.assertEqual(mock_handle.call_count, 2)


if __name__ == "__main__":
    unittest.main()