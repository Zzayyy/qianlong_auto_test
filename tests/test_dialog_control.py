# -*- coding: utf-8 -*-
"""core/dialog_control 单元测试（仅 Windows + pywin32 环境可跑）。

覆盖 close_no_data_prompt：仅按弹窗标题「提示」识别（内容文本自绘不可枚举），
命中后点击"确定"按钮确认，找不到按钮则 PostMessage(WM_CLOSE) 兜底。

测试通过 mock EnumWindows 隔离，避免受本机其它"提示"标题窗口干扰。
"""

import ctypes
import time
import unittest
from ctypes import wintypes
from unittest.mock import patch

import win32con
import win32gui

from core import dialog_control


def _create_prompt_window(title: str, button_text: str = None):
    """创建一个顶层可见窗口，可选带一个 Button 子控件。

    Returns:
        (hwnd, button_hwnd or None)
    """
    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.CreateWindowExW.argtypes = [
        wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
        ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
        wintypes.HWND, ctypes.c_void_p, wintypes.HINSTANCE, ctypes.c_void_p,
    ]
    user32.CreateWindowExW.restype = wintypes.HWND
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.UpdateWindow.argtypes = [wintypes.HWND]
    user32.UpdateWindow.restype = wintypes.BOOL

    instance = ctypes.WinDLL("kernel32", use_last_error=True).GetModuleHandleW(None)
    hwnd = user32.CreateWindowExW(
        0, "#32770", title, 0x00CF0000,
        100, 100, 360, 160, None, None, instance, None,
    )
    if not hwnd:
        raise RuntimeError(f"创建测试窗口失败: title={title}")
    btn = None
    if button_text:
        btn = user32.CreateWindowExW(
            0, "BUTTON", button_text, 0x50000000,
            140, 100, 80, 28, wintypes.HWND(hwnd), None, instance, None,
        )
    user32.ShowWindow(hwnd, 5)  # SW_SHOW
    user32.UpdateWindow(hwnd)
    return hwnd, btn


def _destroy_window(hwnd):
    if hwnd and win32gui.IsWindow(hwnd):
        win32gui.DestroyWindow(hwnd)


def _mock_enum_windows(hwnds):
    """mock win32gui.EnumWindows，只对给定窗口列表执行回调（隔离本机环境）。"""
    def _fake_enum(callback, _):
        for h in hwnds:
            callback(h, None)
    return patch.object(
        dialog_control.win32gui, "EnumWindows", side_effect=_fake_enum
    )


class CloseNoDataPromptTests(unittest.TestCase):
    """close_no_data_prompt 行为：按标题「提示」识别并确认弹窗。"""

    def test_returns_false_when_no_prompt(self):
        """没有任何"提示"标题弹窗时返回 False，不影响执行速度。"""
        with _mock_enum_windows([]):
            result = dialog_control.close_no_data_prompt(settle_delay=0)
        self.assertFalse(result)

    def test_clicks_confirm_button_on_title_match(self):
        """标题为"提示"的弹窗：点击"确定"按钮确认并返回 True。"""
        hwnd, btn = _create_prompt_window("提示", "确定(Y)")
        try:
            with (
                _mock_enum_windows([hwnd]),
                patch.object(dialog_control, "click") as mock_click,
            ):
                result = dialog_control.close_no_data_prompt(settle_delay=0)
            self.assertTrue(result, "标题为'提示'的弹窗应被识别处理")
            mock_click.assert_called_once_with(btn, hwnd)
        finally:
            _destroy_window(hwnd)

    def test_falls_back_to_wm_close_when_no_button(self):
        """标题"提示"但找不到确定按钮时，兜底 PostMessage(WM_CLOSE)。"""
        hwnd, _btn = _create_prompt_window("提示", None)
        try:
            with (
                _mock_enum_windows([hwnd]),
                patch.object(dialog_control, "click") as mock_click,
                patch.object(
                    dialog_control.win32gui, "PostMessage",
                    wraps=dialog_control.win32gui.PostMessage,
                ) as mock_post,
            ):
                result = dialog_control.close_no_data_prompt(settle_delay=0)
            self.assertTrue(result)
            mock_click.assert_not_called()
            self.assertTrue(
                any(
                    call.args[0] == hwnd and call.args[1] == win32con.WM_CLOSE
                    for call in mock_post.call_args_list
                ),
                "无确定按钮时应发送 WM_CLOSE",
            )
        finally:
            _destroy_window(hwnd)

    def test_ignores_other_titles(self):
        """非"提示"标题的窗口不被处理。"""
        hwnd, _btn = _create_prompt_window("数据输出", "确定")
        try:
            with (
                _mock_enum_windows([hwnd]),
                patch.object(dialog_control, "click") as mock_click,
            ):
                result = dialog_control.close_no_data_prompt(settle_delay=0)
            self.assertFalse(result)
            mock_click.assert_not_called()
        finally:
            _destroy_window(hwnd)

    def test_normal_path_overhead_is_small(self):
        """正常路径（无提示弹窗）调用开销极小，不应影响执行速度。"""
        with _mock_enum_windows([]):
            t0 = time.time()
            for _ in range(50):
                dialog_control.close_no_data_prompt(settle_delay=0)
            elapsed = time.time() - t0
        self.assertLess(elapsed, 1.0, f"50 次 close_no_data_prompt 应 <1s，实际 {elapsed:.3f}s")


if __name__ == "__main__":
    unittest.main()