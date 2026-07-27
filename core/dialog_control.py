# -*- coding: utf-8 -*-
"""基于 win32gui/win32api 的对话框控件操作。用于数据输出和系统级另存为

点击采用“方法二”：用 win32api 把光标移到控件中心并发送真实鼠标左键事件
（``SetCursorPos`` + ``mouse_event``）。相比 pywinauto 的 ``click_input()`` 它不
依赖 UIA 树；相比“方法一”的 ``SendMessage(BM_CLICK)``，真实鼠标点击对弹窗
（自定义/非标准按钮）更可靠——弹窗中 ``BM_CLICK`` 常常无效，真实鼠标点击始终有效。

文本设置仍用 ``WM_SETTEXT``（Edit 控件原生处理，可靠且无需焦点）。复选框状态优先
用 UIA 读取以保证正确，失败再回退 ``BM_GETCHECK``。

当控件无法用 win32 方式定位时，调用方应回退到 pywinauto。
"""

import time

import win32api
import win32con
import win32gui

WM_SETTEXT = win32con.WM_SETTEXT
BM_GETCHECK = 0x00F0
BST_CHECKED = 0x0001

# 标准通用对话框常用控件 ID
SAVEAS_FILENAME_ID = 0x047C   # 1148: “另存为”文件名 ComboBox/Edit
IDOK = 0x0001                 # 1: 默认“保存/确定”按钮
IDYES = 0x0006                # 6: “是”按钮


class Win32ControlError(RuntimeError):
    """win32 方式定位或操作控件失败, 调用方应回退到 pywinauto。"""


def find_top_dialog(keywords, timeout: float = 8, visible_only: bool = True):
    """按标题关键字查找顶层窗口, 返回第一个匹配句柄, 超时返回 None。

    Args:
        keywords: 标题需包含的子串（列表或单个字符串）。
        timeout: 等待秒数。
        visible_only: 是否只匹配可见窗口。
    """
    if isinstance(keywords, str):
        keywords = [keywords]
    end = time.time() + timeout
    while time.time() < end:
        found = []

        def _enum(hwnd, _):
            try:
                title = win32gui.GetWindowText(hwnd) or ""
                # 排除本自动化工具自身的窗口, 避免误匹配
                if "GUI自动化工具" in title:
                    return
                if visible_only and not win32gui.IsWindowVisible(hwnd):
                    return
                if any(keyword in title for keyword in keywords):
                    found.append((hwnd, title))
            except Exception:
                return

        win32gui.EnumWindows(_enum, None)
        if found:
            return found[0][0]
        time.sleep(0.1)
    return None


def enum_controls(parent_hwnd: int):
    """枚举 parent 的所有后代控件, 返回 [(hwnd, ctrl_id, class_name, text), ...]。"""
    result = []

    def _enum(hwnd, _):
        try:
            ctrl_id = win32gui.GetDlgCtrlID(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            text = win32gui.GetWindowText(hwnd)
            result.append((hwnd, ctrl_id, class_name, text))
        except Exception:
            return

    win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    return result


def find_by_id(parent_hwnd: int, ctrl_id: int, class_name: str = None):
    """按资源 ID 查找子控件句柄（可选匹配类名）。"""
    for hwnd, cid, cls, _text in enum_controls(parent_hwnd):
        if cid == ctrl_id and (class_name is None or cls == class_name):
            return hwnd
    return None


def find_by_text(parent_hwnd: int, text: str, class_name: str = None,
                 partial: bool = False):
    """按窗口文字查找子控件句柄（可选匹配类名）。"""
    for hwnd, _cid, cls, txt in enum_controls(parent_hwnd):
        if class_name and cls != class_name:
            continue
        target = txt or ""
        if partial:
            if text in target:
                return hwnd
        elif target == text:
            return hwnd
    return None


def find_first_edit(parent_hwnd: int):
    """返回 parent 下第一个 Edit 控件句柄（兜底用）。"""
    for hwnd, _cid, cls, _text in enum_controls(parent_hwnd):
        if cls == "Edit":
            return hwnd
    return None


def find_edit_near_label(parent_hwnd: int, label_substrings):
    """在 parent 下找到与给定标签文字最近的编辑框（类名含 EDIT）。

    适用于“另存为”等自定义对话框：用“文件名”标签定位其右侧/下方的输入框，
    避免误选“文件类型”等其它编辑框。

    Args:
        label_substrings: 标签需包含的子串列表，如 ["文件名", "File name"]。
    Returns:
        编辑框句柄或 None。
    """
    label_cx = label_cy = None
    edits = []
    for hwnd, _cid, cls, txt in enum_controls(parent_hwnd):
        c = (cls or "").upper()
        if "EDIT" in c:
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            edits.append((hwnd, (l + r) / 2, (t + b) / 2))
        elif txt and any(s.lower() in txt.lower() for s in label_substrings):
            l, t, r, b = win32gui.GetWindowRect(hwnd)
            label_cx, label_cy = (l + r) / 2, (t + b) / 2
    if label_cx is None or not edits:
        return None
    edits.sort(
        key=lambda it: abs(it[1][0] - label_cx) + abs(it[1][1] - label_cy)
    )
    return edits[0][0]


def click(hwnd: int, parent_hwnd: int = None):
    """方法二：真实鼠标点击控件中心（适用于弹窗）。

    先把父窗口置前（绕过前台锁），再把光标移到控件中心发送鼠标左键事件。
    """
    if parent_hwnd is not None:
        try:
            if win32gui.IsIconic(parent_hwnd):
                win32gui.ShowWindow(parent_hwnd, win32con.SW_RESTORE)
            # 先置顶再设前台可绕过系统“前台锁”，最后取消置顶
            win32gui.SetWindowPos(
                parent_hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
            )
            win32gui.SetForegroundWindow(parent_hwnd)
            win32gui.SetWindowPos(
                parent_hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
            )
            time.sleep(0.15)
        except Exception:
            pass

    l, t, r, b = win32gui.GetWindowRect(hwnd)
    x = (l + r) // 2
    y = (t + b) // 2
    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.1)


def set_text(hwnd: int, text: str):
    """通过 WM_SETTEXT 消息设置控件文字（Edit 原生处理）。"""
    win32gui.SendMessage(hwnd, WM_SETTEXT, 0, text)


def get_text(hwnd: int) -> str:
    """读取控件文字。"""
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def is_checked(hwnd: int) -> bool:
    """读取复选/单选控件勾选状态（BM_GETCHECK 兜底）。"""
    return win32gui.SendMessage(hwnd, BM_GETCHECK, 0, 0) == BST_CHECKED


def press_enter():
    """向系统发送一次回车键（兜底确认弹窗用）。"""
    win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
    time.sleep(0.05)
    win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)
