# -*- coding: utf-8 -*-
"""Shared Win10/Win11 helpers for the QLOption settings dialog."""

from __future__ import annotations

import time

import win32con
import win32gui
import win32process
from pywinauto import Application


def _same_process(hwnd: int, target_pid: int) -> bool:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid == target_pid
    except Exception:
        return False


def _dialog_title(dialog) -> str:
    try:
        title = dialog.window_text() or ""
        if title:
            return title
    except Exception:
        pass
    try:
        titlebar = dialog.child_window(control_type="TitleBar")
        if titlebar.exists(timeout=0.2):
            legacy = titlebar.legacy_properties()
            return (legacy.get("Value") or titlebar.element_info.name or "").strip()
    except Exception:
        pass
    return ""


def find_settings_dialog(main_window, title: str = "交易系统设置"):
    """Find only a settings dialog owned by the target trading process."""
    main_hwnd = int(main_window.handle)
    _, target_pid = win32process.GetWindowThreadProcessId(main_hwnd)
    candidates: list[int] = []

    def _enum(hwnd, _):
        try:
            if (
                _same_process(hwnd, target_pid)
                and win32gui.IsWindowVisible(hwnd)
                and win32gui.GetClassName(hwnd) == "#32770"
            ):
                candidates.append(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(_enum, None)

    # 部分 Win10 客户端把设置框实现为主窗口的子对话框，而 Win11 版本是
    # 独立顶级对话框；两种层级都纳入，但仍严格限定在交易进程内。
    def _enum_child(hwnd, _):
        try:
            if (
                _same_process(hwnd, target_pid)
                and win32gui.IsWindowVisible(hwnd)
                and win32gui.GetClassName(hwnd) == "#32770"
            ):
                candidates.append(hwnd)
        except Exception:
            pass

    win32gui.EnumChildWindows(main_hwnd, _enum_child, None)
    candidates = list(dict.fromkeys(candidates))
    candidates.sort(
        key=lambda hwnd: (
            0 if win32gui.IsWindowVisible(hwnd) else 1,
            0 if title in (win32gui.GetWindowText(hwnd) or "") else 1,
        )
    )
    for hwnd in candidates:
        native_title = (win32gui.GetWindowText(hwnd) or "").strip()
        if title in native_title:
            try:
                dialog = main_window.app.window(handle=hwnd)
            except Exception:
                dialog = Application(backend="uia").connect(
                    handle=hwnd, timeout=1
                ).window(handle=hwnd)
            print(
                f"[OK] 已找到'{title}'窗口: hwnd={hwnd}, "
                f"title={native_title!r}"
            )
            return dialog

        try:
            try:
                # 复用主窗口已经建立的 UIA Application，避免为同一交易进程
                # 重复 connect（跨完整性级别时可能明显变慢）。
                dialog = main_window.app.window(handle=hwnd)
            except Exception:
                app = Application(backend="uia").connect(handle=hwnd, timeout=1)
                dialog = app.window(handle=hwnd)
            actual_title = _dialog_title(dialog)
            navigation = dialog.child_window(auto_id="2210", control_type="List")
            if title in actual_title or navigation.exists(timeout=0.3):
                print(
                    f"[OK] 已找到'{title}'窗口: hwnd={hwnd}, "
                    f"title={actual_title!r}"
                )
                return dialog
        except Exception:
            continue
    return None


def _click_native_button(parent_hwnd: int, control_id: int) -> bool:
    matches: list[int] = []

    def _enum(hwnd, _):
        try:
            if (
                win32gui.GetDlgCtrlID(hwnd) == control_id
                and win32gui.IsWindowVisible(hwnd)
                and win32gui.IsWindowEnabled(hwnd)
            ):
                matches.append(hwnd)
        except Exception:
            pass

    win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    if len(matches) != 1:
        return False
    # BM_CLICK 会在目标线程内进入 TrackPopupMenu 的模态消息循环；若使用
    # SendMessage，调用方会一直等到菜单关闭。PostMessage 可避免这里卡死。
    win32gui.PostMessage(matches[0], win32con.BM_CLICK, 0, 0)
    return True


def _click_settings_button(main_window, control_id: str) -> None:
    if _click_native_button(int(main_window.handle), int(control_id)):
        print(f"[OK] 已用原生消息点击设置按钮(auto_id={control_id})")
        time.sleep(0.4)
        return

    button = main_window.child_window(auto_id=control_id, control_type="Button")
    button.wait("enabled", timeout=5)
    try:
        button.invoke()
    except Exception:
        button.click_input()
    print(f"[OK] 已通过UIA点击设置按钮(auto_id={control_id})")
    time.sleep(0.4)


def _invoke_settings_menu(main_window, menu_auto_id: str,
                          timeout: float = 4.0) -> bool:
    _, target_pid = win32process.GetWindowThreadProcessId(int(main_window.handle))
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        menus: list[int] = []

        def _enum(hwnd, _):
            try:
                if (
                    _same_process(hwnd, target_pid)
                    and win32gui.IsWindowVisible(hwnd)
                    and win32gui.GetClassName(hwnd) == "#32768"
                ):
                    menus.append(hwnd)
            except Exception:
                pass

        win32gui.EnumWindows(_enum, None)
        for hwnd in menus:
            # 标准 Win32 弹出菜单可以直接读取 HMENU 和命令 ID，不需要读取
            # 对方进程内存。MN_SELECTITEM + Enter 能精确选择 20025，避免
            # Win11 下 UIA 连接菜单超时，也避免用固定坐标或固定 Down 次数。
            try:
                response = win32gui.SendMessageTimeout(
                    hwnd, 0x01E1, 0, 0, win32con.SMTO_ABORTIFHUNG, 1000
                )  # MN_GETHMENU
                menu_handle = response[1] if isinstance(response, tuple) else response
                count = win32gui.GetMenuItemCount(menu_handle)
                command_id = int(menu_auto_id)
                command_ids = [
                    win32gui.GetMenuItemID(menu_handle, index)
                    for index in range(count)
                ]
                if command_id in command_ids:
                    target_index = command_ids.index(command_id)
                    win32gui.SendMessageTimeout(
                        hwnd, 0x01E5, target_index, 0,
                        win32con.SMTO_ABORTIFHUNG, 1000
                    )  # MN_SELECTITEM
                    win32gui.PostMessage(
                        hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0
                    )
                    win32gui.PostMessage(
                        hwnd, win32con.WM_KEYUP, win32con.VK_RETURN, 0
                    )
                    print(
                        f"[OK] 已用原生菜单命令打开交易系统设置"
                        f"(command_id={command_id}, position={target_index})"
                    )
                    return True
            except Exception:
                pass

            try:
                menu = Application(backend="uia").connect(
                    handle=hwnd, timeout=0.5
                ).window(handle=hwnd)
                item = menu.child_window(
                    auto_id=menu_auto_id, control_type="MenuItem"
                )
                if not item.exists(timeout=0.2):
                    continue
                try:
                    item.invoke()
                except Exception:
                    item.click_input()
                print(f"[OK] 已打开交易系统设置(auto_id={menu_auto_id})")
                return True
            except Exception:
                continue
        time.sleep(0.1)
    return False


def open_settings_dialog(main_window, settings_button_auto_id: str = "1008",
                         menu_item_auto_id: str = "20025",
                         title: str = "交易系统设置"):
    """Reuse or open the settings dialog without DPI-dependent coordinates."""
    existing = find_settings_dialog(main_window, title)
    if existing is not None:
        return existing

    for attempt in range(2):
        print(f"[INFO] 正在打开'{title}'（第{attempt + 1}次）...")
        _click_settings_button(main_window, settings_button_auto_id)
        if not _invoke_settings_menu(main_window, menu_item_auto_id):
            continue

        # 当前 Win11 客户端首次创建该对话框有时需要十余秒。等待足够长再
        # 重试，避免第一扇窗口刚出现时又误开第二扇同名窗口。
        deadline = time.monotonic() + 20
        while time.monotonic() < deadline:
            dialog = find_settings_dialog(main_window, title)
            if dialog is not None:
                return dialog
            time.sleep(0.2)

    raise RuntimeError(
        f"无法打开'{title}'。已尝试原生按钮消息和UIA菜单 Invoke；"
        f"设置按钮={settings_button_auto_id}，菜单项={menu_item_auto_id}"
    )


def switch_settings_panel(dialog, panel_name: str,
                          navigation_auto_id: str = "2210") -> bool:
    """Select a settings page with a native ListBox notification first."""
    try:
        dialog_hwnd = int(dialog.handle)
        native_lists: list[int] = []

        def _enum(hwnd, _):
            try:
                if (
                    win32gui.GetDlgCtrlID(hwnd) == int(navigation_auto_id)
                    and win32gui.GetClassName(hwnd) == "ListBox"
                    and win32gui.IsWindowVisible(hwnd)
                ):
                    native_lists.append(hwnd)
            except Exception:
                pass

        win32gui.EnumChildWindows(dialog_hwnd, _enum, None)
        if len(native_lists) == 1:
            list_hwnd = native_lists[0]
            navigation = dialog.app.window(handle=list_hwnd)
        else:
            navigation = dialog.child_window(
                auto_id=navigation_auto_id, control_type="List"
            )
            navigation.wait("ready", timeout=5)
            list_hwnd = int(navigation.wrapper_object().handle)

        items = navigation.descendants(control_type="ListItem")
        names = [(item.window_text() or "").strip() for item in items]
        if panel_name not in names:
            raise RuntimeError(
                f"未找到设置页 {panel_name!r}；当前设置页: "
                f"{'、'.join(name for name in names if name)}"
            )
        target_index = names.index(panel_name)
        parent_hwnd = win32gui.GetParent(list_hwnd)
        control_id = win32gui.GetDlgCtrlID(list_hwnd)

        result = win32gui.SendMessage(
            list_hwnd, win32con.LB_SETCURSEL, target_index, 0
        )
        if result != win32con.LB_ERR:
            win32gui.SendMessage(
                parent_hwnd,
                win32con.WM_COMMAND,
                control_id | (win32con.LBN_SELCHANGE << 16),
                list_hwnd,
            )
            time.sleep(0.6)
            selected = win32gui.SendMessage(
                list_hwnd, win32con.LB_GETCURSEL, 0, 0
            )
            if selected == target_index:
                print(
                    f"[OK] 已切换到'{panel_name}'设置页"
                    f"（原生ListBox，位置={target_index}）"
                )
                return True

        item = items[target_index]
        try:
            item.select()
        except Exception:
            item.click_input()
        time.sleep(0.6)
        print(f"[OK] 已切换到'{panel_name}'设置页（UIA）")
        return True
    except Exception as error:
        print(f"[WARN] 切换到'{panel_name}'设置页失败: {error}")
        return False
