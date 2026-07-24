# -*- coding: utf-8 -*-
"""Windows系统"另存为"对话框处理模块（用于资金持仓等脚本）。

优化：优先用 win32gui 直接发送控件消息（WM_SETTEXT / BM_CLICK），不依赖前台
焦点与真实鼠标；失败时回退到 pywinauto（UI Automation）。
"""

import os
import time

import win32gui
from pywinauto import Application, findwindows

from core.dialog_control import (
    IDOK,
    IDYES,
    SAVEAS_FILENAME_ID,
    Win32ControlError,
    click,
    enum_controls,
    find_by_id,
    find_by_text,
    find_edit_near_label,
    find_first_edit,
    find_top_dialog,
    press_enter,
    set_text,
)


def _find_saveas_window(timeout: float):
    """查找 Windows 系统的"另存为"窗口（win32gui 优先, pywinauto 兜底）。"""
    hwnd = find_top_dialog(
        ["另存为", "保存为", "Save As"], timeout=timeout
    )
    if hwnd is not None:
        return hwnd
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                dlg_app = Application(backend="uia").connect(
                    handle=elem.handle, timeout=0.5
                )
                dlg = dlg_app.window(handle=elem.handle)
                title = dlg.window_text() or ""
                if "另存为" in title or "保存为" in title or "Save As" in title:
                    return elem.handle
            except Exception:
                continue
        time.sleep(0.15)
    return None


def _find_saveas_filename_edit(dlg_hwnd: int):
    """定位"另存为"窗口的文件名输入框。

    识别顺序：
      1) 通用对话框文件名 ComboBox(ID=0x47C) 内的 Edit；
      2) 按“文件名/File name”标签定位其右侧、下方最近的编辑框（自定义对话框）；
      3) 兜底：整窗第一个 Edit。
    """
    # 1) 通用对话框文件名 ComboBox(ID=0x47C) 内的 Edit
    ctrl = find_by_id(dlg_hwnd, SAVEAS_FILENAME_ID)
    if ctrl is not None:
        for hwnd, _cid, cls, _text in enum_controls(ctrl):
            if "EDIT" in (cls or "").upper():
                return hwnd
        if "EDIT" in (win32gui.GetClassName(ctrl) or "").upper():
            return ctrl

    # 2) 按“文件名”标签定位其右侧/下方的编辑框
    near = find_edit_near_label(dlg_hwnd, ["文件名", "File name"])
    if near is not None:
        return near

    # 3) 兜底：整窗第一个 Edit
    return find_first_edit(dlg_hwnd)


def _handle_overwrite_prompt(timeout: float = 4) -> bool:
    """处理"文件已存在,是否替换"弹窗(若有)。pywinauto 兜底用。"""
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(
                    handle=hwnd, timeout=0.5
                )
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if not any(
                    k in title for k in ("确认另存为", "文件已存在", "Confirm Save As")
                ):
                    continue
                dlg.set_focus()
                time.sleep(0.1)
                # 直接点击"是(Y)"按钮，避免焦点在"否"上导致按回车点错
                try:
                    yes_btn = dlg.child_window(
                        title_re=r".*是\(Y\).*|.*Yes.*", control_type="Button"
                    )
                    if yes_btn.exists(timeout=0.5):
                        yes_btn.click_input()
                        print(f"[OK] 已点击'是'确认覆盖 (hwnd={hwnd}, title='{title}')")
                        return True
                except Exception:
                    pass
                # 兜底：尝试按回车
                dlg.type_keys("{ENTER}", with_spaces=False)
                print(f"[OK] 已确认覆盖 (hwnd={hwnd}, title='{title}')")
                return True
            except Exception:
                continue
        time.sleep(0.15)
    return True


def _handle_overwrite_prompt_win32(timeout: float = 4) -> bool:
    """处理"文件已存在,是否替换"弹窗(若有)。win32gui 主路径。"""
    hwnd = find_top_dialog(
        ["确认另存为", "文件已存在", "Confirm Save As"], timeout=timeout
    )
    if hwnd is None:
        # 未出现覆盖确认框，视为主流程已成功
        return True
    # 点击"是"(IDYES=6)；找不到则按文字找"是/Yes"
    btn = find_by_id(hwnd, IDYES)
    if not btn:
        btn = find_by_text(hwnd, "是", class_name="Button", partial=True)
    if btn:
        click(btn, hwnd)
        print(f"[OK] (win32) 已点击'是'确认覆盖 (hwnd={hwnd})")
    else:
        press_enter()
        print(f"[OK] (win32) 已确认覆盖 (hwnd={hwnd})")
    return True


def _saveas_win32(dlg_hwnd: int, full_path: str) -> bool:
    """win32gui 主路径：直接发消息设置文件名并点击保存。"""
    edit = _find_saveas_filename_edit(dlg_hwnd)
    if edit is None:
        raise Win32ControlError("找不到文件名输入框")
    set_text(edit, full_path)
    print(f"[OK] (win32) 已填入文件名: {full_path}")
    time.sleep(0.2)

    # 点击“保存”按钮(IDOK=1)，找不到则回退到回车
    btn = find_by_id(dlg_hwnd, IDOK)
    if btn:
        click(btn, dlg_hwnd)
    else:
        press_enter()
    print("[OK] (win32) 已确认保存")
    time.sleep(0.2)
    return True


def _saveas_pywinauto(dlg_hwnd: int, full_path: str) -> bool:
    """pywinauto 兜底路径（原逻辑）。"""
    dlg_app = Application(backend="uia").connect(handle=dlg_hwnd, timeout=1)
    dlg = dlg_app.window(handle=dlg_hwnd)
    dlg.set_focus()
    time.sleep(0.3)

    name_edit = _find_name_edit(dlg)
    if name_edit is None:
        print("[错误] 找不到文件名输入框")
        return False

    name_edit.set_focus()
    time.sleep(0.1)
    name_edit.set_edit_text(full_path)
    print(f"[OK] 已填入文件名: {full_path}")
    time.sleep(0.2)

    dlg.set_focus()
    time.sleep(0.1)
    dlg.type_keys("{ENTER}", with_spaces=False)
    print("[OK] 已按回车确认保存")

    return _handle_overwrite_prompt(timeout=4)


def _find_name_edit(dlg):
    """在另存为窗口中定位文件名输入框。"""
    for title_re in (r".*文件名.*", r".*File name.*", r".*文件名\(N\).*"):
        try:
            edit = dlg.child_window(title_re=title_re, control_type="Edit")
            if edit.exists(timeout=0.5):
                return edit
        except Exception:
            continue
    edits = dlg.descendants(control_type="Edit")
    return edits[0] if edits else None


def handle_save_as_dialog(save_dir: str, filename: str, timeout: float = 8) -> bool:
    """处理 Windows 系统的"另存为"窗口。

    优先用 win32gui 直接发送控件消息；定位或操作失败时回退到 pywinauto。

    Args:
        save_dir: 保存目录
        filename: 文件名(含扩展名)
        timeout: 等待秒数
    Returns:
        bool: 是否成功
    """
    os.makedirs(save_dir, exist_ok=True)
    full_path = os.path.join(save_dir, filename)

    print(f"[..] 等待'另存为'窗口... (路径={full_path})")
    hwnd = _find_saveas_window(timeout)
    if hwnd is None:
        print(f"[WARN] 等待'另存为'窗口超时({timeout}s)")
        return False

    print(f"[OK] 另存为窗口已弹出 (hwnd={hwnd})")
    try:
        _saveas_win32(hwnd, full_path)
        return _handle_overwrite_prompt_win32(timeout=4)
    except Win32ControlError as exc:
        print(f"[WARN] win32gui 操作失败({exc}), 回退 pywinauto")
        return _saveas_pywinauto(hwnd, full_path)
