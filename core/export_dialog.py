# -*- coding: utf-8 -*-
"""数据输出弹窗处理模块（钱龙自带的"数据输出"对话框）。

优化：优先用 win32gui 直接发送控件消息（BM_CLICK / WM_SETTEXT），不依赖前台
焦点与真实鼠标；失败时回退到 pywinauto（UI Automation）。
"""

import time

from pywinauto import Application, findwindows
import pywinauto.keyboard as keyboard

from core.dialog_control import (
    Win32ControlError,
    click,
    find_by_id,
    find_by_text,
    find_top_dialog,
    is_checked,
    press_enter,
    set_text,
)

# RadioButton auto_id 映射
EXPORT_RADIO_MAP = {
    "txt": "1170",   # 输出到文本文件(.txt)
    "xls": "1172",   # 输出到Excel电子表格(.xls)
}

# 路径输入框 auto_id 映射
PATH_INPUT_MAP = {
    "txt": "1176",   # txt路径输入框
    "xls": "1178",   # xls路径输入框
}

OUTPUT_BTN_TEXT = "输  出"   # 输出按钮实际标题（含空格）


def _read_checkbox_uia(dlg_hwnd: int):
    """用 UIA 读取'主动打开输出的文件'勾选状态, 失败返回 None。"""
    try:
        from pywinauto import Application
        cb = (
            Application(backend="uia")
            .connect(handle=dlg_hwnd, timeout=1)
            .window(handle=dlg_hwnd)
            .child_window(title="主动打开输出的文件", control_type="CheckBox")
        )
        if cb.exists(timeout=1):
            try:
                return cb.get_toggle_state() == 1
            except Exception:
                return cb.is_checked()
    except Exception:
        return None
    return None


def _find_dialog(timeout: float):
    """查找"数据输出"弹窗句柄（win32gui 优先, pywinauto 兜底）。"""
    hwnd = find_top_dialog("数据输出", timeout=timeout)
    if hwnd is not None:
        return hwnd
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True, visible_only=True):
            if "数据输出" in (elem.name or ""):
                return elem.handle
        time.sleep(0.1)
    return None


def _export_win32(dlg_hwnd: int, export_type: str, auto_open: bool,
                  output_path: str) -> bool:
    """win32gui 主路径：直接发消息操作控件。"""
    # 1. 选择导出格式 RadioButton
    radio_id = EXPORT_RADIO_MAP.get(export_type)
    if radio_id:
        rhwnd = find_by_id(dlg_hwnd, int(radio_id))
        if not rhwnd:
            raise Win32ControlError(f"找不到导出格式 RadioButton(id={radio_id})")
        click(rhwnd, dlg_hwnd)
        print(f"[OK] (win32) 已选择导出格式: {export_type}")

    # 2. 设置自定义输出路径
    path_id = PATH_INPUT_MAP.get(export_type)
    if path_id and output_path:
        phwnd = find_by_id(dlg_hwnd, int(path_id))
        if not phwnd:
            raise Win32ControlError(f"找不到路径输入框(id={path_id})")
        set_text(phwnd, output_path)
        print(f"[OK] (win32) 已设置{export_type.upper()}输出路径: {output_path}")

    # 3. 控制"主动打开输出的文件"复选框
    cb = find_by_text(dlg_hwnd, "主动打开输出的文件", class_name="Button")
    if cb:
        # 弹窗中 BM_GETCHECK 可能失效, 优先用 UIA 读取勾选状态
        checked = _read_checkbox_uia(dlg_hwnd)
        if checked is None:
            checked = is_checked(cb)
        if auto_open and not checked:
            click(cb, dlg_hwnd)
            print("[OK] (win32) 已勾选'主动打开输出的文件'")
        elif not auto_open and checked:
            click(cb, dlg_hwnd)
            print("[OK] (win32) 已取消勾选'主动打开输出的文件'")
        else:
            print(
                f"[OK] (win32) '主动打开输出的文件'保持"
                f"{'勾选' if auto_open else '未勾选'}状态"
            )
    else:
        print("[WARN] (win32) 未找到'主动打开输出的文件'复选框")

    # 4. 点击"输  出"按钮：优先按 control id=1，否则按文字
    btn = find_by_id(dlg_hwnd, 1)
    if not btn:
        btn = find_by_text(dlg_hwnd, OUTPUT_BTN_TEXT, class_name="Button")
    if not btn:
        raise Win32ControlError("找不到'输出'按钮")
    click(btn, dlg_hwnd)
    print("[OK] (win32) 已点击'输  出'按钮,导出完成")

    # 确认后续"导出成功"等提示框
    time.sleep(0.5)
    press_enter()
    return True


def _export_pywinauto(dlg_hwnd: int, export_type: str, auto_open: bool,
                      output_path: str) -> bool:
    """pywinauto 兜底路径（原逻辑）。"""
    dlg_app = Application(backend="uia").connect(handle=dlg_hwnd, timeout=1)
    dlg = dlg_app.window(handle=dlg_hwnd)

    # 选择导出格式 RadioButton
    radio_id = EXPORT_RADIO_MAP.get(export_type)
    if radio_id:
        radio = dlg.child_window(auto_id=radio_id, control_type="RadioButton")
        if radio.exists(timeout=1):
            radio.click_input()
            print(f"[OK] 已选择导出格式: {export_type}")
        else:
            print(f"[WARN] RadioButton(auto_id={radio_id}) 未找到")

    time.sleep(0.2)

    # 设置自定义输出路径
    path_input_id = PATH_INPUT_MAP.get(export_type)
    if path_input_id and output_path:
        path_edit = dlg.child_window(auto_id=path_input_id, control_type="Edit")
        if path_edit.exists(timeout=1):
            path_edit.wait("ready", timeout=2)
            path_edit.click_input()
            time.sleep(0.1)
            path_edit.type_keys("^a", with_spaces=False)
            time.sleep(0.05)
            path_edit.set_edit_text(output_path)
            print(f"[OK] 已设置{export_type.upper()}输出路径: {output_path}")
        else:
            print(f"[WARN] 路径输入框(auto_id={path_input_id}) 未找到")

    time.sleep(0.2)

    # 控制"主动打开输出的文件"复选框
    cb = dlg.child_window(title="主动打开输出的文件", control_type="CheckBox")
    if cb.exists(timeout=1):
        try:
            now_checked = cb.get_toggle_state() == 1
        except Exception:
            now_checked = cb.is_checked()

        if auto_open and not now_checked:
            cb.click_input()
            print("[OK] 已勾选'主动打开输出的文件'")
        elif not auto_open and now_checked:
            cb.click_input()
            print("[OK] 已取消勾选'主动打开输出的文件'")
        else:
            print(f"[OK] '主动打开输出的文件'保持{'勾选' if auto_open else '未勾选'}状态")
    else:
        print("[WARN] 未找到'主动打开输出的文件'复选框")

    time.sleep(0.2)

    # 点击"输  出"按钮(auto_id="1")
    btn_output = dlg.child_window(auto_id="1", control_type="Button")
    if btn_output.exists(timeout=1):
        btn_output.wait("ready", timeout=2)
        btn_output.click_input()
        print("[OK] 已点击'输  出'按钮,导出完成")

        # 发送Enter确认后续小弹窗(如"导出成功"提示框)
        time.sleep(0.5)
        keyboard.send_keys("{ENTER}")

        return True
    else:
        print("[错误] 未找到'输  出'按钮")
        return False


def handle_export_dialog(timeout: float = 8, export_type: str = "txt",
                         auto_open: bool = True, output_path: str = "") -> bool:
    """等待并操作"数据输出"弹窗。

    优先用 win32gui 直接发送控件消息；定位或操作失败时回退到 pywinauto。

    Args:
        timeout: 等待秒数
        export_type: 导出格式 "txt" 或 "xls"
        auto_open: 是否自动打开文件
        output_path: 输出路径
    Returns:
        bool: 是否成功完成整个导出流程
    """
    dlg_hwnd = _find_dialog(timeout)
    if not dlg_hwnd:
        print("[WARN] 未找到'数据输出'弹窗")
        return False
    print(f"[OK] 已找到'数据输出'弹窗,句柄={dlg_hwnd}")

    try:
        return _export_win32(dlg_hwnd, export_type, auto_open, output_path)
    except Win32ControlError as exc:
        print(f"[WARN] win32gui 操作失败({exc}), 回退 pywinauto")
        return _export_pywinauto(dlg_hwnd, export_type, auto_open, output_path)
