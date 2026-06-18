# -*- coding: utf-8 -*-
"""数据输出弹窗处理模块（钱龙自带的"数据输出"对话框）"""

import time
import os
import pywinauto.keyboard as keyboard
from pywinauto import Application, findwindows

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


def handle_export_dialog(timeout: float = 8, export_type: str = "txt",
                         auto_open: bool = True, output_path: str = "") -> bool:
    """等待并操作"数据输出"弹窗。

    Args:
        timeout: 等待秒数
        export_type: 导出格式 "txt" 或 "xls"
        auto_open: 是否自动打开文件
        output_path: 输出路径
    Returns:
        bool: 是否成功完成整个导出流程
    """
    end = time.time() + timeout
    dlg_hwnd = None

    # 1. 找到"数据输出"弹窗句柄
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True, visible_only=True):
            if "数据输出" in (elem.name or ""):
                dlg_hwnd = elem.handle
                break
        if dlg_hwnd:
            break
        time.sleep(0.1)
    if not dlg_hwnd:
        print("[WARN] 未找到'数据输出'弹窗")
        return False

    dlg_app = Application(backend="uia").connect(handle=dlg_hwnd, timeout=1)
    dlg = dlg_app.window(handle=dlg_hwnd)
    print(f"[OK] 已找到'数据输出'弹窗,句柄={dlg_hwnd}")

    # 2. 选择导出格式 RadioButton
    radio_id = EXPORT_RADIO_MAP.get(export_type)
    if radio_id:
        radio = dlg.child_window(auto_id=radio_id, control_type="RadioButton")
        if radio.exists(timeout=1):
            radio.click_input()
            print(f"[OK] 已选择导出格式: {export_type}")
        else:
            print(f"[WARN] RadioButton(auto_id={radio_id}) 未找到")

    time.sleep(0.2)

    # 3. 设置自定义输出路径
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

    # 4. 控制"主动打开输出的文件"复选框
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

    # 5. 点击"输  出"按钮(auto_id="1")
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
