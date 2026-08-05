# -*- coding: utf-8 -*-
"""逐面板探索广发证券期权宝'交易系统设置'各页控件 ID（临时诊断脚本）。

用法：
    1. 登录广发证券期权宝客户端
    2. 运行: python z局外_探索广发交易系统设置_面板.py
    3. 脚本自动打开'交易系统设置'窗口，依次切换每个左侧面板，
       并 dump 该页所有**可见**子控件（类名|控件ID|文本）到
       z局外_广发交易系统设置_面板_dump.txt

输出说明：
    - 每段以"===== 面板: <名称> =====" 分隔
    - 控件 ID（ctrl_id）即后续适配 AUTO_ID 映射的依据
    - 若某面板名在导航中不存在会打印 [WARN] 跳过
"""

import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import win32gui

from core.window import find_window, activate_window, close_settings_dialog
from core.settings_window import (
    open_settings_dialog,
    switch_settings_panel,
    _find_native_nav_list,
)

CLIENT_NAME = "广发证券期权宝"
OUT_FILE = "z局外_广发交易系统设置_面板_dump.txt"


def dump_visible_controls(dialog, f, indent="  "):
    """枚举对话框内所有可见子控件（递归），输出 类名|控件ID|文本|相对坐标。"""
    rows = []

    def _enum(child, _):
        try:
            if not win32gui.IsWindowVisible(child):
                return
            cls = win32gui.GetClassName(child)
            cid = win32gui.GetDlgCtrlID(child)
            text = win32gui.GetWindowText(child) or ""
            rows.append((child, cls, cid, text))
        except Exception:
            pass

    dialog_hwnd = int(dialog.handle)
    win32gui.EnumChildWindows(dialog_hwnd, _enum, None)
    base = win32gui.GetWindowRect(dialog_hwnd)

    def _pos(child):
        try:
            r = win32gui.GetWindowRect(child)
            return f"({r[0]-base[0]},{r[1]-base[1]},{r[2]-base[0]},{r[3]-base[1]})"
        except Exception:
            return ""

    for child, cls, cid, text in rows:
        f.write(
            f"{indent}class={cls:<26} ctrl_id={cid:<6} pos={_pos(child):<28} text={text!r}\n"
        )
    f.write(f"{indent}共 {len(rows)} 个可见控件\n")


def main():
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        # 1. 主窗口
        main_hwnd = find_window(CLIENT_NAME)
        f.write(f"主窗口 hwnd={main_hwnd}\n")
        print(f"[OK] 主窗口 hwnd={main_hwnd}")
        win = activate_window(main_hwnd)

        # 2. 打开设置窗口
        dlg = open_settings_dialog(win)
        dlg.wait("ready", timeout=10)
        f.write(f"设置窗口 hwnd={int(dlg.handle)}\n")
        print(f"[OK] 设置窗口 hwnd={int(dlg.handle)}")

        # 3. 读取导航面板列表（UIA ListItem）
        list_hwnd = _find_native_nav_list(int(dlg.handle), 2210)
        if list_hwnd is None:
            f.write("[错误] 未找到导航 ListBox\n")
            print("[错误] 未找到导航 ListBox")
            return
        navigation = dlg.app.window(handle=list_hwnd)
        items = navigation.descendants(control_type="ListItem")
        names = [(item.window_text() or "").strip() for item in items]
        names = [n for n in names if n]
        f.write(f"导航面板列表: {'、'.join(names)}\n")
        print(f"[OK] 导航面板列表: {'、'.join(names)}")

        # 4. 依次切换每个面板并 dump
        for name in names:
            print(f"[..] 切换面板: {name}")
            ok = switch_settings_panel(dlg, name)
            time.sleep(0.6)
            f.write(f"\n===== 面板: {name} (切换{'成功' if ok else '失败'}) =====\n")
            if ok:
                dump_visible_controls(dlg, f)

    print(f"[OK] 结果已写入 {OUT_FILE}")
    try:
        close_settings_dialog(dlg, keep_open=False, main_hwnd=main_hwnd)
    except Exception as e:
        print(f"[WARN] 关闭设置窗口失败: {e}")


if __name__ == "__main__":
    main()
