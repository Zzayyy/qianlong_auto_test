# -*- coding: utf-8 -*-
"""探索广发证券期权宝'交易系统设置'窗口的导航控件结构（临时诊断脚本）。

用法：
    1. 登录广发证券期权宝客户端
    2. 运行: python z局外_探索广发交易系统设置.py
    3. 脚本会自动打开'交易系统设置'窗口并 dump 控件结构到
       z局外_广发交易系统设置_dump.txt

输出说明：
    - 第一部分: Win32 枚举（每个子控件 类名|控件ID|文本|可见|启用），
      重点看左侧导航控件是 ListBox 还是 SysTreeView32 及其控件ID
    - 第二部分: UIA 控件树（含 auto_id / control_type），
      用于确认导航项（ListItem/TreeItem）的层级
"""

import os
import sys
import io

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import win32gui

from core.window import find_window, activate_window, close_settings_dialog
from core.settings_window import open_settings_dialog

CLIENT_NAME = "广发证券期权宝"
OUT_FILE = "z局外_广发交易系统设置_dump.txt"


def dump_win32_controls(hwnd, f):
    """用 EnumChildWindows 枚举设置窗口的所有子控件。"""
    f.write("\n===== Win32 子控件枚举 =====\n")
    rows = []

    def _enum(child, _):
        try:
            cls = win32gui.GetClassName(child)
            cid = win32gui.GetDlgCtrlID(child)
            text = win32gui.GetWindowText(child) or ""
            visible = win32gui.IsWindowVisible(child)
            enabled = win32gui.IsWindowEnabled(child)
            rows.append((cls, cid, text, visible, enabled))
        except Exception:
            pass

    win32gui.EnumChildWindows(hwnd, _enum, None)
    for cls, cid, text, visible, enabled in rows:
        f.write(
            f"class={cls:<26} ctrl_id={cid:<6} visible={int(visible)} "
            f"enabled={int(enabled)} text={text!r}\n"
        )
    f.write(f"\n共 {len(rows)} 个子控件\n")


def main():
    with open(OUT_FILE, "w", encoding="utf-8") as f:
        # 1. 找到主窗口并激活
        main_hwnd = find_window(CLIENT_NAME)
        f.write(f"主窗口 hwnd={main_hwnd}\n")
        print(f"[OK] 主窗口 hwnd={main_hwnd}")
        win = activate_window(main_hwnd)

        # 2. 打开（或复用）交易系统设置窗口
        dlg = open_settings_dialog(win)
        dlg.wait("ready", timeout=10)
        dlg_hwnd = int(dlg.handle)
        f.write(f"设置窗口 hwnd={dlg_hwnd} title={dlg.window_text()!r}\n")
        print(f"[OK] 设置窗口 hwnd={dlg_hwnd}")

        # 3. Win32 枚举（类名 / 控件ID / 文本）
        dump_win32_controls(dlg_hwnd, f)

        # 4. UIA 控件树
        f.write("\n===== UIA 控件树 (print_control_identifiers) =====\n")
        try:
            buf = io.StringIO()
            dlg.print_control_identifiers(depth=None, stdout=buf)
            f.write(buf.getvalue())
        except Exception as e:
            f.write(f"UIA dump 失败: {e}\n")

    print(f"[OK] 结果已写入 {OUT_FILE}")
    # 关闭设置窗口
    try:
        close_settings_dialog(dlg, keep_open=False, main_hwnd=main_hwnd)
    except Exception as e:
        print(f"[WARN] 关闭设置窗口失败: {e}")


if __name__ == "__main__":
    main()
