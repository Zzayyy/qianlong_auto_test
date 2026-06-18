# -*- coding: utf-8 -*-
"""Windows系统"另存为"对话框处理模块（用于资金持仓等脚本）"""

import time
import os
from pywinauto import Application, findwindows


def _find_saveas_window(timeout: float):
    """查找 Windows 系统的"另存为"窗口。"""
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if "另存为" in title or "保存为" in title or "Save As" in title:
                    return hwnd, dlg
            except Exception:
                continue
        time.sleep(0.15)
    return None, None


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


def _handle_overwrite_prompt(timeout: float = 4) -> bool:
    """处理"文件已存在,是否替换"弹窗(若有)。"""
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if not any(k in title for k in ("确认另存为", "文件已存在", "Confirm Save As")):
                    continue
                dlg.set_focus()
                time.sleep(0.1)
                # 直接点击"是(Y)"按钮，避免焦点在"否"上导致按回车点错
                try:
                    yes_btn = dlg.child_window(title_re=r".*是\(Y\).*|.*Yes.*", control_type="Button")
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


def handle_save_as_dialog(save_dir: str, filename: str, timeout: float = 8) -> bool:
    """处理 Windows 系统的"另存为"窗口。

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
    hwnd, dlg = _find_saveas_window(timeout)
    if hwnd is None:
        print(f"[WARN] 等待'另存为'窗口超时({timeout}s)")
        return False

    print(f"[OK] 另存为窗口已弹出 (hwnd={hwnd})")
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
