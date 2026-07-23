# -*- coding: utf-8 -*-
"""
交易系统设置 - 自动追单设置界面自动化测试
============================================
功能：
    打开"交易系统设置"对话框，进入"自动追单设置"标签页，
    逐项读取每个参数的当前值（复选框状态、下拉框选择、数值等），
    与标准值（恢复默认后的参数）比对，记录差异并截图。

界面元素（自动追单设置）：
    - 启用自动追单（CheckBox，默认未勾选）
    - 使用（ComboBox，默认"对手价"）
    - 追价，调整（ComboBox，默认 1）
    - 自动追单时间间隔（ComboBox，默认 4 秒）
    - 秒，最多重复（Edit，默认 2 次）
    - 未完成则自动撤单（CheckBox，默认未勾选）

使用方法：
    1. 打开钱龙旗舰版，登录交易账号
    2. 运行本脚本

依赖：
    pip install pywinauto pillow
"""

import os
import sys
import time
import ctypes
import win32gui
import win32con
from datetime import datetime
from typing import Any, Dict, List, Optional

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image
from pywinauto import Application, findwindows
from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.settings_result import SettingsTestResult


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "自动追单设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对
STANDARD_VALUES = {
    "启用自动追单": False,
    "使用": "对手价",
    "使用_选项列表": ["对手价", "挂盘价", "涨停价", "跌停价", "最新价"],  # 常见候选项，可据实调整
    "追价，调整": 1,
    "追价，调整_选项列表": ["0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"],  # 可据实调整
    "自动追单时间间隔": 4,
    "自动追单时间间隔_选项列表": ["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20"],  # 可据实调整
    "秒，最多重复": 2,
    "未完成则自动撤单": False,
}

# 控件 auto_id 映射（来自交易系统设置_自动追单设置.txt 抓取）
AUTO_ID = {
    "启用自动追单": "2024",
    "使用": "2098",
    "追价，调整": "2124",
    "自动追单时间间隔": "2104",
    "秒，最多重复": "2205",
    "未完成则自动撤单": "2064",
}

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "自动追单设置"
COUNTDOWN_SEC = 3  # 倒计时秒数
# ========================================================


def open_settings_dialog(win) -> Any:
    """自动打开'交易系统设置'对话框。"""
    # 先检查是否已打开
    dlg = _find_existing_settings_dlg(win)
    if dlg is not None:
        return dlg

    print("\n正在通过工具栏按钮打开'交易系统设置'...")
    _click_settings_button(win)
    _click_context_menu_item(win)

    print("  等待设置对话框弹出...")
    time.sleep(1.0)
    end = time.time() + 10
    while time.time() < end:
        dlg = _find_existing_settings_dlg(win)
        if dlg is not None:
            return dlg
        time.sleep(0.3)

    raise RuntimeError(
        f"无法打开'{SETTINGS_DIALOG_TITLE}'对话框，请确认:\n"
        f"  1. 钱龙软件已登录交易账号\n"
        f"  2. 工具栏设置按钮(auto_id={SETTINGS_BUTTON_AUTO_ID})可见"
    )


def _find_existing_settings_dlg(win=None) -> Optional[Any]:
    """查找已打开的设置对话框。"""
    # 1. 顶级窗口
    try:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                title = elem.window_text() or ""
                if SETTINGS_DIALOG_TITLE in title:
                    app = Application(backend="uia").connect(handle=elem.handle)
                    dlg = app.window(handle=elem.handle)
                    dlg.wait("ready", timeout=3)
                    print(f"[OK] 已找到设置对话框(顶级): {title}")
                    return dlg
            except Exception:
                continue
    except Exception:
        pass

    # 2. #32770 对话框
    try:
        for elem in findwindows.find_elements(top_level_only=True, class_name="#32770"):
            try:
                app = Application(backend="uia").connect(handle=elem.handle, timeout=1)
                dlg = app.window(handle=elem.handle)
                tb = dlg.child_window(control_type="TitleBar")
                if tb.exists(timeout=0.5):
                    val = ""
                    try:
                        val = tb.legacy_properties().get("Value", "")
                    except Exception:
                        pass
                    if not val:
                        try:
                            val = tb.element_info.name or ""
                        except Exception:
                            pass
                    if SETTINGS_DIALOG_TITLE in val:
                        dlg.wait("ready", timeout=3)
                        print(f"[OK] 已找到设置对话框(#32770): {val}")
                        return dlg
            except Exception:
                continue
    except Exception:
        pass

    # 3. 主窗口的子窗口
    if win is not None:
        try:
            for child in win.descendants(control_type="Window"):
                try:
                    title = child.window_text() or ""
                    if SETTINGS_DIALOG_TITLE in title:
                        child.wait("ready", timeout=3)
                        print(f"[OK] 已找到设置对话框(子窗口): {title}")
                        return child
                except Exception:
                    continue
        except Exception:
            pass

    return None


def _click_settings_button(win):
    """点击主窗口上的设置按钮。"""
    try:
        btn = win.child_window(auto_id=SETTINGS_BUTTON_AUTO_ID, control_type="Button")
        btn.wait("enabled", timeout=5)
        btn.click_input()
        print(f"[OK] 已点击设置按钮 (auto_id={SETTINGS_BUTTON_AUTO_ID})")
        time.sleep(0.5)
    except Exception:
        print(f"[WARN] 用 auto_id 查找设置按钮失败，改用 win32gui...")
        _click_button_by_auto_id(win.handle, SETTINGS_BUTTON_AUTO_ID)


def _click_button_by_auto_id(parent_hwnd: int, auto_id: str):
    """用 win32gui 枚举子窗口查找指定 auto_id 的按钮并点击。"""
    import win32gui
    import win32api
    import win32con

    def callback(hwnd, _):
        try:
            ctrl_id = win32gui.GetDlgCtrlID(hwnd)
            if str(ctrl_id) == auto_id:
                rect = win32gui.GetWindowRect(hwnd)
                x = (rect[0] + rect[2]) // 2
                y = (rect[1] + rect[3]) // 2
                win32api.SetCursorPos((x, y))
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
                win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
                print(f"[OK] 已用 win32gui 点击设置按钮 (auto_id={auto_id})")
        except Exception:
            pass

    win32gui.EnumChildWindows(parent_hwnd, callback, None)
    time.sleep(0.5)


def _click_context_menu_item(win):
    """在弹出菜单中选择'交易系统设置'。"""
    end = time.time() + 4
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                menu_app = Application(backend="uia").connect(handle=elem.handle, timeout=0.5)
                menu_dlg = menu_app.window(handle=elem.handle)
                item = menu_dlg.child_window(auto_id=SETTINGS_MENU_ITEM_AUTO_ID, control_type="MenuItem")
                if item.exists(timeout=0.3):
                    item.wait("visible", timeout=1)
                    item.click_input()
                    print(f"[OK] 已点击菜单项 (auto_id={SETTINGS_MENU_ITEM_AUTO_ID})")
                    return True
            except Exception:
                continue
        time.sleep(0.15)
    return False


def switch_to_settings_panel(dlg, panel_name: str = PANEL_NAME) -> bool:
    """在设置对话框中切换到指定标签页（左侧 ListItem 导航菜单）。"""
    try:
        nav_list = dlg.child_window(auto_id="2210", control_type="List")
        nav_list.wait("ready", timeout=5)
        nav_list.set_focus()

        item = nav_list.child_window(title=panel_name, control_type="ListItem")
        item.wait("visible", timeout=3)
        item.click_input()
        time.sleep(0.5)
        print(f"[OK] 已切换到'{panel_name}'面板")
        return True
    except Exception as e:
        print(f"[WARN] 切换到'{panel_name}'面板失败: {e}")
        return False


# ============ Win32 消息加速（与 1_委托设置.py 一致）============
def _win32_user32():
    """返回已配置好参数类型的 user32 句柄，用于直接向 Win32 控件发消息。"""
    from ctypes import wintypes
    user32 = ctypes.windll.user32
    user32.SendMessageW.argtypes = [
        wintypes.HWND, wintypes.UINT, wintypes.WPARAM, ctypes.c_void_p
    ]
    user32.SendMessageW.restype = wintypes.LPARAM
    return user32


# ---- Win32 消息常量（来自 win32con）----
WM_GETTEXT = win32con.WM_GETTEXT
WM_SETTEXT = win32con.WM_SETTEXT
BM_GETCHECK = win32con.BM_GETCHECK
BM_SETCHECK = win32con.BM_SETCHECK
BM_CLICK = win32con.BM_CLICK
BST_CHECKED = win32con.BST_CHECKED
BST_UNCHECKED = win32con.BST_UNCHECKED
CB_GETCOUNT = win32con.CB_GETCOUNT
CB_GETCURSEL = win32con.CB_GETCURSEL
CB_GETLBTEXT = win32con.CB_GETLBTEXT


def _get_control_hwnd(dlg, auto_id):
    """用 win32gui 按控件 ID (DlgCtrlID) 查找子窗口句柄（最快，不触发 UIA 树遍历）。"""
    try:
        target = int(auto_id)
    except (TypeError, ValueError):
        return None
    found = []
    def _cb(hwnd, _):
        try:
            if win32gui.GetDlgCtrlID(hwnd) == target:
                found.append(hwnd)
        except Exception:
            pass
    try:
        win32gui.EnumChildWindows(dlg.handle, _cb, None)
    except Exception:
        pass
    return found[0] if found else None


def _get_window_text(hwnd, maxlen=256):
    """通过 WM_GETTEXT 读取窗口/控件文本（不展开、不遍历）。"""
    user32 = _win32_user32()
    buf = ctypes.create_unicode_buffer(maxlen)
    user32.SendMessageW(hwnd, WM_GETTEXT, maxlen, ctypes.addressof(buf))
    return buf.value or ""


def _get_combo_hwnd(dlg, auto_id):
    """获取组合框的原生窗口句柄（直接复用通用定位）。"""
    return _get_control_hwnd(dlg, auto_id)


def _read_combobox_items_win32(hwnd):
    """通过 CB_GETCOUNT / CB_GETLBTEXT 直接读取 Win32 组合框的候选项（不展开）。"""
    user32 = _win32_user32()
    count = user32.SendMessageW(hwnd, CB_GETCOUNT, 0, 0)
    if count is None or count <= 0:
        return None
    buf = ctypes.create_unicode_buffer(256)
    items = []
    seen = set()
    for i in range(count):
        user32.SendMessageW(hwnd, CB_GETLBTEXT, i, ctypes.addressof(buf))
        txt = buf.value.strip()
        if txt and txt not in seen:
            seen.add(txt)
            items.append(txt)
    return items if items else None


def get_checkbox_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取复选框的选中状态（Win32 BM_GETCHECK，速度快）。

    Returns:
        True=已勾选, False=未勾选, None=找不到控件
    """
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            state = _win32_user32().SendMessageW(hwnd, BM_GETCHECK, 0, 0)
            return bool(state & 1)
    except Exception as e:
        print(f"  [WARN] win32 读取复选框(auto_id={auto_id})失败，降级到 UIA: {e}")
    # 降级：UIA
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        cb.wait("ready", timeout=2)
        return bool(cb.get_toggle_state())
    except Exception as e:
        print(f"  [WARN] 获取复选框(auto_id={auto_id})失败: {e}")
        return None


def click_checkbox_by_id(dlg, auto_id: str):
    """通过 auto_id 点击复选框（切换其勾选状态，Win32 BM_CLICK）。"""
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            _win32_user32().SendMessageW(hwnd, BM_CLICK, 0, 0)
            print(f"[OK] 已点击复选框(auto_id={auto_id})(win32)")
            return
    except Exception as e:
        print(f"  [WARN] win32 点击复选框(auto_id={auto_id})失败，降级到 UIA: {e}")
    # 降级：UIA
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        cb.wait("ready", timeout=3)
        cb.click_input()
        print(f"[OK] 已点击复选框(auto_id={auto_id})")
    except Exception as e:
        print(f"  [WARN] 点击复选框(auto_id={auto_id})失败: {e}")


def get_combobox_selection_by_id(dlg, auto_id: str) -> Optional[str]:
    """通过 auto_id 获取下拉框当前选择的文本（Win32 CB_GETCURSEL/CB_GETLBTEXT）。"""
    try:
        hwnd = _get_combo_hwnd(dlg, auto_id)
        if hwnd is not None:
            user32 = _win32_user32()
            sel = user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0)
            if sel is not None and sel >= 0:
                buf = ctypes.create_unicode_buffer(256)
                user32.SendMessageW(hwnd, CB_GETLBTEXT, sel, ctypes.addressof(buf))
                return buf.value.strip() or None
    except Exception as e:
        print(f"  [WARN] win32 读取下拉选择失败，降级到 UIA: {e}")
    # 降级：UIA
    try:
        combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox")
        combo.wait("ready", timeout=2)
        selected_text = combo.selected_text()
        return selected_text.strip() if selected_text else None
    except Exception as e:
        print(f"  [WARN] 获取下拉框(auto_id={auto_id})失败: {e}")
        return None


# 左侧导航菜单项（读取下拉候选项时需排除）
_NAV_MENU_ITEMS = {
    "委托设置", "期权设置", "自动拆单设置", "自动追单设置",
    "快捷设置", "键盘下单设置", "价格提醒设置",
}


def get_combobox_items_by_id(dlg, auto_id: str) -> Optional[List[str]]:
    """读取下拉框全部候选项（Win32 CB_GETCOUNT/CB_GETLBTEXT 快速路径，不展开）。"""
    try:
        hwnd = _get_combo_hwnd(dlg, auto_id)
        if hwnd is not None:
            items = _read_combobox_items_win32(hwnd)
            if items:
                print(f"  [INFO] 已读取到 {len(items)} 个下拉候选项(win32快速路径)")
                return items
    except Exception as e:
        print(f"  [WARN] win32 读取下拉候选项失败，降级到 UIA: {e}")
    # 降级：UIA 展开 + ListItem
    try:
        combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox", found_index=0)
        combo.wait("ready", timeout=2)

        combo.expand()
        print(f"  [INFO] 已点开下拉框(auto_id={auto_id})，读取候选项...")
        time.sleep(0.6)

        items: List[str] = []
        seen: set = set()
        try:
            for li in combo.descendants(control_type="ListItem"):
                txt = (li.window_text() or "").strip()
                if not txt or txt in _NAV_MENU_ITEMS:
                    continue
                if txt not in seen:
                    seen.add(txt)
                    items.append(txt)
        except Exception:
            items = []

        if not items:
            try:
                texts = combo.item_texts()
                items = [t.strip() for t in texts if t and t.strip()]
            except Exception:
                items = []

        combo.collapse()
        print(f"  [INFO] 已关闭下拉框(auto_id={auto_id})")
        time.sleep(0.4)

        return items if items else None
    except Exception as e:
        print(f"  [WARN] 获取下拉框候选项(auto_id={auto_id})失败: {e}")
        return None


def get_edit_value_by_id(dlg, auto_id: str, as_number: bool = True) -> Optional[Any]:
    """通过 auto_id 获取数值输入框(Edit)的值（Win32 WM_GETTEXT，速度快）。

    as_number=True 时尝试返回 int/float；为字符串值则返回 str。
    """
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            text = _get_window_text(hwnd).strip().replace(",", "")
            if not as_number:
                return text
            if text == "" or text == "-":
                return None
            try:
                return int(text)
            except ValueError:
                return float(text)
    except Exception as e:
        print(f"  [WARN] win32 读取数值框(auto_id={auto_id})失败，降级到 UIA: {e}")
    # 降级：UIA
    try:
        edit = dlg.child_window(auto_id=auto_id, control_type="Edit")
        edit.wait("exists", timeout=5)
        for _ in range(2):
            try:
                edit.wait("ready", timeout=2)
            except Exception:
                pass
            try:
                text = edit.get_value().strip()
                text = text.replace(",", "")
                if not as_number:
                    return text
                if text == "" or text == "-":
                    return None
                try:
                    return int(text)
                except ValueError:
                    return float(text)
            except Exception:
                time.sleep(0.5)
        return None
    except Exception as e:
        print(f"  [WARN] 获取数值框(auto_id={auto_id})失败: {e}")
        return None


def take_screenshot(dlg, save_path: str):
    """对“交易系统设置”对话框整体截图保存。"""
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        rect = dlg.rectangle()
        left, top = int(rect.left), int(rect.top)
        width = int(rect.right - rect.left)
        height = int(rect.bottom - rect.top)
        if width <= 0 or height <= 0:
            raise ValueError("窗口矩形无效，无法截图")

        try:
            from mss import MSS
            from mss.tools import to_png
            with MSS() as sct:
                monitor = {"top": top, "left": left, "width": width, "height": height}
                img = sct.grab(monitor)
                to_png(img.rgb, img.size, output=save_path)
            print(f"[OK] 截图已保存(mss): {save_path}")
            return
        except Exception as e:
            print(f"  [WARN] mss 截图失败: {e}")
    except Exception as e:
        print(f"[WARN] 截图失败: {e}")


def test_auto_order_followup(dlg, result: SettingsTestResult):
    """测试自动追单设置各项参数。

    流程：
        1. 检测"启用自动追单"初始状态
        2. 若未勾选 → 点击启用以暴露下方参数，检查完后再点击关闭恢复
        3. 在启用状态下逐项检查下方参数（使用 / 追价 / 间隔 / 重复 / 自动撤单）
    """
    print("\n--- 自动追单设置检查 ---")

    # 1. 检测"启用自动追单"初始状态
    initial_enabled = get_checkbox_state_by_id(dlg, AUTO_ID["启用自动追单"])
    result.add_result("启用自动追单_初始状态", initial_enabled,
                      STANDARD_VALUES["启用自动追单"])

    # 2. 若未启用，点击启用以暴露下方参数
    need_restore = False
    if not initial_enabled:
        print("  [INFO] '启用自动追单'未勾选，点击启用以暴露下方参数...")
        click_checkbox_by_id(dlg, AUTO_ID["启用自动追单"])
        need_restore = True
        time.sleep(0.6)

        now_enabled = get_checkbox_state_by_id(dlg, AUTO_ID["启用自动追单"])
        result.add_result("启用自动追单_启用后", now_enabled, True)
    else:
        print("  [INFO] '启用自动追单'已勾选，直接检查下方参数")

    # 3. 检查下方参数（启用状态下）

    # 3.1 使用（下拉框）
    use = get_combobox_selection_by_id(dlg, AUTO_ID["使用"])
    result.add_result("使用", use or "(未知)", STANDARD_VALUES["使用"])

    use_items = get_combobox_items_by_id(dlg, AUTO_ID["使用"])
    use_actual = "、".join(use_items) if use_items else "(无法获取)"
    use_expect = "、".join(STANDARD_VALUES["使用_选项列表"])
    result.add_result("使用_选项列表", use_actual, use_expect)

    # 3.2 追价，调整（下拉框）
    adjust = get_combobox_selection_by_id(dlg, AUTO_ID["追价，调整"])
    try:
        adjust_val = int(adjust) if adjust is not None and adjust.isdigit() else adjust
    except Exception:
        adjust_val = adjust
    result.add_result("追价，调整", adjust_val, STANDARD_VALUES["追价，调整"])

    adjust_items = get_combobox_items_by_id(dlg, AUTO_ID["追价，调整"])
    adjust_actual = "、".join(adjust_items) if adjust_items else "(无法获取)"
    adjust_expect = "、".join(STANDARD_VALUES["追价，调整_选项列表"])
    result.add_result("追价，调整_选项列表", adjust_actual, adjust_expect)

    # 3.3 自动追单时间间隔（下拉框）
    interval = get_combobox_selection_by_id(dlg, AUTO_ID["自动追单时间间隔"])
    try:
        interval_val = int(interval) if interval is not None and interval.isdigit() else interval
    except Exception:
        interval_val = interval
    result.add_result("自动追单时间间隔", interval_val, STANDARD_VALUES["自动追单时间间隔"])

    interval_items = get_combobox_items_by_id(dlg, AUTO_ID["自动追单时间间隔"])
    interval_actual = "、".join(interval_items) if interval_items else "(无法获取)"
    interval_expect = "、".join(STANDARD_VALUES["自动追单时间间隔_选项列表"])
    result.add_result("自动追单时间间隔_选项列表", interval_actual, interval_expect)

    # 3.4 秒，最多重复（Edit）
    repeat = get_edit_value_by_id(dlg, AUTO_ID["秒，最多重复"])
    result.add_result("秒，最多重复", repeat if repeat is not None else "(未知)",
                      STANDARD_VALUES["秒，最多重复"])

    # 3.5 未完成则自动撤单（CheckBox）
    cancel = get_checkbox_state_by_id(dlg, AUTO_ID["未完成则自动撤单"])
    result.add_result("未完成则自动撤单", cancel, STANDARD_VALUES["未完成则自动撤单"])

    # 4. 若之前未启用，检查完成后恢复关闭状态
    if need_restore:
        print("  [INFO] 检查完成，恢复'启用自动追单'为关闭状态...")
        click_checkbox_by_id(dlg, AUTO_ID["启用自动追单"])
        time.sleep(0.4)
        restored = get_checkbox_state_by_id(dlg, AUTO_ID["启用自动追单"])
        result.add_result("启用自动追单_恢复后", restored,
                          STANDARD_VALUES["启用自动追单"])


def explore_dialog_controls(dlg):
    """探索对话框内所有控件（调试用）。"""
    print("\n=== 控件探索 ===")
    try:
        descendants = dlg.descendants()
        for ctrl in descendants:
            try:
                ctrl_type = ctrl.element_info.control_type
                ctrl_name = ctrl.element_info.name or ""
                try:
                    ctrl_id = ctrl.element_info.automation_id or ""
                except Exception:
                    ctrl_id = ""

                if ctrl_type in ("CheckBox", "ComboBox", "Edit", "RadioButton",
                                  "Button", "Spinner", "ListItem"):
                    extra = ""
                    if ctrl_type == "CheckBox":
                        try:
                            extra = f" [状态={ctrl.get_toggle_state()}]"
                        except Exception:
                            pass
                    elif ctrl_type in ("Edit", "Spinner"):
                        try:
                            val = ctrl.get_value()
                            extra = f" [值={val}]"
                        except Exception:
                            pass
                    elif ctrl_type == "ComboBox":
                        try:
                            sel = ctrl.selected_text()
                            extra = f" [选中={sel}]"
                        except Exception:
                            pass
                    print(f"  [{ctrl_type}] name='{ctrl_name}' id='{ctrl_id}'{extra}")
            except Exception as e:
                print(f"  [?] 获取信息失败: {e}")
    except Exception as e:
        print(f"  探索失败: {e}")


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 自动追单设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult(PANEL_NAME)
    hwnd = None
    dlg = None

    try:
        # 1. 倒计时
        countdown(COUNTDOWN_SEC)

        # 2. 查找主窗口
        hwnd = find_window(WINDOW_KEYWORD)
        print(f"[OK] 已找到主窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)

        # 3. 自动打开设置对话框
        dlg = open_settings_dialog_compat(
            win, SETTINGS_BUTTON_AUTO_ID, SETTINGS_MENU_ITEM_AUTO_ID, SETTINGS_DIALOG_TITLE
        )
        dlg.wait("ready", timeout=10)

        # 4. 切换到自动追单设置面板
        if not switch_settings_panel_compat(dlg, PANEL_NAME):
            print("[错误] 无法切换到自动追单设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉）
        #print("\n正在进行控件探索...")
        #explore_dialog_controls(dlg)

        # 6. 执行测试
        test_auto_order_followup(dlg, result)

        # 7. 截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"自动追单设置_{timestamp}.png")
        take_screenshot(dlg, screenshot_path)

        # 8. 输出结果
        result.print_summary()

        # 9. 保存报告
        report_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"自动追单设置测试报告_{timestamp}.txt")
        result.to_file(report_path)

        print(f"\n=== 测试完成 ===")

    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        # 无论正常完成、异常还是用户中断，都执行安全收尾。
        if dlg is not None:
            keep_open = os.environ.get("GUI_NEXT_CATEGORY", "") == "交易系统设置"
            close_ok = close_settings_dialog(
                dlg, keep_open=keep_open, main_hwnd=hwnd
            )
            if not close_ok:
                print("[WARN] 交易系统设置窗口未正常关闭，请确认后再执行后续非设置类任务")


if __name__ == "__main__":
    main()
