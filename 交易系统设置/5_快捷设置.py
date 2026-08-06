# -*- coding: utf-8 -*-
"""
交易系统设置 - 快捷设置界面自动化测试
====================================
功能：
    打开"交易系统设置"对话框，进入"快捷设置"标签页，
    逐项读取每个参数的当前值（复选框状态、数值输入框等），
    与标准值（恢复默认后的参数）比对，记录差异并截图。

界面元素（快捷设置）：
    一、鼠标快捷输入
        - 鼠标快捷输入（CheckBox，默认勾选）
        - 数量 1~5（Edit，默认 1/2/3/4/5）
        - 百分比% 1~5（Edit，默认 10/20/30/40/50）

    二、委托数量最小跳动
        - 委托数量最小跳动（CheckBox，默认未勾选）
        - 数值（Edit + Spinner，默认 1）

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
from typing import Any, Optional

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.settings import SettingsTestResult
from core.settings_standard import load_standard


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "快捷设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对。
# 优先从 交易系统设置/标准/<客户端>/快捷设置.json 读取（可自定义/抓取覆盖）；
# 找不到时使用下方内嵌兜底（与 qianlong 默认标准一致），保证离线不崩。
DEFAULT_STANDARD_VALUES = {
    # 一、鼠标快捷输入
    "鼠标快捷输入": True,
    "数量_1": 1,
    "数量_2": 2,
    "数量_3": 3,
    "数量_4": 4,
    "数量_5": 5,
    "百分比_1": 10,
    "百分比_2": 20,
    "百分比_3": 30,
    "百分比_4": 40,
    "百分比_5": 50,

    # 二、委托数量最小跳动
    "委托数量最小跳动": False,
    "委托数量最小跳动_数值": 1,
}

# 当前客户端（GUI 启动时由 GUI_CLIENT_ID 环境变量注入；空则用内嵌兜底）
CLIENT_ID = os.environ.get("GUI_CLIENT_ID", "") or ""
STANDARD_VALUES = load_standard(PANEL_NAME, CLIENT_ID, DEFAULT_STANDARD_VALUES)

# 控件 auto_id 映射（来自交易系统设置_快捷设置.txt 抓取）
AUTO_ID = {
    # 一、鼠标快捷输入
    "鼠标快捷输入": "2067",
    "数量_1": "2133",
    "数量_2": "2135",
    "数量_3": "2136",
    "数量_4": "2137",
    "数量_5": "2138",
    "百分比_1": "2139",
    "百分比_2": "2140",
    "百分比_3": "2141",
    "百分比_4": "2142",
    "百分比_5": "2134",

    # 二、委托数量最小跳动
    "委托数量最小跳动": "2066",
    "委托数量最小跳动_数值": "2161",
}

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "快捷设置"
COUNTDOWN_SEC = int(os.environ.get("GUI_COUNTDOWN", "3"))  # 倒计时秒数(GUI可配)
# ========================================================














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
BM_GETCHECK = win32con.BM_GETCHECK
BM_SETCHECK = win32con.BM_SETCHECK
BM_CLICK = win32con.BM_CLICK
BST_CHECKED = win32con.BST_CHECKED
BST_UNCHECKED = win32con.BST_UNCHECKED


def _get_control_hwnd(dlg, auto_id: str):
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


def _get_window_text(hwnd: int, maxlen: int = 256) -> str:
    """通过 WM_GETTEXT 读取窗口/控件文本（不展开、不遍历）。"""
    user32 = _win32_user32()
    buf = ctypes.create_unicode_buffer(maxlen)
    user32.SendMessageW(hwnd, WM_GETTEXT, maxlen, ctypes.addressof(buf))
    return buf.value or ""


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
            # 优先转 int，否则转 float
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
                # 优先转 int，否则转 float
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
    """对"交易系统设置"对话框整体截图保存。"""
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
                monitor = {"top": top, "left": left,
                           "width": width, "height": height}
                img = sct.grab(monitor)
                to_png(img.rgb, img.size, output=save_path)
            print(f"[OK] 截图已保存(mss): {save_path}")
            return
        except Exception as e:
            print(f"  [WARN] mss 截图失败: {e}")
    except Exception as e:
        print(f"[WARN] 截图失败: {e}")


def test_mouse_shortcut(dlg, result: SettingsTestResult):
    """测试一、鼠标快捷输入"""
    print("\n--- [1/2] 鼠标快捷输入 ---")

    mouse = get_checkbox_state_by_id(dlg, AUTO_ID["鼠标快捷输入"])
    result.add_result("鼠标快捷输入", mouse, STANDARD_VALUES["鼠标快捷输入"])

    # 数量 1~5
    for i in range(1, 6):
        key = f"数量_{i}"
        val = get_edit_value_by_id(dlg, AUTO_ID[key])
        result.add_result(key, val if val is not None else "(未知)", STANDARD_VALUES[key])

    # 百分比 1~5
    for i in range(1, 6):
        key = f"百分比_{i}"
        val = get_edit_value_by_id(dlg, AUTO_ID[key])
        result.add_result(key, val if val is not None else "(未知)", STANDARD_VALUES[key])


def test_min_tick(dlg, result: SettingsTestResult):
    """测试二、委托数量最小跳动"""
    print("\n--- [2/2] 委托数量最小跳动 ---")

    tick = get_checkbox_state_by_id(dlg, AUTO_ID["委托数量最小跳动"])
    result.add_result("委托数量最小跳动", tick, STANDARD_VALUES["委托数量最小跳动"])

    if tick:
        tick_val = get_edit_value_by_id(dlg, AUTO_ID["委托数量最小跳动_数值"])
        result.add_result("委托数量最小跳动_数值",
                          tick_val if tick_val is not None else 0,
                          STANDARD_VALUES["委托数量最小跳动_数值"])
    else:
        result.add_not_enabled("委托数量最小跳动_数值")


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
                        except:
                            pass
                    elif ctrl_type in ("Edit", "Spinner"):
                        try:
                            val = ctrl.get_value()
                            extra = f" [值={val}]"
                        except:
                            pass
                    elif ctrl_type == "ComboBox":
                        try:
                            sel = ctrl.selected_text()
                            extra = f" [选中={sel}]"
                        except:
                            pass

                    print(f"  [{ctrl_type}] name='{ctrl_name}' id='{ctrl_id}'{extra}")
            except Exception as e:
                print(f"  [?] 获取信息失败: {e}")
    except Exception as e:
        print(f"  探索失败: {e}")


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


def collect_current_settings(dlg) -> dict:
    """读取当前面板全部设置值，返回与 STANDARD_VALUES 同构的字典。

    供“抓取自定义标准”脚本把当前客户端界面值采集为新的比对标准。
    逻辑与下方 test_* 函数一致（委托数量最小跳动未启用时临时点击启用以读取
    下方数值，读取完再恢复），但只返回字典、不写报告、不改任何设置。
    """
    data: dict = {}

    # 一、鼠标快捷输入
    mouse = get_checkbox_state_by_id(dlg, AUTO_ID["鼠标快捷输入"])
    data["鼠标快捷输入"] = bool(mouse) if mouse is not None else False
    for i in range(1, 6):
        val = get_edit_value_by_id(dlg, AUTO_ID[f"数量_{i}"])
        data[f"数量_{i}"] = val if val is not None else 0
    for i in range(1, 6):
        val = get_edit_value_by_id(dlg, AUTO_ID[f"百分比_{i}"])
        data[f"百分比_{i}"] = val if val is not None else 0

    # 二、委托数量最小跳动（未启用时临时启用以读取下方数值，再恢复）
    tick = get_checkbox_state_by_id(dlg, AUTO_ID["委托数量最小跳动"])
    data["委托数量最小跳动"] = bool(tick) if tick is not None else False
    need_restore = False
    if not tick:
        print("  [INFO] '委托数量最小跳动'未勾选，点击启用以暴露下方数值...")
        click_checkbox_by_id(dlg, AUTO_ID["委托数量最小跳动"])
        need_restore = True
        time.sleep(0.6)
    val = get_edit_value_by_id(dlg, AUTO_ID["委托数量最小跳动_数值"])
    data["委托数量最小跳动_数值"] = val if val is not None else 0
    if need_restore:
        print("  [INFO] 检查完成，恢复'委托数量最小跳动'为未启用状态...")
        click_checkbox_by_id(dlg, AUTO_ID["委托数量最小跳动"])
        time.sleep(0.4)

    return data


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 快捷设置自动化测试")
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

        # 4. 切换到快捷设置面板
        if not switch_settings_panel_compat(dlg, PANEL_NAME):
            print("[错误] 无法切换到快捷设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉）
        #print("\n正在进行控件探索...")
        #explore_dialog_controls(dlg)

        # 6. 执行各项测试
        test_mouse_shortcut(dlg, result)
        test_min_tick(dlg, result)

        # 7. 截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"快捷设置_{timestamp}.png")
        take_screenshot(dlg, screenshot_path)

        # 8. 输出结果
        result.print_summary()

        # 9. 保存报告
        report_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"快捷设置测试报告_{timestamp}.txt")
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
