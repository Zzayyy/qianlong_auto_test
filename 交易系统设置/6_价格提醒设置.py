# -*- coding: utf-8 -*-
"""
交易系统设置 - 价格提醒设置界面自动化测试
==========================================
功能：
    打开"交易系统设置"对话框，进入"价格提醒设置"标签页，
    逐项读取每个参数的当前值（复选框状态、倍数输入框数值等），
    与标准值（恢复默认后的参数）比对，记录差异并截图。

界面元素（价格提醒设置，来自交易系统设置_价格提醒设置.txt）：
    一、合约委托价格超过限定价格提醒
        - 买开、买平、备平委托价格高于最新价格的（CheckBox, auto_id=2065）
          + 倍数输入框（Edit, auto_id=2150，说明"倍提醒（参数设置大于1）"）
        - 卖开、卖平、备开委托价格低于最新价格的（CheckBox, auto_id=2068）
          + 倍数输入框（Edit, auto_id=2179，说明"倍提醒（参数设置大于0小于1）"）

比对规则：
    - 复选框勾选状态直接与标准值比对
    - 若复选框未勾选，则对应倍数输入框视为"未启用"（不计入差异）
    - 若复选框已勾选，则读取倍数并比对

使用方法：
    1. 打开钱龙旗舰版，登录交易账号
    2. 运行本脚本

依赖：
    pip install pywinauto pillow mss
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
PANEL_NAME = "价格提醒设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对。
# 优先从 交易系统设置/标准/<客户端>/价格提醒设置.json 读取（可自定义/抓取覆盖）；
# 找不到时使用下方内嵌兜底（与 qianlong 默认标准一致），保证离线不崩。
# 已确认（2026-07-08 实测）：
#   默认两个复选框均未勾选（未启用），倍数输入框为空（空白）。
#   因此复选框标准值 = False，倍数在"未启用"时记为空、不计入差异。
#   倍数对比仅在复选框已勾选时才进行；若你启用后需要校验具体倍数，
#   请把启用状态下的默认倍数回填到下列"倍数"项（参数需 >1 / 0~1）。
DEFAULT_STANDARD_VALUES = {
    # 一、合约委托价格超过限定价格提醒
    "买开买平备平_委托价格高于最新价格_勾选": False,   # 已确认默认未勾选
    "买开买平备平_委托价格高于最新价格_倍数": None,   # 未启用时空白；启用后按实际默认倍数回填
    "卖开卖平备开_委托价格低于最新价格_勾选": False,   # 已确认默认未勾选
    "卖开卖平备开_委托价格低于最新价格_倍数": None,   # 未启用时空白；启用后按实际默认倍数回填
}

# 当前客户端（GUI 启动时由 GUI_CLIENT_ID 环境变量注入；空则用内嵌兜底）
CLIENT_ID = os.environ.get("GUI_CLIENT_ID", "") or ""
STANDARD_VALUES = load_standard(PANEL_NAME, CLIENT_ID, DEFAULT_STANDARD_VALUES)

# 控件 auto_id 映射（来自交易系统设置_价格提醒设置.txt 抓取）
AUTO_ID = {
    # 一、合约委托价格超过限定价格提醒
    "买开买平备平_委托价格高于最新价格_勾选": "2065",
    "买开买平备平_委托价格高于最新价格_倍数": "2150",
    "卖开卖平备开_委托价格低于最新价格_勾选": "2068",
    "卖开卖平备开_委托价格低于最新价格_倍数": "2179",
}

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "价格提醒设置"
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


def test_price_reminder(dlg, result: SettingsTestResult):
    """测试价格提醒设置各项参数。

    界面结构：
        一、合约委托价格超过限定价格提醒
            - 买开、买平、备平委托价格高于最新价格的（CheckBox）
              + 倍数输入框（Edit，说明"倍提醒（参数设置大于1）"）
            - 卖开、卖平、备开委托价格低于最新价格的（CheckBox）
              + 倍数输入框（Edit，说明"倍提醒（参数设置大于0小于1）"）

    比对规则：
        - 复选框勾选状态直接与标准值比对
        - 复选框未勾选 → 对应倍数视为"未启用"（不计入差异）
        - 复选框已勾选 → 读取倍数并比对
    """
    print("\n--- 价格提醒设置检查 ---")

    # 1. 买开、买平、备平委托价格高于最新价格的
    buy_checked = get_checkbox_state_by_id(dlg, AUTO_ID["买开买平备平_委托价格高于最新价格_勾选"])
    result.add_result("买开买平备平_委托价格高于最新价格_勾选", buy_checked,
                      STANDARD_VALUES["买开买平备平_委托价格高于最新价格_勾选"])

    if buy_checked:
        buy_mult = get_edit_value_by_id(dlg, AUTO_ID["买开买平备平_委托价格高于最新价格_倍数"])
        result.add_result("买开买平备平_委托价格高于最新价格_倍数",
                          buy_mult if buy_mult is not None else "(未知)",
                          STANDARD_VALUES["买开买平备平_委托价格高于最新价格_倍数"])
    else:
        result.add_not_enabled("买开买平备平_委托价格高于最新价格_倍数")

    # 2. 卖开、卖平、备开委托价格低于最新价格的
    sell_checked = get_checkbox_state_by_id(dlg, AUTO_ID["卖开卖平备开_委托价格低于最新价格_勾选"])
    result.add_result("卖开卖平备开_委托价格低于最新价格_勾选", sell_checked,
                      STANDARD_VALUES["卖开卖平备开_委托价格低于最新价格_勾选"])

    if sell_checked:
        sell_mult = get_edit_value_by_id(dlg, AUTO_ID["卖开卖平备开_委托价格低于最新价格_倍数"])
        result.add_result("卖开卖平备开_委托价格低于最新价格_倍数",
                          sell_mult if sell_mult is not None else "(未知)",
                          STANDARD_VALUES["卖开卖平备开_委托价格低于最新价格_倍数"])
    else:
        result.add_not_enabled("卖开卖平备开_委托价格低于最新价格_倍数")


def explore_dialog_controls(dlg):
    """探索对话框内所有控件（调试用，可用于核对控件 auto_id 与默认值）。"""
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


def collect_current_settings(dlg) -> dict:
    """读取当前面板全部设置值，返回与 STANDARD_VALUES 同构的字典。

    供“抓取自定义标准”脚本把当前客户端界面值采集为新的比对标准。
    逻辑与 test_price_reminder 一致（复选框未勾选时对应倍数记为空 None），
    但只返回字典、不写报告、不改任何设置（不点击、不录入倍数）。
    """
    data: dict = {}

    # 1. 买开、买平、备平委托价格高于最新价格的
    buy_checked = get_checkbox_state_by_id(dlg, AUTO_ID["买开买平备平_委托价格高于最新价格_勾选"])
    data["买开买平备平_委托价格高于最新价格_勾选"] = bool(buy_checked) if buy_checked is not None else False
    if buy_checked:
        buy_mult = get_edit_value_by_id(dlg, AUTO_ID["买开买平备平_委托价格高于最新价格_倍数"])
        data["买开买平备平_委托价格高于最新价格_倍数"] = buy_mult
    else:
        data["买开买平备平_委托价格高于最新价格_倍数"] = None

    # 2. 卖开、卖平、备开委托价格低于最新价格的
    sell_checked = get_checkbox_state_by_id(dlg, AUTO_ID["卖开卖平备开_委托价格低于最新价格_勾选"])
    data["卖开卖平备开_委托价格低于最新价格_勾选"] = bool(sell_checked) if sell_checked is not None else False
    if sell_checked:
        sell_mult = get_edit_value_by_id(dlg, AUTO_ID["卖开卖平备开_委托价格低于最新价格_倍数"])
        data["卖开卖平备开_委托价格低于最新价格_倍数"] = sell_mult
    else:
        data["卖开卖平备开_委托价格低于最新价格_倍数"] = None

    return data


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 价格提醒设置自动化测试")
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

        # 4. 切换到价格提醒设置面板
        if not switch_settings_panel_compat(dlg, PANEL_NAME):
            print("[错误] 无法切换到价格提醒设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉；便于核对 auto_id 与真实默认值）
        #print("\n正在进行控件探索...")
        #explore_dialog_controls(dlg)

        # 6. 执行测试
        test_price_reminder(dlg, result)

        # 7. 截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"价格提醒设置_{timestamp}.png")
        take_screenshot(dlg, screenshot_path)

        # 8. 输出结果
        result.print_summary()

        # 9. 保存报告
        report_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"价格提醒设置测试报告_{timestamp}.txt")
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
