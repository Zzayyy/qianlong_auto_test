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
from typing import Any, Dict, List, Optional

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image
from pywinauto import Application, findwindows
from core.window import find_window, activate_window, countdown, close_settings_dialog


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "价格提醒设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对
# 已确认（2026-07-08 实测）：
#   默认两个复选框均未勾选（未启用），倍数输入框为空（空白）。
#   因此复选框标准值 = False，倍数在"未启用"时记为空、不计入差异。
#   倍数对比仅在复选框已勾选时才进行；若你启用后需要校验具体倍数，
#   请把启用状态下的默认倍数回填到下列"倍数"项（参数需 >1 / 0~1）。
STANDARD_VALUES = {
    # 一、合约委托价格超过限定价格提醒
    "买开买平备平_委托价格高于最新价格_勾选": False,   # 已确认默认未勾选
    "买开买平备平_委托价格高于最新价格_倍数": None,   # 未启用时空白；启用后按实际默认倍数回填
    "卖开卖平备开_委托价格低于最新价格_勾选": False,   # 已确认默认未勾选
    "卖开卖平备开_委托价格低于最新价格_倍数": None,   # 未启用时空白；启用后按实际默认倍数回填
}

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
COUNTDOWN_SEC = 3  # 倒计时秒数
# ========================================================


class SettingsTestResult:
    """测试结果收集器"""

    def __init__(self):
        self.results: List[Dict[str, Any]] = []
        self.differences: List[Dict[str, Any]] = []
        self.not_enabled: int = 0

    def add_result(self, name: str, actual_value: Any, expected_value: Any):
        """添加一条测试结果"""
        matched = actual_value == expected_value
        result = {
            "名称": name,
            "期望值": expected_value,
            "实际值": actual_value,
            "是否一致": "✓" if matched else "✗ 差异"
        }
        self.results.append(result)
        if not matched:
            self.differences.append(result)

    def add_not_enabled(self, name: str, detail: str = "(未启用)"):
        """记录一个因上级开关未勾选而未生效的参数。"""
        self.results.append({
            "名称": name,
            "期望值": "—",
            "实际值": detail,
            "是否一致": "○ 未启用"
        })
        self.not_enabled += 1

    def print_summary(self):
        """打印测试摘要"""
        total = len(self.results)
        diff = len(self.differences)
        passed = total - diff - self.not_enabled
        print(f"\n{'='*60}")
        print(f"测试结果汇总")
        print(f"{'='*60}")
        print(f"总项目数: {total}")
        print(f"通过: {passed}")
        print(f"差异: {diff}")
        print(f"未启用(默认预期，不计差异): {self.not_enabled}")

        if self.results:
            print(f"\n{'名称':<35} {'期望值':<25} {'实际值':<25} {'状态'}")
            print(f"{'-'*100}")
            for r in self.results:
                name = str(r["名称"])[:33]
                exp = str(r["期望值"])[:23]
                act = str(r["实际值"])[:23]
                status = r["是否一致"]
                print(f"{name:<35} {exp:<25} {act:<25} {status}")

    def to_file(self, filepath: str):
        """将结果写入文件"""
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"价格提醒设置测试报告\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*60}\n\n")

            total = len(self.results)
            diff = len(self.differences)
            passed = total - diff - self.not_enabled
            f.write(f"总项目数: {total}\n")
            f.write(f"通过: {passed}\n")
            f.write(f"差异: {diff}\n")
            f.write(f"未启用(默认预期，不计差异): {self.not_enabled}\n\n")

            f.write(f"{'名称':<35} {'期望值':<25} {'实际值':<25} {'状态'}\n")
            f.write(f"{'-'*100}\n")
            for r in self.results:
                name = str(r["名称"])[:33]
                exp = str(r["期望值"])[:23]
                act = str(r["实际值"])[:23]
                status = r["是否一致"]
                f.write(f"{name:<35} {exp:<25} {act:<25} {status}\n")

            if self.differences:
                f.write(f"\n{'='*60}\n")
                f.write("差异详情:\n")
                for d in self.differences:
                    f.write(f"  - {d['名称']}: 期望={d['期望值']}, 实际={d['实际值']}\n")

        print(f"[OK] 测试报告已保存: {filepath}")


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


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 价格提醒设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult()
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
        dlg = open_settings_dialog(win)
        dlg.wait("ready", timeout=10)

        # 4. 切换到价格提醒设置面板
        if not switch_to_settings_panel(dlg, PANEL_NAME):
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
