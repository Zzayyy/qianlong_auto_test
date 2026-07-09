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
from datetime import datetime
from typing import Any, Dict, List, Optional

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from PIL import Image
from pywinauto import Application, findwindows
from core.window import find_window, activate_window, countdown


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "快捷设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对
STANDARD_VALUES = {
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
COUNTDOWN_SEC = 3  # 倒计时秒数
# ========================================================


class SettingsTestResult:
    """测试结果收集器"""

    def __init__(self):
        self.results: List[Dict[str, Any]] = []
        self.differences: List[Dict[str, Any]] = []
        self.not_enabled: int = 0          # 因上级开关未启用而未生效的参数数

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
        """记录一个因上级开关未启用而未生效的参数。

        属于默认预期，不计入差异，仅在报告中以"○ 未启用"提示。
        """
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
            f.write(f"快捷设置测试报告\n")
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
    """自动打开'交易系统设置'对话框。

    流程：
        1. 先检查是否已打开，已打开则直接连接
        2. 点击工具栏设置按钮 (auto_id=1008)，弹出上下文菜单
        3. 在弹出菜单中点击"交易系统设置"项 (auto_id=20025)
        4. 等待设置对话框出现（兼容顶级窗口与子窗口）
    """
    # ── 方案A: 先检查是否已打开 ──
    dlg = _find_existing_settings_dlg(win)
    if dlg is not None:
        return dlg

    # ── 方案B: 自动点击打开 ──
    print("\n正在通过工具栏按钮打开'交易系统设置'...")

    # 1. 点击设置按钮 (auto_id=1008)
    _click_settings_button(win)

    # 2. 等待弹出菜单出现，选择"交易系统设置"
    _click_context_menu_item(win)

    # 3. 等待设置对话框出现（先给一点渲染时间，再轮询）
    print("  等待设置对话框弹出...")
    time.sleep(1.0)
    end = time.time() + 10
    while time.time() < end:
        dlg = _find_existing_settings_dlg(win)
        if dlg is not None:
            return dlg
        time.sleep(0.3)

    # 降级：尝试键盘方向键选择
    print("[WARN] 菜单项未点击成功，尝试键盘降级方案...")
    time.sleep(2)
    dlg = _find_existing_settings_dlg(win)
    if dlg is not None:
        return dlg

    raise RuntimeError(
        f"无法打开'{SETTINGS_DIALOG_TITLE}'对话框，请确认:\n"
        f"  1. 钱龙软件已登录交易账号\n"
        f"  2. 工具栏设置按钮(auto_id={SETTINGS_BUTTON_AUTO_ID})可见"
    )


def _find_existing_settings_dlg(win=None) -> Optional[Any]:
    """查找已打开的设置对话框。"""
    # 1. 顶级窗口，直接按标题匹配
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

    # 2. #32770 对话框：标题仅存在于 TitleBar 的 Value 中
    try:
        for elem in findwindows.find_elements(top_level_only=True, class_name="#32770"):
            try:
                app = Application(backend="uia").connect(handle=elem.handle, timeout=1)
                dlg = app.window(handle=elem.handle)
                tb = dlg.child_window(control_type="TitleBar")
                if tb.exists(timeout=0.5):
                    val = ""
                    try:
                        legacy = tb.legacy_properties()
                        val = legacy.get("Value", "") or ""
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

    # 3. 主窗口的子窗口 / 后代窗口兜底
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
    """点击主窗口上的设置按钮 (auto_id=1008)。"""
    try:
        btn = win.child_window(
            auto_id=SETTINGS_BUTTON_AUTO_ID,
            control_type="Button"
        )
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
    """在弹出菜单中选择'交易系统设置' (auto_id=20025)。"""
    end = time.time() + 4
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                menu_app = Application(backend="uia").connect(handle=elem.handle, timeout=0.5)
                menu_dlg = menu_app.window(handle=elem.handle)
                item = menu_dlg.child_window(
                    auto_id=SETTINGS_MENU_ITEM_AUTO_ID,
                    control_type="MenuItem",
                )
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
    """在设置对话框中切换到指定标签页（左侧 ListItem 导航菜单）。

    控件结构（来自抓取文档）：
        ListBox - '' (auto_id="2210", control_type="List")
           ├── ListItem - '委托设置'
           ├── ListItem - '期权设置'
           └── ...
    """
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


def get_checkbox_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取复选框的选中状态。"""
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        cb.wait("ready", timeout=2)
        return bool(cb.get_toggle_state())
    except Exception as e:
        print(f"  [WARN] 获取复选框(auto_id={auto_id})失败: {e}")
        return None


def get_edit_value_by_id(dlg, auto_id: str, as_number: bool = True) -> Optional[Any]:
    """通过 auto_id 获取数值输入框(Edit)的值。

    as_number=True 时尝试返回 int/float；为字符串值则返回 str。
    """
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


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 快捷设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult()

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

        # 4. 切换到快捷设置面板
        if not switch_to_settings_panel(dlg, PANEL_NAME):
            print("[错误] 无法切换到快捷设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉）
        print("\n正在进行控件探索...")
        explore_dialog_controls(dlg)

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


if __name__ == "__main__":
    main()
