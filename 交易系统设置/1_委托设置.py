# -*- coding: utf-8 -*-
"""
交易系统设置 - 委托设置界面自动化测试
==========================================
功能：
    打开"交易系统设置"对话框，进入"委托设置"标签页，
    逐项读取每个参数的当前值（复选框状态、下拉框选择、数值等），
    与标准值（恢复默认后的参数）比对，记录差异并截图。

界面元素（委托设置）：
    一、股票买卖委托价格跟盘设置
        - 买入缺省价（CheckBox + ComboBox: 现价/最新价/...）
        - 卖出缺省价（CheckBox + ComboBox: 现价/最新价/...）

    二、大单自动分单设置
        - 股票拆单（CheckBox + SpinBox: 每单1000000股）
        - 基金拆单（CheckBox + SpinBox: 每单1000000份）

    三、委托数量设置
        - 股票买入自动填入数量（CheckBox + RadioButton组 + SpinBox）
            * 确定数量 / 全部数量 / 上一次交易数量
        - 股票卖出自动填入数量（同上）
        - 期权交易自动填入数量（CheckBox + SpinBox: 100张）
        - 期货交易自动填入数量（CheckBox + SpinBox: 100手）

    四、底部复选框
        - 静默委托下单模式
        - 显示期权下单成功提示
        - 显示期权宝软件风险揭示书
        - 委托成交时，发出提示音
        - 点击持仓/委托/成交记录联动行情

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
from typing import Any, Dict, List, Optional, Tuple

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pywinauto import Application, findwindows, mouse
from core.window import find_window, activate_window, countdown


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "委托设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对
STANDARD_VALUES = {
    # 股票买卖委托价格跟盘设置
    "买入缺省价_勾选": True,
    "买入缺省价_选项": "现价",
    "买入缺省价_下拉_选项列表": ["现价", "卖一", "卖二", "卖三", "卖四", "卖五"],
    "卖出缺省价_勾选": True,
    "卖出缺省价_选项": "现价",
    "卖出缺省价_下拉_选项列表": ["现价", "买一", "买二", "买三", "买四", "买五"],

    # 大单自动分单设置
    "股票拆单_勾选": True,
    "股票拆单_数值": 1000000,
    "基金拆单_勾选": False,
    "基金拆单_数值": 1000000,

    # 委托数量设置
    "股票买入自动填入数量_勾选": False,
    "股票买入_选项": "全部数量",       # RadioButton: 确定数量/全部数量/上一次交易数量
    "股票买入_数值": 100,
    "股票卖出自动填入数量_勾选": False,
    "股票卖出_选项": "全部数量",
    "股票卖出_数值": 100,
    "期权交易自动填入数量_勾选": False,
    "期权交易_数值": 100,
    "期货交易自动填入数量_勾选": False,
    "期货交易_数值": 100,

    # 底部复选框
    "静默委托下单模式": False,
    "显示期权下单成功提示": True,
    "显示期权宝软件风险揭示书": False,
    "委托成交时发出提示音": False,
    "点击持仓联动行情": True,
}

# 控件 auto_id 映射（来自交易系统设置_委托设置.txt 抓取）
AUTO_ID = {
    # 一、股票买卖委托价格跟盘设置
    "买入缺省价": "2026",
    "买入缺省价_下拉": "2080",
    "卖出缺省价": "2053",
    "卖出缺省价_下拉": "2105",

    # 二、大单自动分单设置
    "股票拆单": "2060",
    "股票拆单_数值": "2144",
    "基金拆单": "2062",
    "基金拆单_数值": "2147",

    # 三、委托数量设置
    "股票买入自动填入数量": "2025",
    "股票买入_确定数量": "2220",
    "股票买入_全部数量": "2221",
    "股票买入_上一次交易数量": "2222",
    "股票买入_数值": "2151",
    "股票卖出自动填入数量": "2052",
    "股票卖出_确定数量": "2223",
    "股票卖出_全部数量": "2224",
    "股票卖出_上一次交易数量": "2225",
    "股票卖出_数值": "2180",
    "期权交易自动填入数量": "2023",
    "期权交易_数值": "2159",
    "期货交易自动填入数量": "2022",
    "期货交易_数值": "2158",

    # 四、底部复选框
    "静默委托下单模式": "2044",
    "显示期权下单成功提示": "2056",
    "显示期权宝软件风险揭示书": "2047",
    "委托成交时发出提示音": "2030",
    "点击持仓联动行情": "2037",
}

# RadioButton auto_id -> 名称（买入/卖出共用同一组名称）
RADIO_NAMES = {
    "2220": "确定数量", "2221": "全部数量", "2222": "上一次交易数量",  # 股票买入
    "2223": "确定数量", "2224": "全部数量", "2225": "上一次交易数量",  # 股票卖出
}

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "委托设置"
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

        属于默认预期，不计入差异，仅在报告中以“○ 未启用”提示。
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
            print(f"\n{'名称':<30} {'期望值':<30} {'实际值':<30} {'状态'}")
            print(f"{'-'*100}")
            for r in self.results:
                name = r["名称"][:28]
                # 列宽30，超出时完整显示（下拉候选项等长文本不再截断）
                exp = str(r["期望值"])[:60]
                act = str(r["实际值"])[:60]
                status = r["是否一致"]
                print(f"{name:<30} {exp:<30} {act:<30} {status}")

    def to_file(self, filepath: str):
        """将结果写入文件"""
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"委托设置测试报告\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'='*60}\n\n")

            total = len(self.results)
            diff = len(self.differences)
            passed = total - diff - self.not_enabled
            f.write(f"总项目数: {total}\n")
            f.write(f"通过: {passed}\n")
            f.write(f"差异: {diff}\n")
            f.write(f"未启用(默认预期，不计差异): {self.not_enabled}\n\n")

            f.write(f"{'名称':<35} {'期望值':<30} {'实际值':<30} {'状态'}\n")
            f.write(f"{'-'*100}\n")
            for r in self.results:
                name = r["名称"][:33]
                # 列宽30，超出时完整显示（下拉候选项等长文本不再截断）
                exp = str(r["期望值"])[:60]
                act = str(r["实际值"])[:60]
                status = r["是否一致"]
                f.write(f"{name:<35} {exp:<30} {act:<30} {status}\n")

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
    #_try_keyboard_select(win)
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
    """查找已打开的设置对话框。

    兼容多种形态：
        - 顶级窗口（window_text 直接含标题）
        - #32770 对话框（标题只存在于 TitleBar 的 Value 属性中，
          窗口自身 window_text() 为空，如本软件的交易系统设置）
        - 主窗口的子窗口 / 后代窗口
    """
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
        # 用 auto_id 直接匹配
        btn = win.child_window(
            auto_id=SETTINGS_BUTTON_AUTO_ID,
            control_type="Button"
        )
        btn.wait("enabled", timeout=5)
        btn.click_input()
        print(f"[OK] 已点击设置按钮 (auto_id={SETTINGS_BUTTON_AUTO_ID})")
        time.sleep(0.5)
    except Exception:
        # 降级：用 win32gui 枚举子窗口按 auto_id 查找
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
                # 计算中心坐标
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
    """在弹出菜单中选择'交易系统设置' (auto_id=20025)。

    弹出菜单是动态创建的，需要遍历顶级窗口匹配 MenuItem。
    如果 UIA 匹配不到，降级为键盘方向键。
    """
    # 尝试用 UIA 匹配菜单项
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


def _try_keyboard_select(win):
    """降级方案：用键盘方向键选择菜单项。"""
    try:
        win.set_focus()
        time.sleep(0.2)
        # 先确定几级菜单才到达"交易系统设置"，尝试不同次数
        # 通常设置按钮的上下文菜单中"交易系统设置"是第二个或第三个
        for down_count in [2, 3, 4, 1]:
            win.type_keys(f"{{DOWN {down_count}}}{{ENTER}}", with_spaces=False)
            time.sleep(1.5)
            if _find_existing_settings_dlg(win) is not None:
                print(f"[OK] 键盘降级方案成功 (按Down {down_count}次 + Enter)")
                return True
        print("[WARN] 键盘降级方案未找到确切菜单项")
    except Exception as e:
        print(f"[WARN] 键盘选定方案失败: {e}")


def switch_to_settings_panel(dlg, panel_name: str = PANEL_NAME) -> bool:
    """在设置对话框中切换到指定标签页（左侧导航：Tree/TreeItem 或 List/ListItem）。

    注意：打开对话框后默认选中的未必是目标面板（软件会记住上次打开的页签），
    因此必须主动切换，而不是假设默认即目标面板。仅在“已确认当前就在目标面板”
    时才跳过点击；切换失败时返回 False，由调用方决定是否终止。
    """
    # 1. 收集可能的左侧导航容器（兼容不同软件版本/形态）
    nav_candidates = []
    try:
        tree = dlg.child_window(control_type="Tree")
        if tree.exists(timeout=1):
            nav_candidates.append(tree)
    except Exception:
        pass
    try:
        lst = dlg.child_window(auto_id="2210", control_type="List")
        if lst.exists(timeout=1):
            nav_candidates.append(lst)
    except Exception:
        pass

    if not nav_candidates:
        print(f"[WARN] 未找到左侧导航容器，无法切换到'{panel_name}'面板")
        return False

    # 2. 依次尝试各导航容器，找到目标项并点击切换
    for nav in nav_candidates:
        try:
            nav.wait("ready", timeout=3)
            nav.set_focus()

            # 目标项可能是 TreeItem 或 ListItem，两种都尝试
            item = None
            for ctype in ("TreeItem", "ListItem"):
                try:
                    cand = nav.child_window(title=panel_name, control_type=ctype, found_index=0)
                    if cand.exists(timeout=1):
                        item = cand
                        break
                except Exception:
                    continue
            if item is None:
                continue

            item.wait("visible", timeout=3)

            # 已选中则无需点击
            try:
                if item.is_selected():
                    print(f"[OK] 当前已在'{panel_name}'面板，无需切换")
                    return True
            except Exception:
                pass

            item.click_input()
            time.sleep(0.6)
            print(f"[OK] 已切换到'{panel_name}'面板")
            return True
        except Exception as e:
            print(f"  [WARN] 通过导航容器切换'{panel_name}'失败: {e}")
            continue

    print(f"[WARN] 切换'{panel_name}'面板失败")
    return False


def get_checkbox_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取复选框的选中状态。

    Returns:
        True=已勾选, False=未勾选, None=找不到控件
    """
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        cb.wait("ready", timeout=2)
        return bool(cb.get_toggle_state())
    except Exception as e:
        print(f"  [WARN] 获取复选框(auto_id={auto_id})失败: {e}")
        return None


def set_checkbox_by_id(dlg, auto_id: str, value: bool) -> bool:
    """通过 auto_id 设置复选框的状态（点击以切换）。"""
    try:
        cb = dlg.child_window(auto_id=auto_id, control_type="CheckBox")
        cb.wait("ready", timeout=2)
        current = bool(cb.get_toggle_state())
        if current != value:
            cb.click_input()
            time.sleep(0.2)
            print(f"  [OK] 复选框(auto_id={auto_id})已设为{value}")
        else:
            print(f"  [INFO] 复选框(auto_id={auto_id})已经是{value}，无需更改")
        return True
    except Exception as e:
        print(f"  [ERROR] 设置复选框(auto_id={auto_id})失败: {e}")
        return False


def get_combobox_selection_by_id(dlg, auto_id: str) -> Optional[str]:
    """通过 auto_id 获取下拉框当前选择的文本。

    Returns:
        当前选中的文本，找不到则返回None
    """
    try:
        combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox")
        combo.wait("ready", timeout=2)
        selected_text = combo.selected_text()
        return selected_text.strip() if selected_text else None
    except Exception as e:
        print(f"  [WARN] 获取下拉框(auto_id={auto_id})失败: {e}")
        return None


# 左侧导航菜单项（与下拉候选项无关，读取 ListItem 时需排除，避免被当成候选项）
_NAV_MENU_ITEMS = {
    "委托设置", "期权设置", "自动拆单设置", "自动追单设置",
    "快捷设置", "键盘下单设置", "价格提醒设置",
}


def _toggle_combobox(combo, open_it: bool):
    """展开/收起下拉框。

    优先点击下拉框内部的“打开”按钮（auto_id=DropDown 或 仅 Button），
    若找不到按钮（运行时 auto_id 可能不一致），则直接点击 combobox 右侧
    的下拉箭头区域（绝对坐标），以保证一定能点开/收起。
    """
    # 策略1：内部 Button（先按精确 auto_id，再退化为任意 Button）
    for kwargs in ({"auto_id": "DropDown", "control_type": "Button"},
                   {"control_type": "Button"}):
        try:
            btn = combo.child_window(**kwargs)
            if btn.exists(timeout=1):
                btn.click_input()
                return
        except Exception:
            continue

    # 策略2：点击 combobox 右侧下拉箭头区域（相对控件右边缘约 8px，垂直居中）
    try:
        rect = combo.rectangle()
        x = rect.right - 8
        y = (rect.top + rect.bottom) // 2
        mouse.click(coords=(x, y))
    except Exception as e:
        raise RuntimeError(f"无法点击下拉箭头: {e}")


def get_combobox_items_by_id(dlg, auto_id: str) -> Optional[List[str]]:
    """点击打开下拉框，读取其包含的所有候选项，然后再次点击一次关闭。

    流程：
        1. 点击下拉框右侧箭头展开列表
        2. 读取全部候选项文本（优先用 item_texts，失败则降级读取弹出的 ListItem）
        3. 再次点击一次收起下拉列表

    Returns:
        候选项文本列表（已去除空白），找不到或无法读取返回None
    """
    try:
        combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox", found_index=0)
        combo.wait("ready", timeout=2)

        # 1. 点击箭头展开下拉列表
        combo.expand()
        #_toggle_combobox(combo, open_it=True)
        print(f"  [INFO] 已点开下拉框(auto_id={auto_id})，读取候选项...")
        time.sleep(0.6)

        # 2. 读取所有候选项
        #    下拉弹层与左侧导航菜单都是 ListItem，且弹层会被重复枚举，
        #    因此需：(a) 排除导航菜单项 (b) 按名称去重并保留首次出现顺序。
        items: List[str] = []
        seen: set = set()
        try:
            for li in combo.descendants(control_type="ListItem"):
                txt = (li.window_text() or "").strip()
                if not txt:
                    continue
                if txt in _NAV_MENU_ITEMS:      # 跳过左侧导航菜单项
                    continue
                if txt not in seen:              # 去重
                    seen.add(txt)
                    items.append(txt)
        except Exception:
            items = []
        print(f"  [INFO] 已读取到 {len(items)} 个下拉候选项（已排除菜单噪声并去重）")

        # 兜底：若 ListItem 方式未取到，再尝试 item_texts
        if not items:
            try:
                texts = combo.item_texts()
                items = [t.strip() for t in texts if t and t.strip()]
            except Exception:
                items = []
        # 3. 再次点击一次收起下拉列表
        #_toggle_combobox(combo, open_it=False)
        combo.collapse()
        print(f"  [INFO] 已关闭下拉框(auto_id={auto_id})")
        time.sleep(0.4)

        return items if items else None
    except Exception as e:
        print(f"  [WARN] 获取下拉框候选项(auto_id={auto_id})失败: {e}")
        return None


def get_edit_value_by_id(dlg, auto_id: str) -> Optional[int]:
    """通过 auto_id 获取数值输入框(Edit)的值。

    Returns:
        整数值，找不到返回None
    """
    try:
        edit = dlg.child_window(auto_id=auto_id, control_type="Edit")
        edit.wait("exists", timeout=5)
        # 子控件刚随上级开关启用时可能尚未完全就绪，放宽 ready 要求并重试
        for _ in range(2):
            try:
                edit.wait("ready", timeout=2)
            except Exception:
                pass
            try:
                text = edit.get_value().strip()
                # 清理可能的逗号分隔符
                text = text.replace(",", "")
                return int(text) if text.isdigit() else None
            except Exception:
                time.sleep(0.5)
        return None
    except Exception as e:
        print(f"  [WARN] 获取数值框(auto_id={auto_id})失败: {e}")
        return None


def get_radiobutton_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取 RadioButton 的选中状态。

    多策略检测，兼容不同控件实现：
        1. UIA TogglePattern (get_toggle_state)
        2. SelectionItemPattern (is_selected / element_info.selection_item_is_selected)
        3. LegacyAccessible State 文本含 "checked"
    任一策略命中即返回；全部失败返回 None。

    Returns:
        True=已选中, False=未选中, None=找不到
    """
    try:
        rb = dlg.child_window(auto_id=auto_id, control_type="RadioButton")
        rb.wait("exists", timeout=5)
        # 子控件刚随上级开关启用时可能尚未完全就绪，放宽 ready 要求并重试
        for _ in range(2):
            try:
                rb.wait("ready", timeout=2)
            except Exception:
                pass

            # 策略1: UIA TogglePattern
            try:
                return bool(rb.get_toggle_state())
            except Exception:
                pass

            # 策略2: SelectionItemPattern (is_selected)
            try:
                return bool(rb.is_selected())
            except Exception:
                pass

            # 策略3: 直接读取 element_info 选中属性
            try:
                return bool(rb.element_info.selection_item_is_selected)
            except Exception:
                pass

            # 策略4: LegacyAccessible State 文本
            try:
                state = rb.legacy_properties().get("State", "") or ""
                return "checked" in state.lower()
            except Exception:
                pass

            time.sleep(0.5)
        return None
    except Exception as e:
        print(f"  [WARN] 获取RadioButton(auto_id={auto_id})失败: {e}")
        return None


# 兼容旧接口（保留，便于其它模块调用）
get_checkbox_state = lambda dlg, title: get_checkbox_state_by_id(dlg, AUTO_ID.get(title, ""))
get_combobox_selection = lambda dlg, title: get_combobox_selection_by_id(dlg, AUTO_ID.get(title, ""))
get_spinbox_value = lambda dlg, auto_id=None, spinbox_title=None: get_edit_value_by_id(dlg, auto_id or AUTO_ID.get(spinbox_title, ""))
get_radiobutton_state = lambda dlg, title: get_radiobutton_state_by_id(dlg, AUTO_ID.get(title, ""))


def take_screenshot(dlg, save_path: str):
    """对“交易系统设置”对话框整体截图保存。

    优先使用 mss 抓取（基于 GDI/DXGI 的整屏区域截取，兼容性好，可避免
    PIL ImageGrab 在高 DPI / 硬件加速窗口下截出黑图的问题）；若 mss 不可用
    再回退到 PIL ImageGrab。
    """
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        rect = dlg.rectangle()
        left, top = int(rect.left), int(rect.top)
        width = int(rect.right - rect.left)
        height = int(rect.bottom - rect.top)
        if width <= 0 or height <= 0:
            raise ValueError("窗口矩形无效，无法截图")

        # 优先：mss 按对话框矩形区域截图
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


def test_price_tracking_settings(dlg, result: SettingsTestResult):
    """测试一、股票买卖委托价格跟盘设置"""
    print("\n--- [1/4] 股票买卖委托价格跟盘设置 ---")

    # 买入缺省价
    buy_checked = get_checkbox_state_by_id(dlg, AUTO_ID["买入缺省价"])
    result.add_result("买入缺省价_勾选", buy_checked, STANDARD_VALUES["买入缺省价_勾选"])

    if buy_checked:
        buy_option = get_combobox_selection_by_id(dlg, AUTO_ID["买入缺省价_下拉"])
        result.add_result("买入缺省价_选项", buy_option or "(未知)", STANDARD_VALUES["买入缺省价_选项"])

        # 下拉框候选项列表（点开检测后关闭）
        buy_items = get_combobox_items_by_id(dlg, AUTO_ID["买入缺省价_下拉"])
        result.add_result("买入缺省价_下拉_选项列表",
                          "、".join(buy_items) if buy_items else "(无法获取)",
                          "、".join(STANDARD_VALUES["买入缺省价_下拉_选项列表"]))
    else:
        result.add_not_enabled("买入缺省价_选项")
        result.add_not_enabled("买入缺省价_下拉_选项列表")

    # 卖出缺省价
    sell_checked = get_checkbox_state_by_id(dlg, AUTO_ID["卖出缺省价"])
    result.add_result("卖出缺省价_勾选", sell_checked, STANDARD_VALUES["卖出缺省价_勾选"])

    if sell_checked:
        sell_option = get_combobox_selection_by_id(dlg, AUTO_ID["卖出缺省价_下拉"])
        result.add_result("卖出缺省价_选项", sell_option or "(未知)", STANDARD_VALUES["卖出缺省价_选项"])

        # 下拉框候选项列表（点开检测后关闭）
        sell_items = get_combobox_items_by_id(dlg, AUTO_ID["卖出缺省价_下拉"])
        result.add_result("卖出缺省价_下拉_选项列表",
                          "、".join(sell_items) if sell_items else "(无法获取)",
                          "、".join(STANDARD_VALUES["卖出缺省价_下拉_选项列表"]))
    else:
        result.add_not_enabled("卖出缺省价_选项")
        result.add_not_enabled("卖出缺省价_下拉_选项列表")


def _check_split_item(dlg, result: SettingsTestResult, checkbox_key: str,
                      checkbox_id: str, value_id: str) -> None:
    """对一项"拆单"开关执行：

        检测初始状态 → (未启用则先启用) → 检测下方数值 → (检测后恢复为未启用)

    开关本身的初始/恢复状态写入报告，下方数值始终检测。
    """
    # 1. 检测开关初始状态
    initial = get_checkbox_state_by_id(dlg, checkbox_id)
    result.add_result(f"{checkbox_key}_初始状态", initial,
                      STANDARD_VALUES[checkbox_key + "_勾选"])

    # 2. 若未启用，先点击启用以暴露下方数值
    need_restore = False
    if not initial:
        print(f"  [INFO] '{checkbox_key}'未勾选，点击启用以暴露下方数值...")
        set_checkbox_by_id(dlg, checkbox_id, True)
        need_restore = True
        time.sleep(0.6)
        now_enabled = get_checkbox_state_by_id(dlg, checkbox_id)
        result.add_result(f"{checkbox_key}_启用后", now_enabled, True)
    else:
        print(f"  [INFO] '{checkbox_key}'已勾选，直接检查下方数值")

    # 3. 检测下方数值（启用状态下）
    split_val = get_edit_value_by_id(dlg, value_id)
    result.add_result(f"{checkbox_key}_数值", split_val if split_val is not None else 0,
                      STANDARD_VALUES[checkbox_key + "_数值"])

    # 4. 若之前未启用，检测完成后恢复为未启用状态
    if need_restore:
        print(f"  [INFO] 检查完成，恢复'{checkbox_key}'为未启用状态...")
        set_checkbox_by_id(dlg, checkbox_id, False)
        time.sleep(0.4)
        restored = get_checkbox_state_by_id(dlg, checkbox_id)
        result.add_result(f"{checkbox_key}_恢复后", restored,
                          STANDARD_VALUES[checkbox_key + "_勾选"])


def test_auto_split_settings(dlg, result: SettingsTestResult):
    """测试二、大单自动分单设置

    股票拆单默认启用，基金会拆单默认未启用。为能验证默认参数，对未启用的
    项采取与委托数量设置类似的策略：先启用以暴露下方数值，检测完后再恢复。
    """
    print("\n--- [2/4] 大单自动分单设置 ---")

    # 股票拆单（默认启用）
    _check_split_item(
        dlg, result,
        checkbox_key="股票拆单",
        checkbox_id=AUTO_ID["股票拆单"],
        value_id=AUTO_ID["股票拆单_数值"],
    )

    # 基金拆单（默认未启用）
    _check_split_item(
        dlg, result,
        checkbox_key="基金拆单",
        checkbox_id=AUTO_ID["基金拆单"],
        value_id=AUTO_ID["基金拆单_数值"],
    )


def _check_qty_item(dlg, result: SettingsTestResult, checkbox_key: str,
                    checkbox_id: str, check_sub) -> None:
    """对一项"自动填入数量"开关执行：

        检测初始状态 → (未启用则先启用) → 检测下方参数 → (检测后恢复为未启用)

    各子参数检测逻辑由 check_sub 回调提供；开关本身的初始/恢复状态写入报告。
    """
    # 1. 检测开关初始状态
    initial = get_checkbox_state_by_id(dlg, checkbox_id)
    result.add_result(f"{checkbox_key}_初始状态", initial,
                      STANDARD_VALUES[checkbox_key + "_勾选"])

    # 2. 若未启用，先点击启用以暴露下方参数
    need_restore = False
    if not initial:
        print(f"  [INFO] '{checkbox_key}'未勾选，点击启用以暴露下方参数...")
        set_checkbox_by_id(dlg, checkbox_id, True)
        need_restore = True
        time.sleep(0.6)
        now_enabled = get_checkbox_state_by_id(dlg, checkbox_id)
        result.add_result(f"{checkbox_key}_启用后", now_enabled, True)
    else:
        print(f"  [INFO] '{checkbox_key}'已勾选，直接检查下方参数")

    # 3. 检测下方参数（启用状态下）
    check_sub()

    # 4. 若之前未启用，检测完成后恢复为未启用状态
    if need_restore:
        print(f"  [INFO] 检查完成，恢复'{checkbox_key}'为未启用状态...")
        set_checkbox_by_id(dlg, checkbox_id, False)
        time.sleep(0.4)
        restored = get_checkbox_state_by_id(dlg, checkbox_id)
        result.add_result(f"{checkbox_key}_恢复后", restored,
                          STANDARD_VALUES[checkbox_key + "_勾选"])


def _buy_sell_qty_sub(dlg, result: SettingsTestResult, prefix: str):
    """股票买入/卖出自动填入数量 的下方参数检测（RadioButton + 数值）。"""
    option = None
    for rid in (AUTO_ID[f"{prefix}_确定数量"], AUTO_ID[f"{prefix}_全部数量"],
                AUTO_ID[f"{prefix}_上一次交易数量"]):
        if get_radiobutton_state_by_id(dlg, rid):
            option = RADIO_NAMES.get(rid)
            break
    result.add_result(f"{prefix}_选项", option or "(无选中)", STANDARD_VALUES[f"{prefix}_选项"])

    qty_val = get_edit_value_by_id(dlg, AUTO_ID[f"{prefix}_数值"])
    result.add_result(f"{prefix}_数值", qty_val if qty_val is not None else 0,
                      STANDARD_VALUES[f"{prefix}_数值"])


def _trade_qty_sub(dlg, result: SettingsTestResult, prefix: str):
    """期权/期货交易自动填入数量 的下方参数检测（仅数值）。"""
    qty_val = get_edit_value_by_id(dlg, AUTO_ID[f"{prefix}_数值"])
    result.add_result(f"{prefix}_数值", qty_val if qty_val is not None else 0,
                      STANDARD_VALUES[f"{prefix}_数值"])


def test_quantity_settings(dlg, result: SettingsTestResult):
    """测试三、委托数量设置

    本组四项（股票买入/卖出自动填入数量、期权交易/期货交易自动填入数量）
    默认均为未启用，其下方参数（RadioButton、数值）也随之不可检测（检测到则记为
    “○ 未启用”）。为能验证默认参数，采取与自动追单类似的策略：
        1. 检测开关初始状态
        2. 若未启用 → 先点击启用以暴露下方参数，检测完后再恢复为未启用
        3. 在启用状态下逐项检测下方参数
    """
    print("\n--- [3/4] 委托数量设置 ---")

    # 股票买入自动填入数量（RadioButton + 数值）
    _check_qty_item(
        dlg, result,
        checkbox_key="股票买入自动填入数量",
        checkbox_id=AUTO_ID["股票买入自动填入数量"],
        check_sub=lambda: _buy_sell_qty_sub(dlg, result, "股票买入"),
    )

    # 股票卖出自动填入数量（RadioButton + 数值）
    _check_qty_item(
        dlg, result,
        checkbox_key="股票卖出自动填入数量",
        checkbox_id=AUTO_ID["股票卖出自动填入数量"],
        check_sub=lambda: _buy_sell_qty_sub(dlg, result, "股票卖出"),
    )

    # 期权交易自动填入数量（仅数值）
    _check_qty_item(
        dlg, result,
        checkbox_key="期权交易自动填入数量",
        checkbox_id=AUTO_ID["期权交易自动填入数量"],
        check_sub=lambda: _trade_qty_sub(dlg, result, "期权交易"),
    )

    # 期货交易自动填入数量（仅数值）
    _check_qty_item(
        dlg, result,
        checkbox_key="期货交易自动填入数量",
        checkbox_id=AUTO_ID["期货交易自动填入数量"],
        check_sub=lambda: _trade_qty_sub(dlg, result, "期货交易"),
    )


def test_bottom_checkboxes(dlg, result: SettingsTestResult):
    """测试四、底部复选框"""
    print("\n--- [4/4] 底部复选框设置 ---")

    bottom_checks = [
        "静默委托下单模式",
        "显示期权下单成功提示",
        "显示期权宝软件风险揭示书",
        "委托成交时发出提示音",
        "点击持仓联动行情",
    ]

    for key_name in bottom_checks:
        state = get_checkbox_state_by_id(dlg, AUTO_ID[key_name])
        result.add_result(key_name, state, STANDARD_VALUES[key_name])


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

                # 只打印主要控件类型
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
    print("交易系统设置 - 委托设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult()

    try:
        # 1. 倒计时（让用户把焦点切到钱龙窗口）
        countdown(COUNTDOWN_SEC)

        # 2. 查找主窗口
        hwnd = find_window(WINDOW_KEYWORD)
        print(f"[OK] 已找到主窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)

        # 3. 自动打开设置对话框
        dlg = open_settings_dialog(win)
        dlg.wait("ready", timeout=10)

        # 4. 切换到委托设置面板
        if not switch_to_settings_panel(dlg, PANEL_NAME):
            print("[错误] 无法切换到委托设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉）
        #print("\n正在进行控件探索...")
        #explore_dialog_controls(dlg)

        # 6. 执行各项测试
        test_price_tracking_settings(dlg, result)
        test_auto_split_settings(dlg, result)
        test_quantity_settings(dlg, result)
        test_bottom_checkboxes(dlg, result)

        # 7. 截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"委托设置_{timestamp}.png")
        take_screenshot(dlg, screenshot_path)

        # 8. 输出结果
        result.print_summary()

        # 9. 保存报告
        report_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"委托设置测试报告_{timestamp}.txt")
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
