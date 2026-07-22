# -*- coding: utf-8 -*-
"""
交易系统设置 - 期权设置界面自动化测试
==========================================
功能：
    打开"交易系统设置"对话框，进入"期权设置"标签页，
    逐项读取每个参数的当前值（复选框状态、下拉框选择、数值等），
    与标准值（恢复默认后的参数）比对，记录差异并截图。

界面元素（期权设置）：
    一、默认下单条件
        - 上证期权 / 深证期权 / 商品·股指期权 / 商品期货 / 中金期货
          （各为 ComboBox，默认"对手价"）

    二、回报与虚实度
        - 回报信息提示（ComboBox: 左下角/...）
        - 虚实度阀值（Edit: -10.00，单位 %）

    三、平仓与持仓过滤
        - 默认平仓顺序（ComboBox: 先成先平/...）
        - 持仓过滤设置（ComboBox: 已平仓合约不显示/...）
        - 记录上一次标的/月份过滤状态（CheckBox）

    四、止盈止损设置
        - 止盈止损单默认有效期（ComboBox: 当日/...）
        - 止盈止损按照（ComboBox: 对手价/...）
        - FOK 发出委托（CheckBox）

    五、自成交检查 / 委托定时查询
        - 自成交检查（RadioButton: 关闭/开启）
        - 委托定时查询（RadioButton: 关闭/开启）

    六、到期提醒
        - 合约 N 日内到期提醒（CheckBox + Edit）
        - 卖出平仓所得无法覆盖成本时提醒（CheckBox）

    七、费用扣减
        - 启用费用扣减（CheckBox）
        - 开仓费用 / 平仓费用 / 行权费用（Edit，单位 元/张）
        - 卖出开仓不收取（CheckBox）
        - 扣减方式（RadioButton: 智能收取/开仓+行权收取/开平仓双向收取）

    八、反手指令设置
        - 按权利/义务 / 按认购/认沽（RadioButton）

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

import re

import numpy as np
from PIL import Image
from pywinauto import Application, findwindows
from core.window import find_window, activate_window, countdown, close_settings_dialog


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "期权设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对
STANDARD_VALUES = {
    # 一、默认下单条件
    "默认下单条件_上证期权": "对手价",
    "默认下单条件_深证期权": "对手价",
    "默认下单条件_商品股指期权": "（空白）",
    "默认下单条件_商品期货": "（空白）",
    "默认下单条件_中金期货": "（空白）",
    "默认下单条件_上证期权_选项列表": ["对手价", "挂盘价", "涨停价", "跌停价",
                              "限价", "超价", "市价转限", "市价FAK", "市价F0K"],
    "默认下单条件_深证期权_选项列表": ["对手价", "挂盘价", "涨停价", "跌停价", "限价",
                              "超价", "对方最优价格", "本方最优价格",
                              "即时成交剩余撤销", "五档即成剩撤", "全额成交或撤销"],

    # 二、回报与虚实度
    "回报信息提示": "左下角",
    "回报信息提示_选项列表": ["右下角", "左下角", "正中央", "记住位置"],
    "虚实度阀值": -10.00,

    # 三、平仓与持仓过滤
    "默认平仓顺序": "先成先平",
    "默认平仓顺序_选项列表": ["先成先平", "后成先平", "盈利先平", "亏损先平"],
    "持仓过滤设置": "已平仓合约不显示",
    "持仓过滤设置_选项列表": ["已平仓合约保留一天", "已平仓合约不显示"],
    "记录上一次标的_月份过滤状态": False,

    # 四、止盈止损设置
    "止盈止损单默认有效期": "当日",
    "止盈止损单默认有效期_选项列表": ["永久", "当日"],
    "止盈止损按照": "对手价",
    "止盈止损按照_选项列表": ["对手价", "挂盘价", "超价"],
    "FOK发出委托": False,

    # 五、自成交检查 / 委托定时查询
    "自成交检查": "关闭",
    "委托定时查询": "开启",

    # 六、到期提醒
    "合约_日内到期提醒_勾选": True,
    "合约_日内到期提醒_数值": 7,
    "卖出平仓所得无法覆盖成本时提醒": True,

    # 七、费用扣减
    "启用费用扣减": False,
    "开仓费用": 0.0,
    "平仓费用": 0.0,
    "行权费用": 0.0,
    "卖出开仓不收取": False,
    "扣减方式": "智能收取",

    # 八、反手指令设置
    "反手指令设置": "按权利/义务",

    # 九、超价设置弹窗
    "超价的基准价": "对手价",
    "超价的基准价_选项列表": ["对手价", "挂盘价", "限价"],
}

# 控件 auto_id 映射（来自交易系统设置_期权设置.txt 抓取）
AUTO_ID = {
    # 一、默认下单条件
    "默认下单条件_上证期权": "2086",
    "默认下单条件_深证期权": "2087",
    "默认下单条件_商品股指期权": "2085",
    "默认下单条件_商品期货": "2084",
    "默认下单条件_中金期货": "2083",

    # 二、回报与虚实度
    "回报信息提示": "2102",
    "虚实度阀值": "2206",

    # 三、平仓与持仓过滤
    "默认平仓顺序": "2082",
    "持仓过滤设置": "2097",
    "记录上一次标的_月份过滤状态": "2048",

    # 四、止盈止损设置
    "止盈止损单默认有效期": "2121",
    "止盈止损按照": "2099",
    "FOK发出委托": "2033",

    # 五、自成交检查 / 委托定时查询
    "自成交检查_关闭": "2230",
    "自成交检查_开启": "2228",
    "委托定时查询_关闭": "2231",
    "委托定时查询_开启": "2236",

    # 六、到期提醒
    "合约_日内到期提醒": "2482",
    "合约_日内到期提醒_数值": "2483",
    "卖出平仓所得无法覆盖成本时提醒": "2484",

    # 七、费用扣减
    "启用费用扣减": "2031",
    "开仓费用": "2168",
    "平仓费用": "2155",
    "行权费用": "2162",
    "卖出开仓不收取": "2054",
    "扣减方式_智能收取": "2220",
    "扣减方式_开仓行权收取": "2221",
    "扣减方式_开平仓双向收取": "2222",

    # 八、反手指令设置
    "反手指令设置_按权利义务": "2237",
    "反手指令设置_按认购认沽": "2235",

    # 九、超价参数
    "超价参数": "2007",
    # 超价设置弹窗内的“超价的基准价”下拉框
    "超价的基准价": "2098",
    "超价的基准价_选项列表": ["对手价", "挂盘价", "限价"],
}

# RadioButton auto_id -> 名称
RADIO_NAMES = {
    "2230": "关闭", "2228": "开启",            # 自成交检查
    "2231": "关闭", "2236": "开启",            # 委托定时查询
    "2220": "智能收取", "2221": "开仓+行权收取", "2222": "开平仓双向收取",  # 扣减方式
    "2237": "按权利/义务", "2235": "按认购/认沽",  # 反手指令设置
}

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "期权设置"
COUNTDOWN_SEC = 3  # 倒计时秒数

# 超价参数弹窗比对用的模板 Excel（与脚本同目录）
SUPER_PRICE_XLSX = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "超价设置模板.xlsx"
)

# RapidOCR 识别最低置信度
OCR_MIN_CONF = 0.30
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
            f.write(f"期权设置测试报告\n")
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
WM_COMMAND = win32con.WM_COMMAND
BM_GETCHECK = win32con.BM_GETCHECK
BM_SETCHECK = win32con.BM_SETCHECK
BM_CLICK = win32con.BM_CLICK
BST_CHECKED = win32con.BST_CHECKED
BST_UNCHECKED = win32con.BST_UNCHECKED
CB_GETCOUNT = win32con.CB_GETCOUNT
CB_GETCURSEL = win32con.CB_GETCURSEL
CB_GETLBTEXT = win32con.CB_GETLBTEXT
LB_GETCOUNT = win32con.LB_GETCOUNT
LB_GETCURSEL = win32con.LB_GETCURSEL
LB_GETTEXT = win32con.LB_GETTEXT
LB_GETTEXTLEN = win32con.LB_GETTEXTLEN
LB_FINDSTRINGEXACT = win32con.LB_FINDSTRINGEXACT
LB_SETCURSEL = win32con.LB_SETCURSEL
LBN_SELCHANGE = win32con.LBN_SELCHANGE


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


def _get_combo_hwnd(dlg, auto_id: str):
    """获取组合框的原生窗口句柄（直接复用通用定位）。"""
    return _get_control_hwnd(dlg, auto_id)


def _read_combobox_items_win32(hwnd: int):
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


# 左侧导航菜单项（与下拉候选项无关，读取 ListItem 时需排除）
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
                if not txt:
                    continue
                if txt in _NAV_MENU_ITEMS:
                    continue
                if txt not in seen:
                    seen.add(txt)
                    items.append(txt)
        except Exception:
            items = []
        print(f"  [INFO] 已读取到 {len(items)} 个下拉候选项（已排除菜单噪声并去重）")

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


def get_radiobutton_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取 RadioButton 的选中状态（Win32 BM_GETCHECK，速度快）。"""
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            state = _win32_user32().SendMessageW(hwnd, BM_GETCHECK, 0, 0)
            return bool(state & 1)
    except Exception as e:
        print(f"  [WARN] win32 读取RadioButton(auto_id={auto_id})失败，降级到 UIA: {e}")
    # 降级：UIA（多策略兼容）
    try:
        rb = dlg.child_window(auto_id=auto_id, control_type="RadioButton")
        rb.wait("exists", timeout=5)
        for _ in range(2):
            try:
                rb.wait("ready", timeout=2)
            except Exception:
                pass
            for fn in (lambda: bool(rb.get_toggle_state()),
                       lambda: bool(rb.is_selected()),
                       lambda: bool(rb.element_info.selection_item_is_selected)):
                try:
                    return fn()
                except Exception:
                    pass
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


def get_selected_radiobutton(dlg, auto_ids: List[str]) -> Optional[str]:
    """在一组 RadioButton 中查找当前选中的名称。

    调试用：若全部未检测到，打印各按钮的原始检测值（控制台）。
    """
    detected = []
    for rid in auto_ids:
        state = get_radiobutton_state_by_id(dlg, rid)
        detected.append((rid, state))
        if state:
            return RADIO_NAMES.get(rid)

    if all(s is None for _, s in detected):
        print(f"  [DEBUG] RadioButton 组未检测到任何选中项，各按钮原始值: "
              f"{[(RADIO_NAMES.get(r), s) for r, s in detected]}")
    return None


# 兼容旧接口
get_checkbox_state = lambda dlg, title: get_checkbox_state_by_id(dlg, AUTO_ID.get(title, ""))
get_combobox_selection = lambda dlg, title: get_combobox_selection_by_id(dlg, AUTO_ID.get(title, ""))
get_spinbox_value = lambda dlg, auto_id=None, spinbox_title=None: get_edit_value_by_id(dlg, auto_id or AUTO_ID.get(spinbox_title, ""))
get_radiobutton_state = lambda dlg, title: get_radiobutton_state_by_id(dlg, AUTO_ID.get(title, ""))


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


def test_default_order_conditions(dlg, result: SettingsTestResult):
    """测试一、默认下单条件"""
    print("\n--- [1/8] 默认下单条件 ---")

    keys = [
        "默认下单条件_上证期权",
        "默认下单条件_深证期权",
        "默认下单条件_商品股指期权",
        "默认下单条件_商品期货",
        "默认下单条件_中金期货",
    ]

    for key in keys:
        actual = get_combobox_selection_by_id(dlg, AUTO_ID[key])
        # 下拉框未选择时 actual 为 None，统一显示为“（空白）”以便与期望比对
        actual_disp = actual if actual is not None else "（空白）"
        result.add_result(key, actual_disp, STANDARD_VALUES[key])

    # 上证期权下拉框：点开检查候选项列表（检查后关闭）
    sh_items = get_combobox_items_by_id(dlg, AUTO_ID["默认下单条件_上证期权"])
    sh_actual = "、".join(sh_items) if sh_items else "(无法获取)"
    sh_expect = "、".join(STANDARD_VALUES["默认下单条件_上证期权_选项列表"])
    result.add_result("默认下单条件_上证期权_选项列表", sh_actual, sh_expect)

    # 深证期权下拉框：点开检查候选项列表（检查后关闭）
    sz_items = get_combobox_items_by_id(dlg, AUTO_ID["默认下单条件_深证期权"])
    sz_actual = "、".join(sz_items) if sz_items else "(无法获取)"
    sz_expect = "、".join(STANDARD_VALUES["默认下单条件_深证期权_选项列表"])
    result.add_result("默认下单条件_深证期权_选项列表", sz_actual, sz_expect)


def test_report_and_threshold(dlg, result: SettingsTestResult):
    """测试二、回报信息提示与虚实度阀值"""
    print("\n--- [2/8] 回报信息提示与虚实度阀值 ---")

    report = get_combobox_selection_by_id(dlg, AUTO_ID["回报信息提示"])
    result.add_result("回报信息提示", report or "(未知)", STANDARD_VALUES["回报信息提示"])

    # 回报信息提示下拉框：点开检查候选项列表（检查后关闭）
    report_items = get_combobox_items_by_id(dlg, AUTO_ID["回报信息提示"])
    report_actual = "、".join(report_items) if report_items else "(无法获取)"
    report_expect = "、".join(STANDARD_VALUES["回报信息提示_选项列表"])
    result.add_result("回报信息提示_选项列表", report_actual, report_expect)

    threshold = get_edit_value_by_id(dlg, AUTO_ID["虚实度阀值"])
    result.add_result("虚实度阀值", threshold if threshold is not None else "(未知)",
                      STANDARD_VALUES["虚实度阀值"])


def test_close_and_filter(dlg, result: SettingsTestResult):
    """测试三、默认平仓顺序与持仓过滤"""
    print("\n--- [3/8] 默认平仓顺序与持仓过滤 ---")

    close_order = get_combobox_selection_by_id(dlg, AUTO_ID["默认平仓顺序"])
    result.add_result("默认平仓顺序", close_order or "(未知)", STANDARD_VALUES["默认平仓顺序"])

    # 默认平仓顺序下拉框：点开检查候选项列表（检查后关闭）
    order_items = get_combobox_items_by_id(dlg, AUTO_ID["默认平仓顺序"])
    order_actual = "、".join(order_items) if order_items else "(无法获取)"
    order_expect = "、".join(STANDARD_VALUES["默认平仓顺序_选项列表"])
    result.add_result("默认平仓顺序_选项列表", order_actual, order_expect)

    filter_setting = get_combobox_selection_by_id(dlg, AUTO_ID["持仓过滤设置"])
    result.add_result("持仓过滤设置", filter_setting or "(未知)", STANDARD_VALUES["持仓过滤设置"])

    # 持仓过滤设置下拉框：点开检查候选项列表（检查后关闭）
    filter_items = get_combobox_items_by_id(dlg, AUTO_ID["持仓过滤设置"])
    filter_actual = "、".join(filter_items) if filter_items else "(无法获取)"
    filter_expect = "、".join(STANDARD_VALUES["持仓过滤设置_选项列表"])
    result.add_result("持仓过滤设置_选项列表", filter_actual, filter_expect)

    record = get_checkbox_state_by_id(dlg, AUTO_ID["记录上一次标的_月份过滤状态"])
    result.add_result("记录上一次标的/月份过滤状态", record, STANDARD_VALUES["记录上一次标的_月份过滤状态"])


def test_stop_loss_profit(dlg, result: SettingsTestResult):
    """测试四、止盈止损设置"""
    print("\n--- [4/8] 止盈止损设置 ---")

    valid = get_combobox_selection_by_id(dlg, AUTO_ID["止盈止损单默认有效期"])
    result.add_result("止盈止损单默认有效期", valid or "(未知)", STANDARD_VALUES["止盈止损单默认有效期"])

    # 止盈止损单默认有效期下拉框：点开检查候选项列表（检查后关闭）
    valid_items = get_combobox_items_by_id(dlg, AUTO_ID["止盈止损单默认有效期"])
    valid_actual = "、".join(valid_items) if valid_items else "(无法获取)"
    valid_expect = "、".join(STANDARD_VALUES["止盈止损单默认有效期_选项列表"])
    result.add_result("止盈止损单默认有效期_选项列表", valid_actual, valid_expect)

    price_by = get_combobox_selection_by_id(dlg, AUTO_ID["止盈止损按照"])
    result.add_result("止盈止损按照", price_by or "(未知)", STANDARD_VALUES["止盈止损按照"])

    # 止盈止损按照下拉框：点开检查候选项列表（检查后关闭）
    price_items = get_combobox_items_by_id(dlg, AUTO_ID["止盈止损按照"])
    price_actual = "、".join(price_items) if price_items else "(无法获取)"
    price_expect = "、".join(STANDARD_VALUES["止盈止损按照_选项列表"])
    result.add_result("止盈止损按照_选项列表", price_actual, price_expect)

    fok = get_checkbox_state_by_id(dlg, AUTO_ID["FOK发出委托"])
    result.add_result("FOK发出委托", fok, STANDARD_VALUES["FOK发出委托"])


def test_self_trade_and_query(dlg, result: SettingsTestResult):
    """测试五、自成交检查与委托定时查询"""
    print("\n--- [5/8] 自成交检查与委托定时查询 ---")

    self_trade = get_selected_radiobutton(dlg, [
        AUTO_ID["自成交检查_关闭"],
        AUTO_ID["自成交检查_开启"],
    ])
    result.add_result("自成交检查", self_trade or "(无选中)", STANDARD_VALUES["自成交检查"])

    timed_query = get_selected_radiobutton(dlg, [
        AUTO_ID["委托定时查询_关闭"],
        AUTO_ID["委托定时查询_开启"],
    ])
    result.add_result("委托定时查询", timed_query or "(无选中)", STANDARD_VALUES["委托定时查询"])


def test_expiration_reminder(dlg, result: SettingsTestResult):
    """测试六、到期提醒"""
    print("\n--- [6/8] 到期提醒 ---")

    contract_checked = get_checkbox_state_by_id(dlg, AUTO_ID["合约_日内到期提醒"])
    result.add_result("合约_日内到期提醒_勾选", contract_checked,
                      STANDARD_VALUES["合约_日内到期提醒_勾选"])

    if contract_checked:
        contract_days = get_edit_value_by_id(dlg, AUTO_ID["合约_日内到期提醒_数值"])
        result.add_result("合约_日内到期提醒_数值",
                          contract_days if contract_days is not None else 0,
                          STANDARD_VALUES["合约_日内到期提醒_数值"])
    else:
        result.add_not_enabled("合约_日内到期提醒_数值")

    cover = get_checkbox_state_by_id(dlg, AUTO_ID["卖出平仓所得无法覆盖成本时提醒"])
    result.add_result("卖出平仓所得无法覆盖成本时提醒", cover,
                      STANDARD_VALUES["卖出平仓所得无法覆盖成本时提醒"])


def test_fee_deduction(dlg, result: SettingsTestResult):
    """测试七、费用扣减

    流程：
        1. 检测"启用费用扣减"初始状态（默认未启用）
        2. 若未启用 → 点击启用以暴露下方参数，检查完后再点击关闭恢复
        3. 在启用状态下逐项检查下方参数（开仓费用/平仓费用/行权费用/卖出开仓不收取/扣减方式）
    """
    print("\n--- [7/8] 费用扣减 ---")

    # 1. 检测"启用费用扣减"初始状态
    initial_enabled = get_checkbox_state_by_id(dlg, AUTO_ID["启用费用扣减"])
    result.add_result("启用费用扣减_初始状态", initial_enabled,
                      STANDARD_VALUES["启用费用扣减"])

    # 2. 若未启用，点击启用以暴露下方参数
    need_restore = False
    if not initial_enabled:
        print("  [INFO] '启用费用扣减'未勾选，点击启用以暴露下方参数...")
        click_checkbox_by_id(dlg, AUTO_ID["启用费用扣减"])
        need_restore = True
        time.sleep(0.6)

        now_enabled = get_checkbox_state_by_id(dlg, AUTO_ID["启用费用扣减"])
        result.add_result("启用费用扣减_启用后", now_enabled, True)
    else:
        print("  [INFO] '启用费用扣减'已勾选，直接检查下方参数")

    # 3. 检查下方参数（启用状态下）
    open_fee = get_edit_value_by_id(dlg, AUTO_ID["开仓费用"])
    result.add_result("开仓费用", open_fee if open_fee is not None else 0.0,
                      STANDARD_VALUES["开仓费用"])

    close_fee = get_edit_value_by_id(dlg, AUTO_ID["平仓费用"])
    result.add_result("平仓费用", close_fee if close_fee is not None else 0.0,
                      STANDARD_VALUES["平仓费用"])

    exercise_fee = get_edit_value_by_id(dlg, AUTO_ID["行权费用"])
    result.add_result("行权费用", exercise_fee if exercise_fee is not None else 0.0,
                      STANDARD_VALUES["行权费用"])

    sell_open_free = get_checkbox_state_by_id(dlg, AUTO_ID["卖出开仓不收取"])
    result.add_result("卖出开仓不收取", sell_open_free,
                      STANDARD_VALUES["卖出开仓不收取"])

    deduction = get_selected_radiobutton(dlg, [
        AUTO_ID["扣减方式_智能收取"],
        AUTO_ID["扣减方式_开仓行权收取"],
        AUTO_ID["扣减方式_开平仓双向收取"],
    ])
    result.add_result("扣减方式", deduction or "(无选中)", STANDARD_VALUES["扣减方式"])

    # 4. 若之前未启用，检查完成后恢复关闭状态
    if need_restore:
        print("  [INFO] 检查完成，恢复'启用费用扣减'为关闭状态...")
        click_checkbox_by_id(dlg, AUTO_ID["启用费用扣减"])
        time.sleep(0.4)
        restored = get_checkbox_state_by_id(dlg, AUTO_ID["启用费用扣减"])
        result.add_result("启用费用扣减_恢复后", restored,
                          STANDARD_VALUES["启用费用扣减"])


def test_reverse_instruction(dlg, result: SettingsTestResult):
    """测试八、反手指令设置"""
    print("\n--- [8/8] 反手指令设置 ---")

    reverse = get_selected_radiobutton(dlg, [
        AUTO_ID["反手指令设置_按权利义务"],
        AUTO_ID["反手指令设置_按认购认沽"],
    ])
    result.add_result("反手指令设置", reverse or "(无选中)", STANDARD_VALUES["反手指令设置"])


# 弹窗中可点击以关闭的按钮标题（按优先级）
_POPUP_CLOSE_TITLES = ("取消", "关闭", "确定")


def _get_window_title(win) -> str:
    """多源读取窗口标题（兼容标题仅存在于 TitleBar 的各种情况）。"""
    candidates = []
    try:
        candidates.append(win.window_text() or "")
    except Exception:
        pass
    try:
        candidates.append(win.element_info.name or "")
    except Exception:
        pass
    try:
        tb = win.child_window(control_type="TitleBar")
        try:
            candidates.append(tb.window_text() or "")
        except Exception:
            pass
        try:
            candidates.append(tb.legacy_properties().get("Value", "") or "")
        except Exception:
            pass
    except Exception:
        pass
    for c in candidates:
        if c:
            return c
    return ""


def _get_window_title_by_handle(hwnd: int) -> str:
    try:
        app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
        return _get_window_title(app.window(handle=hwnd))
    except Exception:
        return ""


def find_popup_dialog(dlg, timeout: float = 6) -> Optional[Any]:
    """点击'超价参数'后弹出的对话框（标题含“超价设置”）。

    搜索策略（兼容多种形态）：
        1. 设置对话框的直接子 Window
        2. 顶级 #32770 窗口，标题栏读取的标题含“超价”
    一旦识别到标题即返回，不再强依赖 wait("ready")，避免弹窗未就绪导致漏检。
    """
    target_keyword = "超价"
    end = time.time() + timeout
    while time.time() < end:
        # 策略1：设置对话框的子窗口
        try:
            for child in dlg.descendants(control_type="Window"):
                try:
                    title = _get_window_title(child)
                    if title and "交易系统设置" not in title:
                        print(f"[OK] 已找到弹窗(子窗口): {title}")
                        return child
                except Exception:
                    continue
        except Exception:
            pass

        # 策略2：顶级窗口（含 #32770），标题栏读取标题
        try:
            for elem in findwindows.find_elements(top_level_only=True):
                try:
                    if elem.window_text() == "交易系统设置":
                        continue
                    # 直接有标题的窗口：仅当标题含关键字才接受，否则跳过（避免误连其它程序窗口）
                    direct = elem.window_text() or ""
                    if direct and "交易系统设置" in direct:
                        continue
                    if direct and target_keyword in direct:
                        app = Application(backend="uia").connect(handle=elem.handle, timeout=1)
                        print(f"[OK] 已找到弹窗(顶级, 标题={direct})")
                        return app.window(handle=elem.handle)
                    if direct:
                        # 非空且不含关键字 → 不是目标，跳过，避免对普通窗口做 UIA 连接
                        continue
                    # 标题为空 → 可能是弹窗（标题在 TitleBar），读取标题栏确认
                    title = _get_window_title_by_handle(elem.handle)
                    if title and target_keyword in title:
                        app = Application(backend="uia").connect(handle=elem.handle, timeout=1)
                        print(f"[OK] 已找到弹窗(顶级, TitleBar={title})")
                        return app.window(handle=elem.handle)
                except Exception:
                    continue
        except Exception:
            pass

        time.sleep(0.3)
    return None


def find_popup_combobox(popup, auto_id: str, title_hint: str = "") -> Optional[Any]:
    """在弹窗中定位下拉框：先按 auto_id，失败则按标签就近找。"""
    # 1. 直接按 auto_id
    try:
        cb = popup.child_window(auto_id=auto_id, control_type="ComboBox")
        cb.wait("exists", timeout=2)
        print(f"  [DEBUG] 下拉框定位: 策略1 按 auto_id={auto_id} 命中")
        return cb
    except Exception:
        print(f"  [DEBUG] 下拉框定位: 策略1 auto_id={auto_id} 失败，尝试其它策略")

    # 2. 按标题直接匹配（少数下拉框本身带标题）
    if title_hint:
        try:
            cb = popup.child_window(title=title_hint, control_type="ComboBox")
            cb.wait("exists", timeout=2)
            print(f"  [DEBUG] 下拉框定位: 策略2 按标题'{title_hint}' 命中")
            return cb
        except Exception:
            print(f"  [DEBUG] 下拉框定位: 策略2 按标题'{title_hint}' 失败")

    # 3. 收集所有可见 ComboBox，按离标签最近的选
    try:
        combos = [c for c in popup.descendants(control_type="ComboBox")]
    except Exception:
        combos = []
    if not combos:
        print("  [DEBUG] 下拉框定位: 未找到任何 ComboBox")
        return None
    if len(combos) == 1:
        print(f"  [DEBUG] 下拉框定位: 策略3 弹窗内仅 1 个 ComboBox，直接采用")
        return combos[0]

    if title_hint:
        for label in popup.descendants(control_type="Text"):
            if title_hint in (label.window_text() or ""):
                try:
                    lr = label.rectangle()
                    lcx = (lr.left + lr.right) / 2
                    lcy = (lr.top + lr.bottom) / 2
                    best = None
                    best_d = None
                    for c in combos:
                        try:
                            cr = c.rectangle()
                            ccx = (cr.left + cr.right) / 2
                            ccy = (cr.top + cr.bottom) / 2
                            d = (lcx - ccx) ** 2 + (lcy - ccy) ** 2
                            if best is None or d < best_d:
                                best = c
                                best_d = d
                        except Exception:
                            continue
                    print(f"  [DEBUG] 下拉框定位: 策略3 按标签'{title_hint}'就近匹配命中")
                    return best
                except Exception:
                    continue
    print("  [DEBUG] 下拉框定位: 所有策略均未命中")
    return None


def get_popup_combobox_text(cb) -> Optional[str]:
    """读取弹窗下拉框当前选中文本（Win32 优先，多策略容错）。"""
    # 策略0：Win32 CB_GETCURSEL + CB_GETLBTEXT（不展开下拉层）
    try:
        hwnd = cb.element_info.handle
        if hwnd:
            user32 = _win32_user32()
            sel = user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0)
            if sel is not None and sel >= 0:
                buf = ctypes.create_unicode_buffer(256)
                user32.SendMessageW(hwnd, CB_GETLBTEXT, sel, ctypes.addressof(buf))
                s = buf.value.strip()
                if s:
                    print(f"  [DEBUG] 读取下拉框当前值: 策略'win32 CB' 生效 -> {s!r}")
                    return s
    except Exception as e:
        print(f"  [DEBUG] 读取下拉框当前值: 策略'win32 CB' 失败: {e}")

    getters = [
        ("selected_text()", lambda: cb.selected_text()),
        ("legacy Value", lambda: cb.legacy_properties().get("Value")),
        ("texts()[-1]", lambda: (cb.texts() or [None])[-1]),
    ]
    for name, g in getters:
        try:
            v = g()
            if v:
                s = str(v).strip()
                if s:
                    print(f"  [DEBUG] 读取下拉框当前值: 策略'{name}' 生效 -> {s!r}")
                    return s
        except Exception as e:
            print(f"  [DEBUG] 读取下拉框当前值: 策略'{name}' 失败: {e}")
    return None


def get_popup_combobox_items(cb) -> Optional[List[str]]:
    """读取弹窗下拉框候选项（Win32 CB_GETCOUNT/CB_GETLBTEXT 优先，不展开）。"""
    # 策略0：Win32 直接读列表（不展开下拉层）
    try:
        hwnd = cb.element_info.handle
        if hwnd:
            items = _read_combobox_items_win32(hwnd)
            if items:
                print(f"  [DEBUG] 读取下拉框候选项: 策略'win32 CB' 生效 -> {items}")
                return items
    except Exception as e:
        print(f"  [DEBUG] 读取下拉框候选项: 策略'win32 CB' 失败: {e}")

    # 策略1：展开 + ListItem
    try:
        cb.expand()
        time.sleep(0.6)
        items, seen = [], set()
        for li in cb.descendants(control_type="ListItem"):
            txt = (li.window_text() or "").strip()
            if txt and txt not in seen:
                seen.add(txt)
                items.append(txt)
        cb.collapse()
        time.sleep(0.3)
        if items:
            print(f"  [DEBUG] 读取下拉框候选项: 策略'expand+ListItem' 生效 -> {items}")
            return items
        else:
            print("  [DEBUG] 读取下拉框候选项: 策略'expand+ListItem' 展开成功但无候选项")
    except Exception as e:
        print(f"  [DEBUG] 读取下拉框候选项: 策略'expand+ListItem' 失败: {e}")
        try:
            cb.collapse()
        except Exception:
            pass

    # 策略2：item_texts
    try:
        texts = cb.item_texts()
        if texts:
            print(f"  [DEBUG] 读取下拉框候选项: 策略'item_texts' 生效 -> {texts}")
            return [t.strip() for t in texts if t and t.strip()]
    except Exception as e:
        print(f"  [DEBUG] 读取下拉框候选项: 策略'item_texts' 失败: {e}")

    # 策略3：texts
    try:
        texts = cb.texts()
        if texts:
            print(f"  [DEBUG] 读取下拉框候选项: 策略'texts' 生效 -> {texts}")
            return [t.strip() for t in texts if t and t.strip()]
    except Exception as e:
        print(f"  [DEBUG] 读取下拉框候选项: 策略'texts' 失败: {e}")

    return None


def close_popup_dialog(popup):
    """关闭弹窗：优先 Esc，再尝试点击 取消/关闭/确定 按钮。"""
    try:
        popup.set_focus()
        time.sleep(0.2)
        popup.type_keys("{ESC}")
        time.sleep(0.4)
        if not popup.exists(timeout=1):
            print("[OK] 已用 Esc 关闭弹窗")
            return
    except Exception:
        pass

    for title in _POPUP_CLOSE_TITLES:
        try:
            btn = popup.child_window(title=title, control_type="Button")
            if btn.exists(timeout=1):
                btn.click_input()
                print(f"[OK] 已点击弹窗按钮'{title}'关闭弹窗")
                return
        except Exception:
            continue
    print("[WARN] 未能自动关闭弹窗，请手动关闭")


def test_super_price_params(dlg, result: SettingsTestResult):
    """测试九、超价参数：点击按钮 → 弹窗 → RapidOCR 识别表格 → 与模板比对。

    流程：
        1. 点击“超价参数”按钮 (auto_id=2007)，弹出设置对话框
        2. 对弹窗截图（numpy 数组）
        3. RapidOCR 识别弹窗内表格文本
        4. 解析为 {品种, 买超价步长, 卖超价步长} 行
        5. 与 超价设置模板.xlsx 逐行比对
        6. 关闭弹窗
    """
    print("\n--- [9/9] 超价参数弹窗 (RapidOCR 比对) ---")

    # 1. 点击“超价参数”按钮
    try:
        btn = dlg.child_window(auto_id=AUTO_ID["超价参数"], control_type="Button")
        btn.wait("enabled", timeout=3)
        btn.click_input()
        print(f"[OK] 已点击'超价参数'按钮 (auto_id={AUTO_ID['超价参数']})")
    except Exception as e:
        print(f"[WARN] 点击'超价参数'按钮失败: {e}")
        result.add_result("超价参数_弹窗", "未打开", "已打开")
        return

    time.sleep(1.0)

    # 2. 查找弹窗
    popup = find_popup_dialog(dlg)
    if popup is None:
        print("[WARN] 未检测到弹窗，可能未弹出或形态未知")
        result.add_result("超价参数_弹窗", "未检测到", "已打开")
        return
    result.add_result("超价参数_弹窗", "已打开", "已打开")

    # 3. 截图弹窗为 numpy 数组（同时保存 PNG 便于核对）
    try:
        rect = popup.rectangle()
        popup_arr = _grab_rect_as_array(rect)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        shot = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"超价参数弹窗_{ts}.png")
        os.makedirs(os.path.dirname(shot), exist_ok=True)
        Image.fromarray(popup_arr).save(shot)
        print(f"[OK] 弹窗截图已保存: {shot}")
    except Exception as e:
        print(f"  [WARN] 弹窗截图失败: {e}")
        close_popup_dialog(popup)
        return

    # 4. RapidOCR 识别
    try:
        tokens = ocr_image(popup_arr)
        print(f"[OK] 超价参数弹窗 OCR 识别到 {len(tokens)} 条文本")
    except Exception as e:
        print(f"  [WARN] OCR 识别失败: {e}")
        close_popup_dialog(popup)
        return

    # 5. 解析 + 读取模板 + 比对
    ocr_rows = parse_super_price_popup(tokens)
    print(f"[OK] 解析到 {len(ocr_rows)} 行超价参数数据: {ocr_rows}")

    if not os.path.exists(SUPER_PRICE_XLSX):
        print(f"[WARN] 模板文件不存在: {SUPER_PRICE_XLSX}")
        result.add_result("超价参数_模板文件", "不存在", SUPER_PRICE_XLSX)
        close_popup_dialog(popup)
        return

    try:
        template_rows = read_super_price_template(SUPER_PRICE_XLSX)
        print(f"[OK] 模板读取到 {len(template_rows)} 行: {template_rows}")
    except Exception as e:
        print(f"  [WARN] 读取模板失败: {e}")
        close_popup_dialog(popup)
        return

    compare_super_price(template_rows, ocr_rows, result)

    # 7. 超价设置弹窗内的“超价的基准价”下拉框检查
    print("\n  [弹窗] 检查‘超价的基准价’下拉框...")
    base_cb = find_popup_combobox(popup, AUTO_ID["超价的基准价"], "超价的基准价")
    if base_cb is None:
        print("  [WARN] 未能在弹窗中找到‘超价的基准价’下拉框")
        result.add_result("超价的基准价", "(未找到)", STANDARD_VALUES["超价的基准价"])
        result.add_result("超价的基准价_选项列表", "(未找到)",
                          "、".join(STANDARD_VALUES["超价的基准价_选项列表"]))
    else:
        base_price = get_popup_combobox_text(base_cb)
        result.add_result("超价的基准价", base_price or "(未知)", STANDARD_VALUES["超价的基准价"])

        base_items = get_popup_combobox_items(base_cb)
        base_actual = "、".join(base_items) if base_items else "(无法获取)"
        base_expect = "、".join(STANDARD_VALUES["超价的基准价_选项列表"])
        result.add_result("超价的基准价_选项列表", base_actual, base_expect)

    # 8. 关闭弹窗
    close_popup_dialog(popup)
    time.sleep(0.5)


# ---------- RapidOCR ----------
_rapid_ocr_instance = None


def get_rapid_ocr():
    """初始化 RapidOCR 实例（单例模式）"""
    global _rapid_ocr_instance
    if _rapid_ocr_instance is not None:
        return _rapid_ocr_instance
    try:
        from rapidocr import RapidOCR
    except ImportError as e:
        print(f"[致命错误] rapidocr 导入失败: {e}")
        print("请执行: pip install rapidocr")
        raise
    print("[..] 正在初始化 RapidOCR (默认 ONNXRuntime CPU + PP-OCRv4)...")
    _rapid_ocr_instance = RapidOCR()
    print("[OK] RapidOCR 初始化完成")
    return _rapid_ocr_instance


def ocr_image(img_rgb: np.ndarray):
    """使用 RapidOCR 识别图像，返回 [{text, box, conf}, ...]。"""
    ocr = get_rapid_ocr()
    result = ocr(img_rgb)
    out = []
    if not result or not result.txts:
        return out
    for i in range(len(result.txts)):
        text = result.txts[i]
        score = float(result.scores[i])
        if score < OCR_MIN_CONF:
            continue
        text = (text or "").strip()
        if not text:
            continue
        box = result.boxes[i]
        box4 = [(int(box[j][0]), int(box[j][1])) for j in range(4)]
        out.append({"text": text, "box": box4, "conf": score})
    return out


def _grab_rect_as_array(rect) -> np.ndarray:
    """用 mss 按窗口矩形区域截图为 numpy RGB 数组。"""
    import mss
    bbox = {
        "left": int(rect.left),
        "top": int(rect.top),
        "width": max(1, int(rect.width())),
        "height": max(1, int(rect.height())),
    }
    with mss.mss() as sct:
        sct_img = sct.grab(bbox)
        rgb = np.array(sct_img)[:, :, :3][:, :, ::-1].copy()
    return rgb


def _box_center(box):
    xs = [p[0] for p in box]
    ys = [p[1] for p in box]
    return (sum(xs) / 4.0, sum(ys) / 4.0)


def _group_tokens_by_row(tokens, y_tol: int = 12):
    """按 y 坐标把 OCR token 分成行。"""
    if not tokens:
        return []
    sorted_toks = sorted(tokens, key=lambda t: (_box_center(t["box"])[1],
                                                _box_center(t["box"])[0]))
    rows = [[sorted_toks[0]]]
    for tok in sorted_toks[1:]:
        cy = _box_center(tok["box"])[1]
        last_cy = _box_center(rows[-1][-1]["box"])[1]
        if abs(cy - last_cy) <= y_tol:
            rows[-1].append(tok)
        else:
            rows.append([tok])
    return rows


def _extract_int(s: str):
    """从文本中提取带符号整数（兼容 '−1'/'-1'/'1'）。"""
    if s is None:
        return None
    s = s.replace("−", "-").replace("－", "-").replace("—", "-")
    m = re.search(r"-?\d+", s)
    return int(m.group()) if m else None


def parse_super_price_popup(tokens):
    """把 OCR token 解析为超价参数表格数据。

    弹窗表头大致为: 品种 | 买超价步长 | 卖超价步长
    返回 list of {"品种": str, "买": int|None, "卖": int|None}
    """
    rows = _group_tokens_by_row(tokens)
    data_rows = []
    for row in rows:
        # 跳过纯表头行（含“品种”且含“超价/步长”）
        texts = "".join(t["text"] for t in row)
        if "品种" in texts and ("超价" in texts or "步长" in texts):
            continue

        sorted_row = sorted(row, key=lambda t: _box_center(t["box"])[0])

        variety = None
        nums = []
        for t in sorted_row:
            txt = (t["text"] or "").strip()
            if not txt:
                continue
            # 表头残留词跳过
            if txt in ("品种", "买超价步长", "卖超价步长", "买超价", "卖超价"):
                continue
            # 含中文 → 视为品种名
            if re.search(r"[\u4e00-\u9fff]", txt):
                variety = (variety + txt) if variety else txt
                continue
            # 数值（买/卖 超价步长）
            v = _extract_int(txt)
            if v is not None:
                nums.append(v)

        if variety and nums:
            data_rows.append({
                "品种": variety,
                "买": nums[0] if len(nums) > 0 else None,
                "卖": nums[1] if len(nums) > 1 else None,
            })
    return data_rows


def read_super_price_template(filepath: str):
    """读取超价设置模板.xlsx → list of {品种, 买, 卖}。

    要求列名（trim 后）：品种 / 买超价步长 / 卖超价步长
    """
    import pandas as pd
    df = pd.read_excel(filepath)
    df.columns = [str(c).strip() for c in df.columns]

    col_map = {}
    aliases = {
        "品种": ["品种", "合约", "名称"],
        "买超价步长": ["买超价步长", "买超价", "买步长"],
        "卖超价步长": ["卖超价步长", "卖超价", "卖步长"],
    }
    for key, alist in aliases.items():
        for a in alist:
            if a in df.columns:
                col_map[key] = a
                break
    missing = [k for k in ("品种", "买超价步长", "卖超价步长") if k not in col_map]
    if missing:
        raise ValueError(f"Excel 缺少必需列: {missing}（实际列: {list(df.columns)}）")

    rows = []
    for _, row in df.iterrows():
        品种 = str(row[col_map["品种"]]).strip()
        if 品种 in ("", "nan"):
            continue
        rows.append({
            "品种": 品种,
            "买": _extract_int(str(row[col_map["买超价步长"]])),
            "卖": _extract_int(str(row[col_map["卖超价步长"]])),
        })
    return rows


def compare_super_price(template_rows, ocr_rows, result: SettingsTestResult):
    """将模板逐行与 OCR 结果比对，记录差异。"""
    for tpl in template_rows:
        品种 = tpl["品种"]
        # 精确匹配品种
        match = next((o for o in ocr_rows if o["品种"] == 品种), None)
        # 模糊匹配（按数字部分包含）
        if match is None:
            digits = re.sub(r"\D", "", 品种)
            if digits:
                match = next((o for o in ocr_rows if digits in re.sub(r"\D", "", o["品种"])), None)

        if match is None:
            result.add_result(f"超价参数_{品种}", "未在弹窗中找到",
                              f"买={tpl['买']}, 卖={tpl['卖']}")
            continue

        result.add_result(f"超价参数_{品种}_买超价步长",
                          match["买"] if match["买"] is not None else "(缺失)",
                          tpl["买"])
        result.add_result(f"超价参数_{品种}_卖超价步长",
                          match["卖"] if match["卖"] is not None else "(缺失)",
                          tpl["卖"])


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
    print("交易系统设置 - 期权设置自动化测试")
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

        # 4. 切换到期权设置面板
        if not switch_to_settings_panel(dlg, PANEL_NAME):
            print("[错误] 无法切换到期权设置面板")
            sys.exit(1)

        time.sleep(0.5)

        # 5. 控件探索（首次运行时有用，可注释掉）
        #print("\n正在进行控件探索...")
        #explore_dialog_controls(dlg)

        # 6. 执行各项测试
        test_default_order_conditions(dlg, result)
        test_report_and_threshold(dlg, result)
        test_close_and_filter(dlg, result)
        test_stop_loss_profit(dlg, result)
        test_self_trade_and_query(dlg, result)
        test_expiration_reminder(dlg, result)
        test_fee_deduction(dlg, result)
        test_reverse_instruction(dlg, result)
        test_super_price_params(dlg, result)

        # 7. 截图
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"期权设置_{timestamp}.png")
        take_screenshot(dlg, screenshot_path)

        # 8. 输出结果
        result.print_summary()

        # 9. 保存报告
        report_path = os.path.join(OUTPUT_DIR, RESULT_SUBDIR, f"期权设置测试报告_{timestamp}.txt")
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
