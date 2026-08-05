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
import ctypes
import win32gui
import win32con
from datetime import datetime
from typing import List, Optional

# 添加项目根目录到路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.settings import SettingsTestResult
from core.settings_standard import load_standard
from core.settings_auto_id import apply_client_auto_id


# ====================== 可配置参数 ======================
WINDOW_KEYWORD = "钱龙模拟"              # 主窗口标题关键字
SETTINGS_BUTTON_AUTO_ID = "1008"     # 设置按钮 auto_id
SETTINGS_MENU_ITEM_AUTO_ID = "20025" # 弹出菜单中"交易系统设置"项 auto_id
SETTINGS_DIALOG_TITLE = "交易系统设置"  # 设置对话框标题
PANEL_NAME = "委托设置"               # 左侧树节点名称

# 标准值（恢复默认后应呈现的值），用于比对。
# 优先从 交易系统设置/标准/<客户端>/委托设置.json 读取（可自定义/抓取覆盖）；
# 找不到时使用下方内嵌兜底（与 qianlong 默认标准一致），保证离线不崩。
DEFAULT_STANDARD_VALUES = {
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

# 当前客户端（GUI 启动时由 GUI_CLIENT_ID 环境变量注入；空则用内嵌兜底）
CLIENT_ID = os.environ.get("GUI_CLIENT_ID", "") or ""
STANDARD_VALUES = load_standard(PANEL_NAME, CLIENT_ID, DEFAULT_STANDARD_VALUES)

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

# 客户端专属控件 ID 覆盖（广发等与默认 AUTO_ID 不同，见 core/settings_auto_id.py）
apply_client_auto_id(PANEL_NAME, AUTO_ID, RADIO_NAMES, CLIENT_ID)

# 底部复选框清单（广发无 委托成交时发出提示音/点击持仓联动行情）
if CLIENT_ID == "guangfa":
    BOTTOM_CHECK_KEYS = [
        "静默委托下单模式",
        "显示期权下单成功提示",
        "显示期权宝软件风险揭示书",
    ]
else:
    BOTTOM_CHECK_KEYS = [
        "静默委托下单模式",
        "显示期权下单成功提示",
        "显示期权宝软件风险揭示书",
        "委托成交时发出提示音",
        "点击持仓联动行情",
    ]

# 输出目录（可被 GUI 传入的 GUI_OUTPUT_DIR 环境变量覆盖）
_OUTPUT_DIR_DEFAULT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "交易系统设置_测试结果")
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
# 每个脚本的结果（报告+截图）单独存放在同名子文件夹中
RESULT_SUBDIR = "委托设置"
COUNTDOWN_SEC = 3  # 倒计时秒数
# ========================================================
















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


def set_checkbox_by_id(dlg, auto_id: str, value: bool) -> bool:
    """通过 auto_id 设置复选框的状态（Win32 BM_CLICK 切换，触发 BN_CLICKED）。"""
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            user32 = _win32_user32()
            cur = user32.SendMessageW(hwnd, BM_GETCHECK, 0, 0)
            cur_checked = bool(cur & 1)
            if cur_checked != value:
                user32.SendMessageW(hwnd, BM_CLICK, 0, 0)
                time.sleep(0.2)
                print(f"  [OK] 复选框(auto_id={auto_id})已设为{value}(win32)")
            else:
                print(f"  [INFO] 复选框(auto_id={auto_id})已经是{value}，无需更改")
            return True
    except Exception as e:
        print(f"  [WARN] win32 设置复选框(auto_id={auto_id})失败，降级到 UIA: {e}")

    # 降级：UIA
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


def _get_control_hwnd(dlg, auto_id: str) -> Optional[int]:
    """用 win32gui 按控件 ID (DlgCtrlID) 查找子窗口句柄（最快，不触发 UIA 树遍历）。

    本文件所有控件 auto_id 均对应原生对话框控件 ID，故可直接用
    GetDlgCtrlID 枚举定位。找不到返回 None。
    """
    try:
        target = int(auto_id)
    except (TypeError, ValueError):
        return None
    found = []
    def _cb(hwnd, _):
        try:
            # 只匹配可见控件：广发等客户端不同页面共用同一 auto_id（如 16074），
            # 隐藏的是其它页面控件，不可见则不应命中
            if (win32gui.GetDlgCtrlID(hwnd) == target
                    and win32gui.IsWindowVisible(hwnd)):
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


def _get_combo_hwnd(dlg, auto_id: str) -> Optional[int]:
    """获取组合框的原生窗口句柄（直接复用通用定位）。"""
    return _get_control_hwnd(dlg, auto_id)


def get_combobox_selection_by_id(dlg, auto_id: str) -> Optional[str]:
    """通过 auto_id 获取下拉框当前选择的文本。

    优化：优先用 CB_GETCURSEL + CB_GETLBTEXT 直接读取（不展开下拉层，
    速度最快）；失败则降级到原 UIA selected_text()。

    Returns:
        当前选中的文本，找不到则返回None
    """
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


# 左侧导航菜单项（与下拉候选项无关，读取 ListItem 时需排除，避免被当成候选项）
_NAV_MENU_ITEMS = {
    "委托设置", "期权设置", "自动拆单设置", "自动追单设置",
    "快捷设置", "键盘下单设置", "价格提醒设置",
}




def _read_combobox_items_win32(hwnd: int) -> Optional[List[str]]:
    """通过 CB_GETCOUNT / CB_GETLBTEXT 直接读取 Win32 组合框的候选项。

    不走 UIA 树遍历、不展开下拉层，因此极快。返回去重后的候选项列表，
    读取失败或为空时返回 None。
    """
    user32 = _win32_user32()
    count = user32.SendMessageW(hwnd, CB_GETCOUNT, 0, 0)
    if count is None or count <= 0:
        return None
    buf = ctypes.create_unicode_buffer(256)
    items: List[str] = []
    seen: set = set()
    for i in range(count):
        user32.SendMessageW(hwnd, CB_GETLBTEXT, i, ctypes.addressof(buf))
        txt = buf.value.strip()
        if txt and txt not in seen:
            seen.add(txt)
            items.append(txt)
    return items if items else None


def _get_combobox_items_uia(dlg, auto_id: str) -> Optional[List[str]]:
    """原 UIA 方案（兼容非标准 / 虚拟化组合框）：展开下拉层并遍历 ListItem。"""
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


def get_combobox_items_by_id(dlg, auto_id: str) -> Optional[List[str]]:
    """读取下拉框的全部候选项（已去重）。

    优化（重点提速项）：优先用 win32 直接对组合框发送 CB_GETCOUNT /
    CB_GETLBTEXT 消息读取列表数据 —— 不展开下拉层、不做 UIA 全树遍历，
    速度比原方案快一个数量级；仅当控件非标准 Win32 组合框（拿不到句柄
    或读取为空）时，才降级回原 UIA 方案（展开 + 遍历 ListItem），保证
    与旧逻辑行为一致、结果完全相同。

    Returns:
        候选项文本列表（已去除空白并去重），找不到或无法读取返回None
    """
    # ── 快速路径：win32 消息直接读列表（不展开下拉层）──
    try:
        hwnd = _get_combo_hwnd(dlg, auto_id)
        if hwnd is not None:
            items = _read_combobox_items_win32(hwnd)
            if items:
                print(f"  [INFO] 已读取到 {len(items)} 个下拉候选项(win32快速路径)")
                return items
    except Exception as e:
        print(f"  [WARN] win32 读取下拉候选项失败，降级到 UIA: {e}")

    # ── 降级路径：原 UIA 方案 ──
    return _get_combobox_items_uia(dlg, auto_id)


def get_edit_value_by_id(dlg, auto_id: str) -> Optional[int]:
    """通过 auto_id 获取数值输入框(Edit)的值（Win32 WM_GETTEXT，速度快）。

    Returns:
        整数值，找不到返回None
    """
    try:
        hwnd = _get_control_hwnd(dlg, auto_id)
        if hwnd is not None:
            text = _get_window_text(hwnd).strip().replace(",", "")
            return int(text) if text.isdigit() else None
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
                text = edit.get_value().strip().replace(",", "")
                return int(text) if text.isdigit() else None
            except Exception:
                time.sleep(0.5)
        return None
    except Exception as e:
        print(f"  [WARN] 获取数值框(auto_id={auto_id})失败: {e}")
        return None


def get_radiobutton_state_by_id(dlg, auto_id: str) -> Optional[bool]:
    """通过 auto_id 获取 RadioButton 的选中状态（Win32 BM_GETCHECK，速度快）。

    Returns:
        True=已选中, False=未选中, None=找不到
    """
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
    “未启用”）。为能验证默认参数，采取与自动追单类似的策略：
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

    for key_name in BOTTOM_CHECK_KEYS:
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


def collect_current_settings(dlg) -> dict:
    """读取当前面板全部设置值，返回与 STANDARD_VALUES 同构的字典。

    供“抓取自定义标准”脚本把当前客户端界面值采集为新的比对标准。
    逻辑与下方 test_* 函数一致（含为读取数值临时启用再恢复），但只返回字典、
    不写报告、不改任何设置。
    """
    data: dict = {}

    # 一、股票买卖委托价格跟盘设置
    buy_checked = get_checkbox_state_by_id(dlg, AUTO_ID["买入缺省价"])
    data["买入缺省价_勾选"] = bool(buy_checked) if buy_checked is not None else False
    if buy_checked:
        data["买入缺省价_选项"] = get_combobox_selection_by_id(dlg, AUTO_ID["买入缺省价_下拉"]) or ""
        buy_items = get_combobox_items_by_id(dlg, AUTO_ID["买入缺省价_下拉"])
        data["买入缺省价_下拉_选项列表"] = buy_items or []
    else:
        data["买入缺省价_选项"] = None
        data["买入缺省价_下拉_选项列表"] = []

    sell_checked = get_checkbox_state_by_id(dlg, AUTO_ID["卖出缺省价"])
    data["卖出缺省价_勾选"] = bool(sell_checked) if sell_checked is not None else False
    if sell_checked:
        data["卖出缺省价_选项"] = get_combobox_selection_by_id(dlg, AUTO_ID["卖出缺省价_下拉"]) or ""
        sell_items = get_combobox_items_by_id(dlg, AUTO_ID["卖出缺省价_下拉"])
        data["卖出缺省价_下拉_选项列表"] = sell_items or []
    else:
        data["卖出缺省价_选项"] = None
        data["卖出缺省价_下拉_选项列表"] = []

    # 二、大单自动分单设置
    def _collect_split(key: str, cb_id: str, val_id: str):
        initial = get_checkbox_state_by_id(dlg, cb_id)
        data[f"{key}_勾选"] = bool(initial) if initial is not None else False
        need_restore = False
        if not initial:
            set_checkbox_by_id(dlg, cb_id, True)
            need_restore = True
            time.sleep(0.6)
        val = get_edit_value_by_id(dlg, val_id)
        data[f"{key}_数值"] = val if val is not None else 0
        if need_restore:
            set_checkbox_by_id(dlg, cb_id, False)
            time.sleep(0.4)

    _collect_split("股票拆单", AUTO_ID["股票拆单"], AUTO_ID["股票拆单_数值"])
    _collect_split("基金拆单", AUTO_ID["基金拆单"], AUTO_ID["基金拆单_数值"])

    # 三、委托数量设置
    def _collect_qty(key: str, cb_id: str, sub):
        initial = get_checkbox_state_by_id(dlg, cb_id)
        data[f"{key}_勾选"] = bool(initial) if initial is not None else False
        need_restore = False
        if not initial:
            set_checkbox_by_id(dlg, cb_id, True)
            need_restore = True
            time.sleep(0.6)
        sub()
        if need_restore:
            set_checkbox_by_id(dlg, cb_id, False)
            time.sleep(0.4)

    def _buy_sell_sub(prefix: str):
        option = None
        for rid in (AUTO_ID[f"{prefix}_确定数量"], AUTO_ID[f"{prefix}_全部数量"],
                    AUTO_ID[f"{prefix}_上一次交易数量"]):
            if get_radiobutton_state_by_id(dlg, rid):
                option = RADIO_NAMES.get(rid)
                break
        data[f"{prefix}_选项"] = option or ""
        qty = get_edit_value_by_id(dlg, AUTO_ID[f"{prefix}_数值"])
        data[f"{prefix}_数值"] = qty if qty is not None else 0

    def _trade_sub(prefix: str):
        qty = get_edit_value_by_id(dlg, AUTO_ID[f"{prefix}_数值"])
        data[f"{prefix}_数值"] = qty if qty is not None else 0

    _collect_qty("股票买入自动填入数量", AUTO_ID["股票买入自动填入数量"],
                 lambda: _buy_sell_sub("股票买入"))
    _collect_qty("股票卖出自动填入数量", AUTO_ID["股票卖出自动填入数量"],
                 lambda: _buy_sell_sub("股票卖出"))
    _collect_qty("期权交易自动填入数量", AUTO_ID["期权交易自动填入数量"],
                 lambda: _trade_sub("期权交易"))
    _collect_qty("期货交易自动填入数量", AUTO_ID["期货交易自动填入数量"],
                 lambda: _trade_sub("期货交易"))

    # 四、底部复选框
    for key in BOTTOM_CHECK_KEYS:
        st = get_checkbox_state_by_id(dlg, AUTO_ID[key])
        data[key] = bool(st) if st is not None else False

    return data


def main():
    """主流程"""
    print("=" * 60)
    print("交易系统设置 - 委托设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult(PANEL_NAME)
    hwnd = None
    dlg = None

    try:
        # 1. 倒计时（让用户把焦点切到钱龙窗口）
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

        # 4. 切换到委托设置面板
        if not switch_settings_panel_compat(dlg, PANEL_NAME):
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
    finally:
        # 无论正常完成、异常还是用户中断，都执行安全收尾；公共函数只向设置窗口的
        # 明确句柄发送 WM_CLOSE，不再使用可能落到主窗口上的 Alt+F4。
        if dlg is not None:
            keep_open = os.environ.get("GUI_NEXT_CATEGORY", "") == "交易系统设置"
            close_ok = close_settings_dialog(
                dlg, keep_open=keep_open, main_hwnd=hwnd
            )
            if not close_ok:
                print("[WARN] 交易系统设置窗口未正常关闭，请确认后再执行后续非设置类任务")


if __name__ == "__main__":
    main()
