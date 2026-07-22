# -*- coding: utf-8 -*-
"""
钱龙组合/拆分申报 - 全自动轮询脚本（Win32 兼容版）
================================================
相比原 pywinauto 版本，本版本优先用 Win32 消息操作控件，并兼容 Win10/Win11：

  - 下拉框(交易所/策略)：优先用 CB_GETLBTEXT 只读建立“文本→索引”映射；原生读取
    不可用时才用 UIA 遍历，并恢复原选择。之后用 CB_SETCURSEL + CBN_SELCHANGE 选择。
  - 数据列表(SysListView32)：LVM_GETITEMCOUNT 即时判断是否有数据，不再 UIA 遍历子元素。
  - 选中第一条：依次尝试 LVM_SETITEMSTATE、控件定向键盘消息、UIA；任一路径均需
    回读确认首行确实被选中，不使用坐标点击。
  - 弹窗(组合申报)：EnumWindows 定位 + WM_SETTEXT 填数量（不再每 0.2s 连 uia）。
  - 弹窗仅匹配目标客户端进程；警告默认不确认。

通过 GUI_COMBINATION_ACTION=combine/split 选择组合或拆分流程。组合模式遍历
全部交易所和策略，拆分模式遍历全部交易所；每个有数据的项目提交第一条记录。
"""
import sys
import os
import time
import ctypes
import struct
import win32gui
import win32con
import win32process
import uiautomation as auto

_here = os.path.dirname(os.path.abspath(__file__))
for _d in (_here, os.path.dirname(_here), os.path.dirname(os.path.dirname(_here))):
    if os.path.isdir(os.path.join(_d, "core")) and _d not in sys.path:
        sys.path.insert(0, _d)
# 统一复用多客户端窗口和菜单逻辑；直接运行时也会按窗口标题识别客户端。
from core.window import (
    find_window as find_main_window,
    activate_window as activate_main_window,
    switch_panel,
)
from core.native_tree import RemoteProcessMemory
from pywinauto import Application

# ====================== 可配置参数 ======================
WINDOW_KEY = "钱龙模拟期权宝"        # 窗口标题关键字
ACTION = os.environ.get("GUI_COMBINATION_ACTION", "combine").strip().lower()
if ACTION not in ("combine", "split"):
    raise ValueError(f"GUI_COMBINATION_ACTION 无效: {ACTION!r}")
IS_COMBINE = ACTION == "combine"
ACTION_NAME = "组合" if IS_COMBINE else "拆分"
DIALOG_TITLE = f"{ACTION_NAME}申报"
PANEL_PATH = rf"\组合申报\{DIALOG_TITLE}"
ORDER_QTY = int(os.environ.get("GUI_ORDER_QTY", "1"))   # 委托数量（可由 GUI 参数覆盖）
COUNTDOWN = int(os.environ.get("GUI_COUNTDOWN", "3"))   # 操作前倒计时秒数
SKIP_NAV = os.environ.get("GUI_SKIP_NAV") == "1"        # 已在该页面时跳过切换面板
# ========================================================

# 所有交易所
EXCHANGES = ["上证", "深证"]

# 所有策略类型
STRATEGIES = [
    "认购牛市价差",
    "认购熊市价差",
    "宽跨式空头",
    "跨式空头",
    "认沽牛市价差",
    "认沽熊市价差",
]

# 组合至少需要两条腿；拆分列表一条记录即可执行。
MIN_DATA_ROWS = 2 if IS_COMBINE else 1

# ---------------- Win32 消息常量 ----------------
CB_GETCOUNT = 0x0146
CB_SETCURSEL = 0x014E
CBN_SELCHANGE = 1
WM_COMMAND = 0x0111

LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4      # 0x1004
LVM_GETNEXTITEM = LVM_FIRST + 12      # 0x100C
LVM_ENSUREVISIBLE = LVM_FIRST + 19    # 0x1013
LVM_SETITEMSTATE = LVM_FIRST + 43     # 0x102B
LVM_SETSELECTIONMARK = LVM_FIRST + 67 # 0x1043
LVIS_SELECTED = 0x0002
LVIS_FOCUSED = 0x0001
LVIF_STATE = 0x0008
LVNI_SELECTED = 0x0002

BM_CLICK = 0x00F5
WM_SETTEXT = 0x000C
WM_CHAR = 0x0102
WM_CLEAR = 0x0303
EM_SETSEL = 0x00B1
GWL_STYLE = -16
BS_DEFPUSHBUTTON = 0x0001

DIALOG_KEYWORDS = (DIALOG_TITLE, "警告", "确认", "提示")
AFFIRMATIVE_IDS = (1, 6, 5051)  # IDOK / IDYES / 客户端自定义“确定”
CANCEL_IDS = (2, 7, 5052)       # IDCANCEL / IDNO / 客户端自定义“取消”


class DialogCleanupError(RuntimeError):
    """交易弹窗无法关闭；继续遍历可能导致客户端消息循环卡死。"""


# ============================================================
# 窗口 / 控件查找
# ============================================================
def find_window(keyword):
    """按关键字找主窗口（win32gui 实现，避免 pywinauto 启动开销）。"""
    found = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if keyword in title and "GUI自动化工具" not in title:
            found.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    if not found:
        raise RuntimeError(f"未找到包含'{keyword}'的窗口，请确认软件已启动")
    return found[0]


def activate_window(hwnd):
    """置前并显示窗口，返回 pywinauto 包装对象（供 switch_panel 兼容使用）。"""
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    except Exception:
        pass
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.2)
    app = Application(backend="uia").connect(handle=hwnd)
    return app.window(handle=hwnd)


def find_top_windows_with(substrings, target_pid=None):
    """返回目标进程中标题包含任一子串的顶层窗口句柄列表。"""
    found = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        if target_pid is not None and _dlg_pid(hwnd) != target_pid:
            return True
        title = win32gui.GetWindowText(hwnd)
        if "GUI自动化工具" in title:
            return True
        if any(s in title for s in substrings):
            found.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return found


def find_visible_child(hwnd, ctrl_id):
    """在 hwnd 的全部后代中，找指定控件 ID 且可见的控件（同 ID 在多面板复用，必须取可见的那个）。"""
    out = []
    def cb(h, _):
        try:
            if win32gui.GetDlgCtrlID(h) == ctrl_id and win32gui.IsWindowVisible(h):
                out.append(h)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(hwnd, cb, None)
    if out:
        return out[0]
    out = []
    def cb2(h, _):
        try:
            if win32gui.GetDlgCtrlID(h) == ctrl_id:
                out.append(h)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(hwnd, cb2, None)
    return out[0] if out else None


# ============================================================
# 下拉框：建立 文本→索引 映射（UIA 只读一次），之后纯 win32gui 选择
# ============================================================
_combo_maps = {}   # combo_hwnd -> {text: index}
_combo_sources = {}


def _send_timeout(hwnd, message, wparam=0, lparam=0, timeout=3000):
    """发送有超时保护的消息，并统一返回 LRESULT。"""
    response = win32gui.SendMessageTimeout(
        hwnd, message, wparam, lparam, win32con.SMTO_ABORTIFHUNG, timeout
    )
    return response[1] if isinstance(response, tuple) else response


def _read_combo_value(combo_hwnd):
    try:
        return auto.ControlFromHandle(combo_hwnd).GetValuePattern().Value
    except Exception:
        return ""


def _native_combo_map(combo_hwnd):
    """不改变选择，直接读取标准 ComboBox 的全部候选项。"""
    count = int(_send_timeout(combo_hwnd, CB_GETCOUNT))
    if count <= 0:
        return {}

    user32 = ctypes.WinDLL("user32", use_last_error=True)
    user32.SendMessageW.argtypes = [
        ctypes.c_void_p, ctypes.c_uint, ctypes.c_size_t, ctypes.c_ssize_t
    ]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    CB_GETLBTEXT = 0x0148
    CB_GETLBTEXTLEN = 0x0149
    result = {}
    for index in range(count):
        length = user32.SendMessageW(
            combo_hwnd, CB_GETLBTEXTLEN, index, 0
        )
        if length < 0:
            return {}
        buffer = ctypes.create_unicode_buffer(max(int(length) + 1, 2))
        copied = user32.SendMessageW(
            combo_hwnd, CB_GETLBTEXT, index, ctypes.addressof(buffer)
        )
        if copied < 0 or not buffer.value.strip():
            return {}
        result[buffer.value.strip()] = index
    return result


def build_combo_map(combo_hwnd):
    """建立文本到索引映射；优先只读，降级遍历时恢复原选择。"""
    if combo_hwnd in _combo_maps:
        return _combo_maps[combo_hwnd]

    native = _native_combo_map(combo_hwnd)
    if native:
        _combo_maps[combo_hwnd] = native
        _combo_sources[combo_hwnd] = "原生只读"
        return native

    n = int(_send_timeout(combo_hwnd, CB_GETCOUNT))
    parent = win32gui.GetParent(combo_hwnd)
    cid = win32gui.GetDlgCtrlID(combo_hwnd)
    original = int(_send_timeout(combo_hwnd, 0x0147))  # CB_GETCURSEL
    m = {}
    try:
        for i in range(n):
            _send_timeout(combo_hwnd, CB_SETCURSEL, i, 0)
            _send_timeout(
                parent, WM_COMMAND, (CBN_SELCHANGE << 16) | cid, combo_hwnd
            )
            time.sleep(0.05)
            v = _read_combo_value(combo_hwnd).strip()
            if v:
                m[v] = i
    finally:
        if 0 <= original < n:
            _send_timeout(combo_hwnd, CB_SETCURSEL, original, 0)
            _send_timeout(
                parent, WM_COMMAND,
                (CBN_SELCHANGE << 16) | cid, combo_hwnd
            )
    _combo_maps[combo_hwnd] = m
    _combo_sources[combo_hwnd] = "UIA遍历后恢复"
    return m


def combo_select(combo_hwnd, text):
    """选中下拉框指定文本（用已建好的映射取索引，再 CB_SETCURSEL + 通知父窗口）。"""
    m = build_combo_map(combo_hwnd)
    idx = m.get(text, -1)
    if idx < 0:
        return False
    _send_timeout(combo_hwnd, CB_SETCURSEL, idx, 0)
    parent = win32gui.GetParent(combo_hwnd)
    cid = win32gui.GetDlgCtrlID(combo_hwnd)
    _send_timeout(
        parent, WM_COMMAND, (CBN_SELCHANGE << 16) | cid, combo_hwnd
    )
    time.sleep(0.1)
    selected = int(_send_timeout(combo_hwnd, 0x0147))  # CB_GETCURSEL
    if selected != idx:
        return False
    actual = _read_combo_value(combo_hwnd).strip()
    return not actual or actual == text


# ============================================================
# 列表（SysListView32）操作
# ============================================================
def list_count(list_hwnd):
    return int(_send_timeout(list_hwnd, LVM_GETITEMCOUNT, 0, 0))


def _pack_lvitem_state(target_bits, iitem, state, state_mask):
    """打包只含状态字段的 LVITEM；这些字段在32/64位偏移相同。"""
    if target_bits not in (32, 64):
        raise ValueError(f"不支持的目标进程位数: {target_bits}")
    buf = bytearray(24)
    struct.pack_into("<I", buf, 0, LVIF_STATE)
    struct.pack_into("<i", buf, 4, iitem)
    struct.pack_into("<I", buf, 12, state)
    struct.pack_into("<I", buf, 16, state_mask)
    return bytes(buf)


def _selected_list_index(list_hwnd):
    result = int(_send_timeout(
        list_hwnd, LVM_GETNEXTITEM, -1, LVNI_SELECTED
    ))
    if result in (0xFFFFFFFF, 0xFFFFFFFFFFFFFFFF):
        return -1
    return result


def _list_select_first_native(list_hwnd):
    """使用远程 LVITEM 选中首行；适用于 Win10/同权限进程。"""
    memory = RemoteProcessMemory(list_hwnd)
    try:
        packed = _pack_lvitem_state(
            memory.target_bits, 0, LVIS_SELECTED | LVIS_FOCUSED,
                                    LVIS_SELECTED | LVIS_FOCUSED)
        address = memory.allocate(len(packed))
        memory.write(address, packed)
        _send_timeout(list_hwnd, LVM_SETITEMSTATE, 0, address)
        if _selected_list_index(list_hwnd) != 0:
            raise RuntimeError("原生消息发送后首行没有成为选中项")
    finally:
        memory.close()


def _list_select_first_uia(main_hwnd, list_hwnd):
    """远程内存不可用时，用 UIA 精确选择可见列表的第一行。"""
    app = Application(backend="uia").connect(handle=main_hwnd, timeout=2)
    main = app.window(handle=main_hwnd)
    try:
        list_box = app.window(handle=list_hwnd)
        rows = list_box.children(control_type="ListItem")
    except Exception:
        list_box = main.child_window(auto_id="1229", control_type="List")
        list_box.wait("ready", timeout=5)
        rows = list_box.children(control_type="ListItem")

    visible_rows = []
    for row in rows:
        try:
            rect = row.rectangle()
            if rect.right > rect.left and rect.bottom > rect.top:
                visible_rows.append(row)
        except Exception:
            continue
    if not visible_rows:
        raise RuntimeError("UIA 未暴露任何可见列表行")
    try:
        visible_rows[0].select()
    except Exception:
        visible_rows[0].click_input()
    time.sleep(0.15)
    selected = _selected_list_index(list_hwnd)
    if selected != 0:
        raise RuntimeError(f"UIA 选择后客户端报告选中行={selected}")


def _list_select_first_keyboard(main_hwnd, list_hwnd):
    """用控件定向键盘消息选中首行，不需要远程内存或坐标。"""
    if win32gui.GetClassName(list_hwnd) != "SysListView32":
        raise RuntimeError("目标控件不是 SysListView32")
    if not win32gui.IsWindowVisible(list_hwnd):
        raise RuntimeError("目标列表当前不可见")
    if list_count(list_hwnd) <= 0:
        raise RuntimeError("目标列表没有可选择的行")
    try:
        win32gui.SetForegroundWindow(main_hwnd)
    except Exception:
        pass
    _send_timeout(list_hwnd, LVM_ENSUREVISIBLE, 0, 0)
    for key in (win32con.VK_HOME, win32con.VK_DOWN):
        _send_timeout(list_hwnd, win32con.WM_SETFOCUS, 0, 0)
        _send_timeout(list_hwnd, win32con.WM_KEYDOWN, key, 0)
        _send_timeout(list_hwnd, win32con.WM_KEYUP, key, 0)
        time.sleep(0.12)
        if _selected_list_index(list_hwnd) == 0:
            return

    _send_timeout(list_hwnd, LVM_SETSELECTIONMARK, 0, 0)
    _send_timeout(list_hwnd, win32con.WM_SETFOCUS, 0, 0)
    _send_timeout(list_hwnd, win32con.WM_KEYDOWN, win32con.VK_SPACE, 0)
    _send_timeout(list_hwnd, win32con.WM_KEYUP, win32con.VK_SPACE, 0)
    time.sleep(0.12)
    if _selected_list_index(list_hwnd) != 0:
        raise RuntimeError("Home/Down/Space消息均未确认选中首行")


def list_select_first(main_hwnd, list_hwnd):
    """按远程Win32、无指针键盘消息、UIA的顺序选择首行。"""
    try:
        _list_select_first_native(list_hwnd)
        return {"ok": True, "method": "Win32远程内存", "error": ""}
    except Exception as native_error:
        print(f"[INFO] 远程内存列表选择不可用，改用键盘消息: {native_error}")
        try:
            _list_select_first_keyboard(main_hwnd, list_hwnd)
            return {"ok": True, "method": "Win32键盘消息", "error": ""}
        except Exception as keyboard_error:
            print(f"[INFO] 键盘消息选择不可用，改用UIA: {keyboard_error}")
            try:
                _list_select_first_uia(main_hwnd, list_hwnd)
                return {"ok": True, "method": "UIA", "error": ""}
            except Exception as uia_error:
                return {
                    "ok": False,
                    "method": "",
                    "error": (
                        f"远程Win32={native_error}; "
                        f"键盘消息={keyboard_error}; UIA={uia_error}"
                    ),
                }


# ============================================================
# 按钮 / 弹窗
# ============================================================
def enum_buttons(dlg_hwnd):
    out = []
    def cb(h, _):
        try:
            if win32gui.GetClassName(h) == "Button":
                out.append(h)
        except Exception:
            pass
        return True
    win32gui.EnumChildWindows(dlg_hwnd, cb, None)
    return out


def _usable_button(button):
    return bool(
        button
        and win32gui.IsWindow(button)
        and win32gui.IsWindowVisible(button)
        and win32gui.IsWindowEnabled(button)
    )


def _normalized_button_text(button):
    text = (win32gui.GetWindowText(button) or "").replace("&", "").strip()
    return text.split("(", 1)[0].strip().lower()


def _find_dialog_button(dlg_hwnd, affirmative=True):
    """查找肯定/取消按钮；兼容国泰海通的自绘“确定(Y)”按钮。"""
    ids = AFFIRMATIVE_IDS if affirmative else CANCEL_IDS
    for ctrl_id in ids:
        try:
            button = win32gui.GetDlgItem(dlg_hwnd, ctrl_id)
        except Exception:
            button = 0
        if _usable_button(button):
            return button

    positive = ("确定", "确认", "是", "继续", "提交", "ok", "yes")
    negative = ("取消", "否", "关闭", "cancel", "no")
    prefixes = positive if affirmative else negative
    for button in enum_buttons(dlg_hwnd):
        if _usable_button(button) and _normalized_button_text(button).startswith(prefixes):
            return button

    if affirmative:
        for button in enum_buttons(dlg_hwnd):
            if not _usable_button(button):
                continue
            try:
                style = win32gui.GetWindowLong(button, GWL_STYLE)
            except Exception:
                continue
            if (style & 0xF) == BS_DEFPUSHBUTTON:
                return button
    return 0


def click_default_button(dlg_hwnd):
    """兼容函数：点击明确的肯定按钮，包括自绘 IDOK。"""
    button = _find_dialog_button(dlg_hwnd, affirmative=True)
    if not button:
        return False
    win32gui.PostMessage(button, BM_CLICK, 0, 0)
    return True


def _wait_for_dialog_button(dlg_hwnd, affirmative=True, timeout=2.5):
    """等待异步创建的自绘按钮进入可用状态。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not win32gui.IsWindow(dlg_hwnd) or not win32gui.IsWindowVisible(dlg_hwnd):
            return 0
        button = _find_dialog_button(dlg_hwnd, affirmative=affirmative)
        if button:
            return button
        time.sleep(0.1)
    return 0


def _dlg_pid(hwnd):
    """返回 hwnd 所在进程 PID（用于只处理目标软件的弹窗）。"""
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return pid
    except Exception:
        return None


def has_control(hwnd, ctrl_id):
    """判断对话框是否含某控件 ID（GetDlgItem 找不到会抛 1421，故包 try）。"""
    try:
        return win32gui.GetDlgItem(hwnd, ctrl_id) is not None
    except Exception:
        return False


def find_all_dialogs(target_pid=None):
    """返回目标进程的 #32770 顶层对话框句柄。"""
    found = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if "GUI自动化工具" in title:
            return True
        if target_pid is not None and _dlg_pid(hwnd) != target_pid:
            return True
        if win32gui.GetClassName(hwnd) == "#32770":
            found.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return found


def _relevant_dialogs(main_hwnd):
    """只返回目标客户端进程中可见的本次申报链弹窗。"""
    target_pid = _dlg_pid(main_hwnd)
    dialogs = []
    for dialog in find_all_dialogs(target_pid):
        title = win32gui.GetWindowText(dialog) or ""
        if any(keyword in title for keyword in DIALOG_KEYWORDS):
            dialogs.append(dialog)
    return dialogs


def click_dialog_button(dlg, ctrl_id):
    """点击对话框内指定控件 ID 的按钮（找不到返回 False）。

    用 PostMessage 异步点击，避免按钮打开/结束模态对话框时 SendMessage 永久阻塞。
    """
    try:
        btn = win32gui.GetDlgItem(dlg, ctrl_id)
    except Exception:
        return False
    if not _usable_button(btn):
        return False
    win32gui.PostMessage(btn, BM_CLICK, 0, 0)
    return True


def get_edit_text(edit_hwnd):
    """读取 Edit 控件文本（只读，不触发模态，安全）。"""
    try:
        n = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
    except Exception:
        return ""
    buf = ctypes.create_unicode_buffer(n + 1)
    try:
        win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXT, n + 1, buf)
    except Exception:
        return ""
    return buf.value


def fill_qty_in_dialog(main_hwnd, timeout=8):
    """定位当前申报数量弹窗，填写委托数量，返回弹窗句柄（失败返回 None）。

    注意：部分标题含申报名称的窗口（如提示/结果框）并不含委托数量编辑框 9057，
    win32gui.GetDlgItem 对不存在的控件 ID 会抛 1421 异常（而非返回 0），
    因此必须放在 try 中并跳过；只认真正的 #32770 对话框，避免误匹配主窗口等。
    写完后回读校验，确保数量确实写入（避免点到"确定"时数量仍为空）。
    """
    target_pid = _dlg_pid(main_hwnd)
    end = time.time() + timeout
    while time.time() < end:
        for dlg in find_top_windows_with([DIALOG_TITLE], target_pid):
            if win32gui.GetClassName(dlg) != "#32770":
                continue
            try:
                edit = win32gui.GetDlgItem(dlg, 9057)
            except Exception:
                continue
            if not edit or win32gui.GetClassName(edit) != "Edit":
                continue
            value = str(ORDER_QTY)
            # 先按真实键盘输入路径发送给 Edit，使 MFC 的校验/通知逻辑收到变化；
            # 若客户端不接受 WM_CHAR，再降级为 WM_SETTEXT。
            try:
                _send_timeout(edit, EM_SETSEL, 0, -1)
                _send_timeout(edit, WM_CLEAR, 0, 0)
                for char in value:
                    _send_timeout(edit, WM_CHAR, ord(char), 1)
            except Exception:
                pass
            if get_edit_text(edit) != value:
                try:
                    win32gui.SendMessage(edit, WM_SETTEXT, 0, value)
                except Exception:
                    continue
            # 校验是否真的写入（自绘/带串编辑框偶发需要重试）。
            if get_edit_text(edit) != value:
                time.sleep(0.1)
                try:
                    win32gui.SendMessage(edit, WM_SETTEXT, 0, value)
                except Exception:
                    continue
            if get_edit_text(edit) == value:
                print(f"[OK] {DIALOG_TITLE}弹窗: 委托数量已设为 {ORDER_QTY}")
                return dlg
        time.sleep(0.15)
    print(f"[WARN] 等待'{DIALOG_TITLE}'弹窗超时({timeout}s)")
    return None


def cancel_dialog(dialog, timeout=3):
    """取消指定对话框；按钮无效时发送 WM_CLOSE，并验证窗口关闭。"""
    if not dialog or not win32gui.IsWindow(dialog):
        return True
    cancel = _find_dialog_button(dialog, affirmative=False)
    if cancel:
        win32gui.PostMessage(cancel, BM_CLICK, 0, 0)
    else:
        win32gui.PostMessage(dialog, win32con.WM_CLOSE, 0, 0)
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not win32gui.IsWindow(dialog) or not win32gui.IsWindowVisible(dialog):
            return True
        time.sleep(0.1)
    # 某些自绘按钮只把焦点切到“取消”，再补一次窗口关闭。
    win32gui.PostMessage(dialog, win32con.WM_CLOSE, 0, 0)
    deadline = time.time() + 1.0
    while time.time() < deadline:
        if not win32gui.IsWindow(dialog) or not win32gui.IsWindowVisible(dialog):
            return True
        time.sleep(0.1)
    return False


def _wait_dialog_transition(main_hwnd, dialog, known_dialogs=None, timeout=1.4):
    """等待当前弹窗关闭，或等待同一申报链出现新的弹窗。"""
    known = set(known_dialogs or (dialog,))
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not win32gui.IsWindow(dialog) or not win32gui.IsWindowVisible(dialog):
            return True
        if any(candidate not in known for candidate in _relevant_dialogs(main_hwnd)):
            return True
        time.sleep(0.1)
    return False


def confirm_dialog(main_hwnd, dialog, attempts=3):
    """点击肯定按钮并验证状态变化；国泰海通实测偶发需重复点击。"""
    title = win32gui.GetWindowText(dialog) or ""
    for attempt in range(1, attempts + 1):
        if not win32gui.IsWindow(dialog) or not win32gui.IsWindowVisible(dialog):
            return True
        button = _wait_for_dialog_button(
            dialog, affirmative=True, timeout=2.5 if attempt == 1 else 0.8
        )
        if not button:
            print(f"[ERROR] 弹窗没有可确认按钮: {title!r}")
            return False
        known_dialogs = set(_relevant_dialogs(main_hwnd))
        win32gui.PostMessage(button, BM_CLICK, 0, 0)
        if _wait_dialog_transition(
            main_hwnd, dialog, known_dialogs=known_dialogs
        ):
            print(
                f"[OK] 弹窗已确认并发生状态变化: {title!r} "
                f"(尝试 {attempt}/{attempts})"
            )
            return True
        print(f"[WARN] 弹窗确认后未关闭，准备重试: {title!r} ({attempt}/{attempts})")
    return False


def handle_dialogs(main_hwnd, timeout=8.0, quiet_period=1.2):
    """等待并确认后续警告/确认/提示，直到连续 quiet_period 秒无弹窗。"""
    deadline = time.time() + timeout
    quiet_since = None
    while time.time() < deadline:
        dialogs = _relevant_dialogs(main_hwnd)
        followups = [dialog for dialog in dialogs if not has_control(dialog, 9057)]
        if followups:
            quiet_since = None
            dialog = followups[0]
            title = win32gui.GetWindowText(dialog) or ""
            if not confirm_dialog(main_hwnd, dialog, attempts=3):
                print(f"[ERROR] 无法确认后续弹窗: {title!r}")
                return False
            continue

        # 数量框仍可见说明提交按钮没有真正生效，不能假报成功。
        if any(has_control(dialog, 9057) for dialog in dialogs):
            print(f"[ERROR] {DIALOG_TITLE}数量弹窗仍然可见")
            return False

        if quiet_since is None:
            quiet_since = time.time()
        elif time.time() - quiet_since >= quiet_period:
            return True
        time.sleep(0.1)
    print(f"[ERROR] 等待{DIALOG_TITLE}后续弹窗结束超时")
    return False


def abort_transaction_dialogs(main_hwnd, timeout=6.0):
    """关闭本次申报链全部可见弹窗，并验证没有残留。"""
    deadline = time.time() + timeout
    while time.time() < deadline:
        dialogs = _relevant_dialogs(main_hwnd)
        if not dialogs:
            return True
        # 先关嵌套警告/确认，再关底层数量窗口。
        dialogs.sort(key=lambda dialog: has_control(dialog, 9057))
        dialog = dialogs[0]
        title = win32gui.GetWindowText(dialog) or ""
        if not cancel_dialog(dialog, timeout=1.0):
            print(f"[WARN] 常规取消失败，强制关闭弹窗: {title!r}")
            win32gui.PostMessage(dialog, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.2)
    remaining = [
        (dialog, win32gui.GetWindowText(dialog) or "")
        for dialog in _relevant_dialogs(main_hwnd)
    ]
    if remaining:
        print(f"[ERROR] 交易弹窗清理失败，仍残留: {remaining!r}")
        return False
    return True


def close_leftover_qty_dialogs(main_hwnd):
    """启动前清理残留交易弹窗；无法清理时停止运行。"""
    dialogs = _relevant_dialogs(main_hwnd)
    if not dialogs:
        return
    print(f"[INFO] 检测到 {len(dialogs)} 个残留交易弹窗，开始取消")
    if not abort_transaction_dialogs(main_hwnd):
        raise DialogCleanupError("启动前无法清理残留交易弹窗")
    print("[OK] 残留交易弹窗已全部关闭")


# ============================================================
# 业务流程
# ============================================================
def select_exchange(main_hwnd, exchange):
    combo = find_visible_child(main_hwnd, 9059)
    if not combo:
        print("[WARN] 未找到交易所下拉框")
        return False
    if combo_select(combo, exchange):
        print(f"[OK] 已选择交易所: {exchange}")
        return True
    print(f"[WARN] 交易所选择失败: {exchange}")
    return False


def select_strategy(main_hwnd, strategy):
    combo = find_visible_child(main_hwnd, 9040)
    if not combo:
        print("[WARN] 未找到策略下拉框")
        return False
    if combo_select(combo, strategy):
        print(f"[OK] 已选择策略: {strategy}")
        return True
    print(f"[--] 策略选择失败，跳过: {strategy}")
    return False


def check_list_has_data(main_hwnd, wait=0.4, timeout=3.0):
    list_hwnd = find_visible_child(main_hwnd, 1229)
    if not list_hwnd:
        return False
    time.sleep(wait)
    deadline = time.time() + timeout
    last = None
    stable = 0
    cnt = 0
    while time.time() < deadline:
        cnt = list_count(list_hwnd)
        if cnt == last:
            stable += 1
        else:
            last = cnt
            stable = 1
        if stable >= 3:
            break
        time.sleep(0.2)
    if stable < 3:
        print("[WARN] 列表行数在超时前没有稳定，拒绝申报")
        return False
    if cnt >= MIN_DATA_ROWS:
        print(f"[OK] 列表行数已稳定({cnt} 条)，可以继续检查")
        return True
    print(f"[--] 列表无足够数据({cnt} 条)，跳过")
    return False


def click_first_list_row_and_submit(main_hwnd):
    list_hwnd = find_visible_child(main_hwnd, 1229)
    if not list_hwnd:
        return False
    if list_count(list_hwnd) < MIN_DATA_ROWS:
        return False
    print(f"[OK] 列表共 {list_count(list_hwnd)} 条数据，准备选中第一条")
    selection = list_select_first(main_hwnd, list_hwnd)
    if not selection["ok"]:
        print(f"[ERROR] 无法安全选中列表首行: {selection['error']}")
        return False
    print(f"[OK] 已通过{selection['method']}选中列表首行")
    time.sleep(0.2)

    action_btn = find_visible_child(main_hwnd, 9058)
    if not action_btn:
        print(f"[WARN] 未找到'{ACTION_NAME}'按钮")
        return False
    actual_button_text = win32gui.GetWindowText(action_btn).strip()
    if actual_button_text and actual_button_text != ACTION_NAME:
        print(
            f"[ERROR] 控件9058文字异常: 期望={ACTION_NAME!r}，"
            f"实际={actual_button_text!r}"
        )
        return False
    # 注意：动作按钮会打开模态对话框（DoModal），若用 SendMessage(BM_CLICK)
    # 会一直阻塞到对话框关闭，导致后续"填数量/点确定"永远执行不到。
    # 改用 PostMessage 异步点击，点击后立即继续，由 fill_qty_in_dialog 轮询弹窗。
    win32gui.PostMessage(action_btn, BM_CLICK, 0, 0)
    print(f"[OK] 已点击'{ACTION_NAME}'按钮（异步）")
    time.sleep(0.3)

    # 1) 等待数量弹窗并填写委托数量
    dlg = fill_qty_in_dialog(main_hwnd, timeout=8)
    if not dlg:
        print(f"[ERROR] 未找到并验证数量弹窗，停止本次{DIALOG_TITLE}")
        if not abort_transaction_dialogs(main_hwnd):
            raise DialogCleanupError("数量弹窗定位失败且无法清理交易弹窗")
        return False

    if not confirm_dialog(main_hwnd, dlg, attempts=3):
        print("[ERROR] 数量弹窗确认失败，开始取消本次申报")
        if not abort_transaction_dialogs(main_hwnd):
            raise DialogCleanupError("数量弹窗确认失败且无法清理交易弹窗")
        return False
    print(f"[OK] 数量弹窗已确认，等待{DIALOG_TITLE}后续结果")
    if not handle_dialogs(main_hwnd):
        print("[ERROR] 提交后的弹窗未正常结束，开始清理")
        if not abort_transaction_dialogs(main_hwnd):
            raise DialogCleanupError("后续弹窗异常且无法清理，停止全部遍历")
        return False
    return True


def process_item(main_hwnd, exchange, strategy=None):
    item_name = f"{exchange} | {strategy}" if strategy else exchange
    print(f"\n--- 处理: {item_name} ---")
    if IS_COMBINE:
        if not select_strategy(main_hwnd, strategy):
            return False
        time.sleep(0.3)
    if not check_list_has_data(main_hwnd):
        return False
    if click_first_list_row_and_submit(main_hwnd):
        print(f"[OK] {item_name} {DIALOG_TITLE}流程完成")
        return True
    print(f"[--] {DIALOG_TITLE}执行失败")
    return False


def countdown(seconds):
    print(f"将在 {seconds} 秒后开始，请把焦点切到钱龙软件...")
    try:
        for i in range(seconds, 0, -1):
            print(f"  {i}...", end="\r")
            time.sleep(1)
    except KeyboardInterrupt:
        raise
    print(" " * 30, end="\r")


def main():
    try:
        print(f"[INFO] 业务模式: {DIALOG_TITLE}；运行模式: 正式提交（全量遍历）")
        countdown(COUNTDOWN)
        hwnd = find_main_window(WINDOW_KEY)
        print(f"[OK] 已找到窗口，句柄 = {hex(hwnd)}")
        close_leftover_qty_dialogs(hwnd)
        win = activate_main_window(hwnd)

        if not SKIP_NAV:
            switch_panel(win, PANEL_PATH)
            time.sleep(0.5)

        # 建立下拉框 文本→索引 映射；拆分页没有策略下拉框。
        ex_combo = find_visible_child(hwnd, 9059)
        st_combo = find_visible_child(hwnd, 9040) if IS_COMBINE else None
        if not ex_combo or (IS_COMBINE and not st_combo):
            raise RuntimeError(f"{DIALOG_TITLE}页面缺少必要下拉框")
        exchange_map = build_combo_map(ex_combo)
        strategy_map = build_combo_map(st_combo) if st_combo else None
        if not exchange_map or (IS_COMBINE and not strategy_map):
            raise RuntimeError(f"无法建立{DIALOG_TITLE}下拉框映射")
        map_summary = f"交易所={_combo_sources[ex_combo]}"
        if st_combo:
            map_summary += f"，策略={_combo_sources[st_combo]}"
        print(f"[OK] 下拉框映射已建立: {map_summary}")

        exchanges = EXCHANGES
        strategies = STRATEGIES if IS_COMBINE else [None]

        total_executed = 0
        total_skipped = 0

        for exchange in exchanges:
            print(f"\n{'='*50}")
            print(f"开始处理交易所: {exchange}")
            print(f"{'='*50}")
            if not select_exchange(hwnd, exchange):
                print(f"[错误] 交易所选择失败: {exchange}")
                continue
            time.sleep(0.4)

            for strategy in strategies:
                try:
                    if process_item(hwnd, exchange, strategy):
                        total_executed += 1
                    else:
                        total_skipped += 1
                except DialogCleanupError:
                    raise
                except Exception as e:
                    print(f"[错误] 处理异常: {type(e).__name__}: {e}")
                    total_skipped += 1
                time.sleep(0.4)

        print(f"\n{'='*50}")
        print(f"=== 全部完成 ===")
        print(f"执行{DIALOG_TITLE}流程: {total_executed} 个")
        print(f"跳过(无数据): {total_skipped} 个")
        print(f"{'='*50}")
    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
