# -*- coding: utf-8 -*-
"""
钱龙组合申报 - 全自动轮询脚本（win32gui 加速版）
==============================================
相比原 pywinauto 版本，本版本用 win32gui 直接发消息操作控件，大幅提升速度：

  - 下拉框(交易所/策略)：启动时用 UIA 只读一次建立“文本→索引”映射（解决自绘/带串
    下拉框无法用 CB_GETLBTEXT 读文本的问题），之后每次选择只用
    CB_SETCURSEL + WM_COMMAND(CBN_SELCHANGE)，不再模拟鼠标键盘。
  - 数据列表(SysListView32)：LVM_GETITEMCOUNT 即时判断是否有数据，不再 UIA 遍历子元素。
  - 选中第一条：LVM_SETITEMSTATE（向目标进程注入 LVITEM），不再坐标点击。
  - 弹窗(组合申报)：EnumWindows 定位 + WM_SETTEXT 填数量（不再每 0.2s 连 uia）。
  - 确认弹窗：定位默认按钮 BM_CLICK，不再枚举 uia 元素。

代码逻辑与原版保持一致：遍历所有交易所 × 策略，无数据跳过，有数据做第一个组合申报。
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
# switch_panel 已是 win32gui 优先方案（仅 import 时加载 pywinauto，运行期不依赖）
from core.window import switch_panel
from pywinauto import Application

# ====================== 可配置参数 ======================
WINDOW_KEY = "钱龙模拟期权宝"        # 窗口标题关键字
ORDER_QTY = int(os.environ.get("GUI_ORDER_QTY", "1"))   # 委托数量（可由 GUI 参数覆盖）
COUNTDOWN = int(os.environ.get("GUI_COUNTDOWN", "3"))   # 操作前倒计时秒数
TEST_ONE = os.environ.get("GUI_TEST_ONE") == "1"        # 仅跑 1 个组合（自测用）
TEST_EXCHANGE = os.environ.get("GUI_TEST_EXCHANGE")     # 指定交易所（与 TEST_ONE 搭配）
TEST_STRATEGY = os.environ.get("GUI_TEST_STRATEGY")     # 指定策略
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

MIN_DATA_ROWS = 2   # 至少 2 条数据才视为有组合（与原版“每组合至少 2 条腿”逻辑一致）

# ---------------- Win32 消息常量 ----------------
CB_GETCOUNT = 0x0146
CB_SETCURSEL = 0x014E
CBN_SELCHANGE = 1
WM_COMMAND = 0x0111

LVM_FIRST = 0x1000
LVM_GETITEMCOUNT = LVM_FIRST + 4      # 0x1004
LVM_SETITEMSTATE = LVM_FIRST + 43     # 0x102B
LVIS_SELECTED = 0x0002
LVIS_FOCUSED = 0x0001
LVIF_STATE = 0x0008

BM_CLICK = 0x00F5
WM_SETTEXT = 0x000C
GWL_STYLE = -16
BS_DEFPUSHBUTTON = 0x0001


# ============================================================
# 远程内存（向目标进程读写，用于跨进程传结构体指针）
# ============================================================
class RemoteMem:
    def __init__(self, hwnd, target_bits):
        _, self.pid = win32process.GetWindowThreadProcessId(hwnd)
        self.h_proc = ctypes.windll.kernel32.OpenProcess(
            0x0008 | 0x0010 | 0x0020, False, self.pid)  # VM_OP|VM_READ|VM_WRITE
        if not self.h_proc:
            raise OSError(f"无法打开进程 PID={self.pid}，请以管理员权限运行脚本")
        self.target_bits = target_bits

    def alloc(self, size):
        return ctypes.windll.kernel32.VirtualAllocEx(
            self.h_proc, 0, size, 0x1000, 0x04)  # MEM_COMMIT, PAGE_READWRITE

    def free(self, addr):
        ctypes.windll.kernel32.VirtualFreeEx(self.h_proc, addr, 0, 0x8000)

    def write(self, addr, data):
        n = ctypes.c_size_t(0)
        ctypes.windll.kernel32.WriteProcessMemory(
            self.h_proc, addr, data, len(data), ctypes.byref(n))

    def close(self):
        if self.h_proc:
            ctypes.windll.kernel32.CloseHandle(self.h_proc)
            self.h_proc = None


def detect_target_bitness(hwnd):
    """检测 hwnd 所在进程位数（32/64）。"""
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    h_proc = ctypes.windll.kernel32.OpenProcess(0x1000, False, pid)  # QUERY_LIMITED
    if not h_proc:
        return struct.calcsize("P") * 8
    is_wow64 = ctypes.c_int(0)
    ctypes.windll.kernel32.IsWow64Process(h_proc, ctypes.byref(is_wow64))
    ctypes.windll.kernel32.CloseHandle(h_proc)
    python_bits = struct.calcsize("P") * 8
    if python_bits == 64:
        return 32 if is_wow64.value else 64
    return 32


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


def find_top_windows_with(substrings):
    """返回标题包含任一子串的顶层窗口句柄列表。"""
    found = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
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


def _read_combo_value(combo_hwnd):
    try:
        return auto.ControlFromHandle(combo_hwnd).GetValuePattern().Value
    except Exception:
        return ""


def build_combo_map(combo_hwnd):
    """遍历所有索引，用 CB_SETCURSEL + UIA 回读当前值，建立 文本→索引 映射。"""
    if combo_hwnd in _combo_maps:
        return _combo_maps[combo_hwnd]
    n = win32gui.SendMessage(combo_hwnd, CB_GETCOUNT, 0, 0)
    parent = win32gui.GetParent(combo_hwnd)
    cid = win32gui.GetDlgCtrlID(combo_hwnd)
    m = {}
    for i in range(n):
        win32gui.SendMessage(combo_hwnd, CB_SETCURSEL, i, 0)
        win32gui.SendMessage(parent, WM_COMMAND, (CBN_SELCHANGE << 16) | cid, combo_hwnd)
        time.sleep(0.05)
        v = _read_combo_value(combo_hwnd).strip()
        if v:
            m[v] = i
    _combo_maps[combo_hwnd] = m
    return m


def combo_select(combo_hwnd, text):
    """选中下拉框指定文本（用已建好的映射取索引，再 CB_SETCURSEL + 通知父窗口）。"""
    m = build_combo_map(combo_hwnd)
    idx = m.get(text, -1)
    if idx < 0:
        return False
    win32gui.SendMessage(combo_hwnd, CB_SETCURSEL, idx, 0)
    parent = win32gui.GetParent(combo_hwnd)
    cid = win32gui.GetDlgCtrlID(combo_hwnd)
    win32gui.SendMessage(parent, WM_COMMAND, (CBN_SELCHANGE << 16) | cid, combo_hwnd)
    return True


# ============================================================
# 列表（SysListView32）操作
# ============================================================
def list_count(list_hwnd):
    return win32gui.SendMessage(list_hwnd, LVM_GETITEMCOUNT, 0, 0)


def _pack_lvitem_state(target_bits, iitem, state, state_mask):
    """按目标位数打包 LVITEM（仅填 mask/iItem/state/stateMask）。"""
    if target_bits == 32:
        buf = bytearray(40)
        struct.pack_into("<I", buf, 0, LVIF_STATE)
        struct.pack_into("<i", buf, 4, iitem)
        struct.pack_into("<I", buf, 12, state)
        struct.pack_into("<I", buf, 16, state_mask)
        return bytes(buf)
    else:
        buf = bytearray(80)
        struct.pack_into("<I", buf, 0, LVIF_STATE)
        struct.pack_into("<i", buf, 8, iitem)
        struct.pack_into("<I", buf, 24, state)
        struct.pack_into("<I", buf, 32, state_mask)
        return bytes(buf)


def list_select_first(list_hwnd):
    """选中并聚焦列表第 0 项（LVM_SETITEMSTATE，跨进程注入 LVITEM）。

    使用 SendMessageTimeout 而非 SendMessage：若目标线程因模态弹窗等原因卡住，
    也不会让本脚本永久阻塞。"""
    bits = detect_target_bitness(list_hwnd)
    mem = RemoteMem(list_hwnd, bits)
    try:
        packed = _pack_lvitem_state(bits, 0, LVIS_SELECTED | LVIS_FOCUSED,
                                    LVIS_SELECTED | LVIS_FOCUSED)
        addr = mem.alloc(len(packed))
        mem.write(addr, packed)
        try:
            win32gui.SendMessageTimeout(
                list_hwnd, LVM_SETITEMSTATE, 0, addr, 0, 3000)
        except Exception as e:
            print(f"[WARN] 选中列表首项失败: {e}")
        mem.free(addr)
    finally:
        mem.close()


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


def click_default_button(dlg_hwnd):
    """点击对话框默认按钮（BS_DEFPUSHBUTTON），找不到则兜底点第一个按钮。"""
    for b in enum_buttons(dlg_hwnd):
        try:
            style = win32gui.GetWindowLong(b, GWL_STYLE)
        except Exception:
            continue
        if (style & 0xF) == BS_DEFPUSHBUTTON:
            # 用 PostMessage 异步点击：若该按钮会打开/结束模态对话框，
            # SendMessage 会一直阻塞到对话框关闭，导致后续逻辑无法执行。
            win32gui.PostMessage(b, BM_CLICK, 0, 0)
            return True
    btns = enum_buttons(dlg_hwnd)
    if btns:
        win32gui.PostMessage(btns[0], BM_CLICK, 0, 0)
        return True
    return False


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


def find_all_dialogs():
    """返回所有 #32770 顶层对话框句柄（排除本自动化工具窗口）。"""
    found = []
    def cb(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return True
        title = win32gui.GetWindowText(hwnd)
        if "GUI自动化工具" in title:
            return True
        if win32gui.GetClassName(hwnd) == "#32770":
            found.append(hwnd)
        return True
    win32gui.EnumWindows(cb, None)
    return found


def click_dialog_button(dlg, ctrl_id):
    """点击对话框内指定控件 ID 的按钮（找不到返回 False）。

    用 PostMessage 异步点击，避免按钮打开/结束模态对话框时 SendMessage 永久阻塞。
    """
    try:
        btn = win32gui.GetDlgItem(dlg, ctrl_id)
    except Exception:
        return False
    if not btn:
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


def fill_qty_in_dialog(timeout=8):
    """定位"组合申报"数量弹窗，填写委托数量，返回该弹窗句柄（失败返回 None）。

    注意：部分标题含"组合申报"的窗口（如提示/结果框）并不含委托数量编辑框 9057，
    win32gui.GetDlgItem 对不存在的控件 ID 会抛 1421 异常（而非返回 0），
    因此必须放在 try 中并跳过；只认真正的 #32770 对话框，避免误匹配主窗口等。
    写完后回读校验，确保数量确实写入（避免点到"确定"时数量仍为空）。
    """
    end = time.time() + timeout
    while time.time() < end:
        for dlg in find_top_windows_with(["组合申报"]):
            if win32gui.GetClassName(dlg) != "#32770":
                continue
            try:
                edit = win32gui.GetDlgItem(dlg, 9057)
            except Exception:
                continue
            if not edit or win32gui.GetClassName(edit) != "Edit":
                continue
            try:
                win32gui.SendMessage(edit, WM_SETTEXT, 0, str(ORDER_QTY))
            except Exception:
                continue
            # 校验是否真的写入（自绘/带串编辑框偶发需要重试）
            if get_edit_text(edit) != str(ORDER_QTY):
                time.sleep(0.1)
                try:
                    win32gui.SendMessage(edit, WM_SETTEXT, 0, str(ORDER_QTY))
                except Exception:
                    continue
            if get_edit_text(edit) == str(ORDER_QTY):
                print(f"[OK] 组合申报弹窗: 委托数量已设为 {ORDER_QTY}")
                return dlg
        time.sleep(0.15)
    print(f"[WARN] 等待'组合申报'弹窗超时({timeout}s)")
    return None


def handle_dialogs(main_hwnd=None):
    """连续处理点击"组合/确定"后弹出的确认/警告类弹窗。

    1) 优先匹配已知标题（组合申报/警告/确认/提示）的 #32770 弹窗，点默认按钮；
    2) 若仍有目标进程名下、带默认按钮、且非数量录入框的 #32770 弹窗（标题不确定
       的"确定弹窗"），也一并点掉，做兜底。
    最多处理 6 轮，避免无限循环。
    """
    keywords = ["组合申报", "警告", "确认", "提示"]
    target_pid = _dlg_pid(main_hwnd) if main_hwnd else None
    for _ in range(6):
        dlg = None
        for cand in find_top_windows_with(keywords):
            if win32gui.GetClassName(cand) != "#32770":
                continue
            if has_control(cand, 9057):   # 数量录入框已由主流程点过'确定'，跳过
                continue
            dlg = cand
            break
        if dlg is None and target_pid:
            # 兜底：目标进程名下任何带默认按钮的 #32770 弹窗（标题不确定）
            for cand in find_all_dialogs():
                if _dlg_pid(cand) != target_pid:
                    continue
                if has_control(cand, 9057):   # 数量录入框已由主流程单独点过确定
                    continue
                dlg = cand
                break
        if dlg is None:
            break
        try:
            win32gui.SetForegroundWindow(dlg)
        except Exception:
            pass
        time.sleep(0.1)
        title = win32gui.GetWindowText(dlg)
        click_default_button(dlg)
        print(f"[OK] '{title}'弹窗已确认")
        time.sleep(0.4)


def close_leftover_qty_dialogs():
    """启动前清理上一次运行可能残留的“组合申报”/“警告”类对话框。

    这些对话框若是模态的，会卡住目标线程的消息循环，导致后续 SendMessage 永久阻塞。
    统一点“取消”(id=2) 关闭（不提交任何委托），无取消按钮则发 WM_CLOSE。
    """
    cleaned = 0
    for dlg in find_top_windows_with(["组合申报", "警告"]):
        if win32gui.GetClassName(dlg) != "#32770":
            continue
        cancel = None
        try:
            cancel = win32gui.GetDlgItem(dlg, 2)  # 取消
        except Exception:
            cancel = None
        if cancel:
            win32gui.PostMessage(cancel, BM_CLICK, 0, 0)
        else:
            try:
                win32gui.PostMessage(dlg, win32con.WM_CLOSE, 0, 0)
            except Exception:
                pass
        print(f"[INFO] 已关闭残留弹窗 {hex(dlg)} title={win32gui.GetWindowText(dlg)!r}")
        cleaned += 1
        time.sleep(0.3)
    if cleaned:
        print(f"[INFO] 共清理 {cleaned} 个残留弹窗")


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


def check_list_has_data(main_hwnd, wait=0.4):
    list_hwnd = find_visible_child(main_hwnd, 1229)
    if not list_hwnd:
        return False
    time.sleep(wait)  # 等列表随策略切换刷新
    cnt = list_count(list_hwnd)
    if cnt >= MIN_DATA_ROWS:
        print(f"[OK] 列表有数据({cnt} 条)，执行组合申报")
        return True
    print(f"[--] 列表无足够数据({cnt} 条)，跳过")
    return False


def click_first_list_row_and_combine(main_hwnd):
    list_hwnd = find_visible_child(main_hwnd, 1229)
    if not list_hwnd:
        return False
    if list_count(list_hwnd) < MIN_DATA_ROWS:
        return False
    print(f"[OK] 列表共 {list_count(list_hwnd)} 条数据，选中第一条并提交组合")
    list_select_first(list_hwnd)
    time.sleep(0.2)

    combo_btn = find_visible_child(main_hwnd, 9058)  # "组合"按钮
    if not combo_btn:
        print("[WARN] 未找到'组合'按钮")
        return False
    # 注意：'组合'按钮会打开模态对话框（DoModal），若用 SendMessage(BM_CLICK)
    # 会一直阻塞到对话框关闭，导致后续"填数量/点确定"永远执行不到。
    # 改用 PostMessage 异步点击，点击后立即继续，由 fill_qty_in_dialog 轮询弹窗。
    win32gui.PostMessage(combo_btn, BM_CLICK, 0, 0)
    print("[OK] 已点击'组合'按钮（异步）")
    time.sleep(0.3)

    # 1) 等待数量弹窗并填写委托数量
    dlg = fill_qty_in_dialog(timeout=8)
    if dlg:
        # 2) 填写完成，点击"确定"(id=1) 提交组合申报
        if click_dialog_button(dlg, 1):
            print("[OK] 已点击'确定'提交组合申报")
        else:
            click_default_button(dlg)
            print("[OK] 已点击'组合申报'弹窗默认按钮")
        time.sleep(0.4)
        # 3) 处理提交后可能出现的二次确认/警告弹窗（标题不确定也兜底处理）
        handle_dialogs(main_hwnd)
    else:
        # 没等到数量弹窗，仍尝试兜底处理可能出现的弹窗
        handle_dialogs(main_hwnd)
    return True


def process_strategy(main_hwnd, exchange, strategy):
    print(f"\n--- 处理: {exchange} | {strategy} ---")
    if not select_strategy(main_hwnd, strategy):
        return False
    time.sleep(0.3)
    if not check_list_has_data(main_hwnd):
        return False
    if click_first_list_row_and_combine(main_hwnd):
        print(f"[OK] {exchange} | {strategy} 组合申报完成")
        return True
    print(f"[--] 组合申报执行失败")
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
        countdown(COUNTDOWN)
        hwnd = find_window(WINDOW_KEY)
        print(f"[OK] 已找到窗口，句柄 = {hex(hwnd)}")
        close_leftover_qty_dialogs()
        win = activate_window(hwnd)

        if not SKIP_NAV:
            try:
                switch_panel(win, r"\组合申报\组合申报")
            except Exception as e:
                print(f"[WARN] 切换面板失败（{e}），假定已在组合申报页面")
            time.sleep(0.5)

        # 建立下拉框 文本→索引 映射（UIA 只读一次）
        ex_combo = find_visible_child(hwnd, 9059)
        st_combo = find_visible_child(hwnd, 9040)
        if ex_combo:
            build_combo_map(ex_combo)
        if st_combo:
            build_combo_map(st_combo)
        print("[OK] 下拉框映射已建立")

        if TEST_ONE and TEST_EXCHANGE and TEST_STRATEGY:
            exchanges = [TEST_EXCHANGE]
            strategies = [TEST_STRATEGY]
        elif TEST_ONE:
            exchanges = EXCHANGES[:1]
            strategies = STRATEGIES[:1]
        else:
            exchanges = EXCHANGES
            strategies = STRATEGIES

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
                    if process_strategy(hwnd, exchange, strategy):
                        total_executed += 1
                    else:
                        total_skipped += 1
                except Exception as e:
                    print(f"[错误] 处理异常: {type(e).__name__}: {e}")
                    total_skipped += 1
                time.sleep(0.4)

        print(f"\n{'='*50}")
        print(f"=== 全部完成 ===")
        print(f"执行组合申报: {total_executed} 个")
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
