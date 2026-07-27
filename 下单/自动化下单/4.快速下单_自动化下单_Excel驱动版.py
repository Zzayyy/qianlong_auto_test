# -*- coding: utf-8 -*-
"""快速下单 Excel 驱动器。

与三键/四键下单不同，快速下单页面通过“买卖方向 + 开平标志”
动态决定提交按钮，并且报价方式下拉框是 owner-drawn ComboLBox，
不能沿用其他下单页面的“按文字查找下拉项”逻辑。
"""

from __future__ import annotations

from dataclasses import dataclass
import math
import os
import sys
import time

import openpyxl
from pywinauto import Application
from pywinauto.keyboard import send_keys
import win32api
import win32con
import win32gui
import win32process


_PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
)
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.window import activate_window, find_window, switch_panel


PANEL_PATH = r"\快速下单"
MENU_NAME = "快速下单"
WINDOW_KEY = "钱龙模拟期权宝"
COUNTDOWN = int(os.environ.get("GUI_COUNTDOWN", "3"))
INTERVAL = 1.0
SUBMIT_DIALOG_DELAY = 1.0
SUBMIT_DIALOG_TIMEOUT = 3.0
MAX_DIALOGS = 5

REQUIRED_COLUMNS = (
    "菜单",
    "合约代码",
    "报价方式",
    "委托数量",
    "动作",
    "开平",
    "备兑",
    "FOK",
    "联动",
)

# 已在钱龙实机的 owner-drawn ComboLBox 中按顺序核对。
QUOTE_TYPES = (
    "对手价",
    "挂盘价",
    "涨停价",
    "跌停价",
    "限价",
    "超价",
    "市价转限",
    "市价FAK",
    "市价FOK",
)
QUOTE_INDEX = {name: index for index, name in enumerate(QUOTE_TYPES)}

CONTROL_IDS = {
    "合约代码": 18005,
    "买入": 6006,
    "卖出": 6007,
    "开仓": 6009,
    "平仓": 6010,
    "备兑": 5025,
    "FOK": 5027,
    "报价方式": 18059,
    "联动": 5024,
    "委托数量": 306,
    "全": 2384,
    "提交": 18001,
}

ACTION_TITLES = {
    ("买入", "开仓"): "买入开仓",
    ("卖出", "开仓"): "卖出开仓",
    ("买入", "平仓"): "买入平仓",
    ("卖出", "平仓"): "卖出平仓",
}


@dataclass(frozen=True)
class QuickOrder:
    excel_row: int
    contract_code: str
    direction: str
    offset: str
    quote_type: str
    quantity: str
    covered: bool
    fok: bool
    linked: bool

    @property
    def action_title(self) -> str:
        return ACTION_TITLES[(self.direction, self.offset)]


def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def normalize_bool(value, field: str, default: bool | None = None) -> bool:
    if _is_blank(value) and default is not None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if value in (0, 1):
            return bool(value)
    if isinstance(value, str):
        normalized = value.strip().lower()
        if normalized in {"true", "是", "1", "yes", "y"}:
            return True
        if normalized in {"false", "否", "0", "no", "n"}:
            return False
    raise ValueError(f"{field}必须是 True/False（也支持是/否、1/0）")


def normalize_contract_code(value) -> str:
    if isinstance(value, bool) or _is_blank(value):
        raise ValueError("合约代码不能为空")
    if isinstance(value, float):
        if not math.isfinite(value) or not value.is_integer():
            raise ValueError("合约代码必须是整数或文本")
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    result = str(value).strip()
    if not result:
        raise ValueError("合约代码不能为空")
    return result


def normalize_choice(value, field: str, choices: tuple[str, ...]) -> str:
    normalized = "" if value is None else str(value).strip()
    if normalized not in choices:
        raise ValueError(f"{field}必须是 {'/'.join(choices)}")
    return normalized


def normalize_alias_choice(value, field: str, aliases: dict[str, str]) -> str:
    normalized = "" if value is None else str(value).strip()
    if normalized not in aliases:
        raise ValueError(f"{field}必须是 {'/'.join(aliases)}")
    return aliases[normalized]


def normalize_quantity(value) -> str:
    if isinstance(value, bool) or _is_blank(value):
        raise ValueError("委托数量必须是正整数或‘全’")
    if isinstance(value, str):
        normalized = value.strip()
        if normalized == "全":
            return normalized
        if normalized.isdigit() and int(normalized) > 0:
            return str(int(normalized))
        raise ValueError("委托数量必须是正整数或‘全’")
    if isinstance(value, int) and value > 0:
        return str(value)
    if isinstance(value, float) and math.isfinite(value) and value.is_integer() and value > 0:
        return str(int(value))
    raise ValueError("委托数量必须是正整数或‘全’")


def load_orders(path: str) -> list[QuickOrder]:
    """读取并在操作 GUI 前一次性验证所有快速下单行。"""
    workbook = openpyxl.load_workbook(path, data_only=True)
    try:
        sheet = workbook.active
        raw_headers = [cell.value for cell in sheet[1]]
        headers = [str(value).strip() if value is not None else "" for value in raw_headers]

        nonempty_headers = [header for header in headers if header]
        duplicates = sorted({h for h in nonempty_headers if nonempty_headers.count(h) > 1})
        if duplicates:
            raise ValueError(f"Excel 表头重复: {', '.join(duplicates)}")

        missing = [column for column in REQUIRED_COLUMNS if column not in headers]
        if missing:
            raise ValueError(f"Excel 缺少必填列: {', '.join(missing)}")

        orders = []
        errors = []
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            if not row or not any(not _is_blank(value) for value in row):
                continue
            item = dict(zip(headers, row))
            if str(item.get("菜单", "")).strip() != MENU_NAME:
                continue
            try:
                linked = normalize_bool(item.get("联动"), "联动", default=False)
                orders.append(
                    QuickOrder(
                        excel_row=row_number,
                        contract_code=normalize_contract_code(item.get("合约代码")),
                        direction=normalize_alias_choice(
                            item.get("动作"), "动作", {"买入": "买入", "卖出": "卖出"}
                        ),
                        offset=normalize_alias_choice(
                            item.get("开平"),
                            "开平",
                            {"开": "开仓", "平": "平仓", "开仓": "开仓", "平仓": "平仓"},
                        ),
                        quote_type=normalize_choice(
                            item.get("报价方式"), "报价方式", QUOTE_TYPES
                        ),
                        quantity=normalize_quantity(item.get("委托数量")),
                        covered=normalize_bool(item.get("备兑"), "备兑", default=False),
                        fok=normalize_bool(item.get("FOK"), "FOK", default=False),
                        linked=linked,
                    )
                )
            except ValueError as exc:
                errors.append(f"第 {row_number} 行: {exc}")

        if errors:
            raise ValueError("Excel 数据校验失败:\n" + "\n".join(errors))
        if not orders:
            raise ValueError("没有找到‘菜单=快速下单’的有效数据行")
        return orders
    finally:
        workbook.close()


def _find_control(main_hwnd: int, control_id: int) -> int:
    matches = []

    def callback(hwnd, _):
        try:
            if (
                win32gui.GetDlgCtrlID(hwnd) == control_id
                and win32gui.IsWindowVisible(hwnd)
            ):
                matches.append(hwnd)
        except Exception:
            pass

    win32gui.EnumChildWindows(main_hwnd, callback, None)
    if len(matches) != 1:
        raise RuntimeError(
            f"控件 auto_id={control_id} 可见候选数应为 1，实际为 {len(matches)}"
        )
    return matches[0]


def _mouse_click(hwnd: int) -> None:
    root = win32gui.GetAncestor(hwnd, 2) or win32gui.GetParent(hwnd)
    if root:
        try:
            win32gui.SetForegroundWindow(root)
        except Exception:
            pass
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    if right <= left or bottom <= top:
        raise RuntimeError(f"控件无可点击区域: hwnd={hwnd}")
    win32api.SetCursorPos(((left + right) // 2, (top + bottom) // 2))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)


def _wait_until(predicate, timeout: float = 1.5, interval: float = 0.1) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        try:
            if predicate():
                return True
        except Exception:
            pass
        time.sleep(interval)
    return False


def _set_edit_verified(main_hwnd: int, control_id: int, value: str) -> None:
    hwnd = _find_control(main_hwnd, control_id)
    expected = str(value).strip()
    write_result = win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, expected)

    def matches() -> bool:
        # 钱龙/国泰的该类自绘 Edit 不向跨进程 GetWindowText/UIA
        # 暴露实际文本，但 WM_GETTEXTLENGTH 可正确返回字符数。
        return win32gui.SendMessage(
            hwnd, win32con.WM_GETTEXTLENGTH, 0, 0
        ) == len(expected)

    if not write_result or not _wait_until(matches, timeout=0.8):
        _mouse_click(hwnd)
        send_keys("^a", pause=0.03)
        send_keys(expected, with_spaces=True, pause=0.03)
    if not _wait_until(matches, timeout=1.2):
        actual_length = win32gui.SendMessage(
            hwnd, win32con.WM_GETTEXTLENGTH, 0, 0
        )
        raise RuntimeError(
            f"auto_id={control_id} 写入后长度校验失败: "
            f"期望长度={len(expected)}, 实际长度={actual_length}"
        )


def _set_toggle(main_hwnd: int, control_id: int, desired: bool, name: str) -> None:
    hwnd = _find_control(main_hwnd, control_id)

    def checked() -> bool:
        return win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != 0

    if checked() == desired:
        return
    if not win32gui.IsWindowEnabled(hwnd):
        raise RuntimeError(f"控件‘{name}’已置灰，无法设为 {desired}")
    _mouse_click(hwnd)
    if not _wait_until(lambda: checked() == desired):
        raise RuntimeError(f"控件‘{name}’状态设置失败")


def _set_confirmed_toggle(
    main_hwnd: int, control_id: int, desired: bool, name: str
) -> None:
    """设置会弹确认框的复选项，并在确认后校验最终状态。"""
    hwnd = _find_control(main_hwnd, control_id)

    def checked() -> bool:
        current = _find_control(main_hwnd, control_id)
        return win32gui.SendMessage(current, win32con.BM_GETCHECK, 0, 0) != 0

    if checked() == desired:
        return
    if not win32gui.IsWindowEnabled(hwnd):
        raise RuntimeError(f"控件‘{name}’已置灰，无法设为 {desired}")

    pid = _process_id(main_hwnd)
    dialogs_before = _visible_process_windows(pid, main_hwnd)
    _mouse_click(hwnd)
    confirmed = _confirm_new_dialogs(
        main_hwnd,
        dialogs_before,
        first_timeout=2.0,
        context=f"切换‘{name}’",
    )
    if confirmed == 0:
        raise RuntimeError(f"切换‘{name}’后未出现确认提示框")
    if not _wait_until(lambda: checked() == desired, timeout=1.5):
        raise RuntimeError(f"控件‘{name}’确认后状态设置失败")


def _set_radio(main_hwnd: int, selected_name: str) -> None:
    control_id = CONTROL_IDS[selected_name]
    hwnd = _find_control(main_hwnd, control_id)
    if win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) == 0:
        if not win32gui.IsWindowEnabled(hwnd):
            raise RuntimeError(f"单选项‘{selected_name}’不可用")
        _mouse_click(hwnd)
    if not _wait_until(
        lambda: win32gui.SendMessage(hwnd, win32con.BM_GETCHECK, 0, 0) != 0
    ):
        raise RuntimeError(f"单选项‘{selected_name}’未选中")


def _press_virtual_key(vk_code: int) -> None:
    win32api.keybd_event(vk_code, 0, 0, 0)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)


def _select_quote_type(main_hwnd: int, quote_type: str) -> None:
    combo = _find_control(main_hwnd, CONTROL_IDS["报价方式"])
    index = QUOTE_INDEX[quote_type]
    if not win32gui.IsWindowEnabled(combo):
        raise RuntimeError("报价方式下拉框不可用")

    # 该页面是 owner-drawn ComboBox，点击中心只会聚焦编辑区，
    # 必须点击右侧箭头才会真正展开 ComboLBox。
    root = win32gui.GetAncestor(combo, 2) or win32gui.GetParent(combo)
    if root:
        try:
            win32gui.SetForegroundWindow(root)
        except Exception:
            pass
    left, top, right, bottom = win32gui.GetWindowRect(combo)
    win32api.SetCursorPos((right - 12, (top + bottom) // 2))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(0.15)
    _press_virtual_key(win32con.VK_HOME)
    for _ in range(index):
        _press_virtual_key(win32con.VK_DOWN)
    _press_virtual_key(win32con.VK_RETURN)

    cb_getcursel = 0x0147
    if not _wait_until(
        lambda: win32gui.SendMessage(combo, cb_getcursel, 0, 0) == index,
        timeout=1.0,
    ):
        actual_index = win32gui.SendMessage(combo, cb_getcursel, 0, 0)
        raise RuntimeError(
            f"报价方式选择失败: 期望={quote_type}({index}), 实际索引={actual_index}"
        )


def _set_quantity(main_hwnd: int, quantity: str) -> None:
    if quantity != "全":
        _set_edit_verified(main_hwnd, CONTROL_IDS["委托数量"], quantity)
        return

    button = _find_control(main_hwnd, CONTROL_IDS["全"])
    if not win32gui.IsWindowEnabled(button):
        raise RuntimeError("‘全’按钮不可用")
    _mouse_click(button)
    quantity_edit = _find_control(main_hwnd, CONTROL_IDS["委托数量"])

    def has_positive_quantity() -> bool:
        value = (win32gui.GetWindowText(quantity_edit) or "").strip()
        return value.isdigit() and int(value) > 0

    if not _wait_until(has_positive_quantity, timeout=1.5):
        actual = (win32gui.GetWindowText(quantity_edit) or "").strip()
        raise RuntimeError(f"点击‘全’后未得到正整数委托数量: {actual!r}")


def _process_id(hwnd: int) -> int:
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return pid


def _visible_process_windows(pid: int, exclude_hwnd: int) -> set[int]:
    result = set()

    def callback(hwnd, _):
        try:
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if (
                window_pid == pid
                and hwnd != exclude_hwnd
                and win32gui.IsWindowVisible(hwnd)
            ):
                result.add(hwnd)
        except Exception:
            pass

    win32gui.EnumWindows(callback, None)
    return result


def _confirm_new_dialogs(
    main_hwnd: int,
    dialogs_before_click: set[int],
    max_dialogs: int = MAX_DIALOGS,
    first_timeout: float = 3.0,
    next_timeout: float = 1.0,
    context: str = "本次下单",
) -> int:
    """只确认本次点击后在同进程新出现的顶层窗口。"""
    pid = _process_id(main_hwnd)
    handled = set(dialogs_before_click)
    count = 0

    for dialog_number in range(1, max_dialogs + 1):
        timeout = first_timeout if dialog_number == 1 else next_timeout
        print(
            f"[WARN] {context}：等待第 {dialog_number} 个弹窗 "
            f"({timeout:g}s 无新弹窗则结束)..."
        )
        deadline = time.monotonic() + timeout
        dialog = None
        while time.monotonic() < deadline:
            current = _visible_process_windows(pid, main_hwnd)
            candidates = [hwnd for hwnd in current if hwnd not in handled]
            if candidates:
                dialog = candidates[0]
                break
            time.sleep(0.1)

        if dialog is None:
            print(f"[OK] {context}：无更多弹窗，共确认 {count} 个")
            return count

        handled.add(dialog)
        title = win32gui.GetWindowText(dialog) or ""
        closed = False
        last_error = None
        for attempt in range(1, 4):
            try:
                app = Application(backend="uia").connect(handle=dialog, timeout=0.8)
                wrapper = app.window(handle=dialog)
                wrapper.set_focus()
                wrapper.type_keys("{ENTER}", with_spaces=False)
                if _wait_until(
                    lambda: not win32gui.IsWindow(dialog)
                    or not win32gui.IsWindowVisible(dialog),
                    timeout=0.8,
                ):
                    closed = True
                    print(
                        f"[OK] 回车确认 (hwnd={dialog}, title={title!r}, "
                        f"context={context!r})"
                    )
                    break
            except Exception as exc:
                last_error = exc
            print(
                f"[--] {context}弹窗确认后未关闭，准备重试: "
                f"title={title!r} ({attempt}/3)"
            )
            time.sleep(0.15)
        if not closed:
            detail = f": {last_error}" if last_error else ""
            raise RuntimeError(
                f"{context}的新弹窗确认后仍未关闭 "
                f"hwnd={dialog}, title={title!r}{detail}"
            )
        count += 1
        print(f"[OK] {context}：已确认第 {count} 个弹窗")
        time.sleep(0.4)

    print(f"[WARN] {context}：达到最大弹窗数量上限 ({max_dialogs})")
    return count


def execute_order(main_hwnd: int, order: QuickOrder) -> None:
    """填写一笔数据；任一步失败都在点击提交按钮前抛错。"""
    _set_edit_verified(main_hwnd, CONTROL_IDS["合约代码"], order.contract_code)
    print(
        f"[OK] 合约代码已填写并校验: {order.contract_code} "
        f"(auto_id={CONTROL_IDS['合约代码']})"
    )
    time.sleep(0.25)

    _set_radio(main_hwnd, order.direction)
    print(
        f"[OK] 买卖方向已选择: {order.direction} "
        f"(auto_id={CONTROL_IDS[order.direction]})"
    )
    _set_radio(main_hwnd, order.offset)
    print(
        f"[OK] 开平标志已选择: {order.offset} "
        f"(auto_id={CONTROL_IDS[order.offset]})"
    )
    _select_quote_type(main_hwnd, order.quote_type)
    print(
        f"[OK] 报价方式已选择: {order.quote_type} "
        f"(索引={QUOTE_INDEX[order.quote_type]}, auto_id={CONTROL_IDS['报价方式']})"
    )
    # 报价切换会异步刷新自动委托价和“联动”状态。
    time.sleep(0.4)

    _set_toggle(main_hwnd, CONTROL_IDS["备兑"], order.covered, "备兑")
    print(f"[OK] 备兑状态已校验: {'是' if order.covered else '否'}")
    _set_toggle(main_hwnd, CONTROL_IDS["FOK"], order.fok, "FOK")
    print(f"[OK] FOK状态已校验: {'是' if order.fok else '否'}")
    _set_confirmed_toggle(main_hwnd, CONTROL_IDS["联动"], order.linked, "联动")
    print(f"[OK] 联动状态已校验: {'是' if order.linked else '否'}")

    _set_quantity(main_hwnd, order.quantity)
    print(
        f"[OK] 委托数量已填写并校验: {order.quantity} "
        f"(auto_id={CONTROL_IDS['委托数量']})"
    )

    submit = _find_control(main_hwnd, CONTROL_IDS["提交"])
    if not win32gui.IsWindowEnabled(submit):
        raise RuntimeError("快速下单提交按钮不可用")
    actual_title = (win32gui.GetWindowText(submit) or "").strip()
    if actual_title != order.action_title:
        raise RuntimeError(
            f"提交按钮动作校验失败: "
            f"期望={order.action_title!r}, 实际={actual_title!r}"
        )

    pid = _process_id(main_hwnd)
    dialogs_before = _visible_process_windows(pid, main_hwnd)
    _mouse_click(submit)
    print(
        f"[OK] 下单动作已触发: {order.action_title} "
        f"(auto_id={CONTROL_IDS['提交']})"
    )
    print(f"[INFO] 等待交易客户端响应 {SUBMIT_DIALOG_DELAY:g} 秒...")
    time.sleep(SUBMIT_DIALOG_DELAY)
    confirmed = _confirm_new_dialogs(
        main_hwnd,
        dialogs_before,
        max_dialogs=MAX_DIALOGS,
        first_timeout=SUBMIT_DIALOG_TIMEOUT,
        next_timeout=SUBMIT_DIALOG_TIMEOUT,
        context="本次下单",
    )
    print(
        f"[OK] Excel 第 {order.excel_row} 行下单流程完成: "
        f"动作={order.action_title}，确认弹窗={confirmed} 个"
    )


def countdown(seconds: int) -> None:
    print(f"将在 {seconds} 秒后开始，请确保交易客户端处于可操作状态...")
    for remaining in range(seconds, 0, -1):
        print(f"  {remaining}...", end="\r")
        time.sleep(1)
    print(" " * 30, end="\r")


def main() -> None:
    excel_path = os.environ.get("GUI_XLSX_FILE", "").strip()
    if not excel_path or not os.path.isfile(excel_path):
        raise FileNotFoundError("请从 GUI 选择存在的 Excel 配置文件")

    orders = load_orders(excel_path)
    print(f"[OK] Excel 校验通过，快速下单共 {len(orders)} 笔")
    countdown(COUNTDOWN)

    hwnd = find_window(WINDOW_KEY)
    win = activate_window(hwnd)
    switch_panel(win, PANEL_PATH)
    time.sleep(0.6)

    for index, order in enumerate(orders, start=1):
        print(
            f"\n=== [{index}/{len(orders)}] Excel第{order.excel_row}行 "
            f"合约={order.contract_code} 动作={order.action_title} "
            f"报价={order.quote_type} 数量={order.quantity} ==="
        )
        try:
            execute_order(win.handle, order)
        except Exception as exc:
            raise RuntimeError(
                f"Excel 第 {order.excel_row} 行执行失败，已停止后续下单: {exc}"
            ) from exc
        time.sleep(INTERVAL)

    print(f"\n=== 快速下单全部完成: {len(orders)} 笔 ===")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        raise SystemExit(0)
    except Exception as error:
        print(f"\n[错误] {type(error).__name__}: {error}")
        raise SystemExit(1)
