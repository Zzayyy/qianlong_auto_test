# -*- coding: utf-8 -*-
"""东吴证券期权宝「协议行权参数设置」Excel 驱动器。

菜单路径: \\协议行权\\协议行权参数设置（目前仅东吴客户端有该菜单，
其他客户端由 clients.json 的 unsupported 过滤）。

面板为表单录入 + 确认提交：逐行读取 Excel，按 FIELDS 配置把值写入
表单控件并逐项校验，全部通过后才点击【确认】(auto_id=18001) 提交；
提交后按其他下单脚本同样的策略处理客户端弹出的确认/提示弹窗
（等待同进程新顶层窗口并逐个回车确认，静默后进入下一行）。

字段映射（2026-09-03 用户实机确认，控件 auto_id 以实机 dump 为准）：
  市场类别=18063(组合框)  设置方式=18059(组合框)  证券类别=18048(组合框)
  证券代码=1129(输入框)   合约类别=18050(组合框)  合约代码=18005(输入框)
  操作类别=18054(组合框)  策略类别=18047(组合框)  策略值=18049(输入框)
注意：策略类别/策略值 两个控件未出现在本次抓取的 dump 可视区域内
（可能需要滚动或依赖“设置方式”条件显示），运行时找不到控件时的行为：
该列 Excel 留空 -> 跳过；该列有值 -> 明确报错，不会盲目提交。

组合框选择策略（按优先级）：
  1. FIELDS 里配了 options -> 按“点箭头展开 + 键盘上下选择 + CB_GETCURSEL
     校验”模式按索引选择（同快速下单报价方式）；
  2. 未配 options -> 运行时用 CB_GETCOUNT/CB_GETLBTEXT 直接读取下拉项
     （这两条消息由系统跨进程封送，无需预知选项清单），按文本精确匹配
     后键盘选中；下拉项里没有该文本则报错并列出全部可选项；
  3. 读不到下拉项（count<=0 或消息失败）-> 兜底用 WM_SETTEXT 写入编辑区
     并按 WM_GETTEXTLENGTH 校验（QL 系列组合框编辑区跨进程读不到文本
     但能读到长度）。
"""

from __future__ import annotations

import ctypes
import os
import sys
import time
import unicodedata

import openpyxl
from pywinauto import Application
from pywinauto.keyboard import send_keys
import win32api
import win32con
import win32gui
import win32process

_PROJECT_ROOT = os.path.dirname(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
)  # 本脚本位于 行情交易/协议行权/ 下，向上 3 级即项目根目录
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.window import activate_window, find_window, switch_panel


WINDOW_KEY = "东吴证券期权宝"
PANEL_PATH = r"\协议行权\协议行权参数设置"
MENU_NAME = "协议行权参数设置"
COUNTDOWN = int(os.environ.get("GUI_COUNTDOWN", "3"))
INTERVAL = 1.0
SUBMIT_DIALOG_DELAY = 1.0
SUBMIT_DIALOG_TIMEOUT = 3.0
MAX_DIALOGS = 5

# 每行填写前先点【重置】清空表单，避免上一行残留值混入下一行
# （Excel 留空的列不会覆盖，若不清空就会把上一行的值提交出去）。
RESET_BEFORE_ROW = True

CONFIRM_ID = 18001
RESET_ID = 1153

# ====================== 字段配置（Excel 列名 -> 控件） ======================
# kind: "edit"  = 输入框，WM_SETTEXT 写入 + 长度校验；
#       "combo" = 组合框，OPTIONS 为空时按 WM_SETTEXT 写入编辑区并校验；
#                 OPTIONS 非空时按“下拉项索引”键盘选择并 CB_GETCURSEL 校验。
# required: 该列 Excel 是否必填（表头仍必须全部存在）。
FIELDS = [
    {"column": "市场类别", "auto_id": 18063, "kind": "combo", "options": (), "required": False},
    {"column": "设置方式", "auto_id": 18059, "kind": "combo", "options": (), "required": False},
    {"column": "证券类别", "auto_id": 18048, "kind": "combo", "options": (), "required": False},
    {"column": "证券代码", "auto_id": 1129, "kind": "edit", "options": (), "required": False},
    {"column": "合约类别", "auto_id": 18050, "kind": "combo", "options": (), "required": False},
    {"column": "合约代码", "auto_id": 18005, "kind": "edit", "options": (), "required": False},
    {"column": "操作类别", "auto_id": 18054, "kind": "combo", "options": (), "required": False},
    {"column": "策略类别", "auto_id": 18047, "kind": "combo", "options": (), "required": False},
    {"column": "策略值", "auto_id": 18049, "kind": "edit", "options": (), "required": False},
]

REQUIRED_COLUMNS = tuple(field["column"] for field in FIELDS)

CB_GETCURSEL = 0x0147
CB_GETCOUNT = 0x0146
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149


def _is_blank(value) -> bool:
    return value is None or (isinstance(value, str) and not value.strip())


def normalize_text(value, field: str) -> str:
    if isinstance(value, bool):
        raise ValueError(f"{field} 不能是布尔值")
    if isinstance(value, float):
        if not value.is_integer():
            return repr(value)
        return str(int(value))
    if isinstance(value, int):
        return str(value)
    result = str(value).strip()
    if not result:
        raise ValueError(f"{field} 不能为空")
    return result


def load_rows(path: str) -> list[dict]:
    """读取并在操作 GUI 前一次性验证所有协议行权参数行。"""
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

        rows = []
        errors = []
        for row_number, row in enumerate(
            sheet.iter_rows(min_row=2, values_only=True), start=2
        ):
            if not row or not any(not _is_blank(value) for value in row):
                continue
            item = dict(zip(headers, row))
            entry = {"excel_row": row_number, "values": {}}
            try:
                for field in FIELDS:
                    raw = item.get(field["column"])
                    if _is_blank(raw):
                        if field["required"]:
                            raise ValueError(f"{field['column']} 不能为空")
                        continue
                    entry["values"][field["column"]] = normalize_text(
                        raw, field["column"]
                    )
            except ValueError as exc:
                errors.append(f"第 {row_number} 行: {exc}")
                continue
            if not entry["values"]:
                # 所有配置列均为空（如仅有备注等无关内容）：跳过，
                # 避免对着空表单点“确认”。
                print(f"[INFO] Excel 第 {row_number} 行无有效字段，已跳过")
                continue
            rows.append(entry)

        if errors:
            raise ValueError("Excel 数据校验失败:\n" + "\n".join(errors))
        if not rows:
            raise ValueError("Excel 中没有有效数据行")
        return rows
    finally:
        workbook.close()


# ====================== Win32 基础工具（与下单脚本同源） ======================

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


def _find_control_optional(main_hwnd: int, control_id: int) -> int | None:
    """查找控件，找不到（或不可见）时返回 None，用于条件显示的字段。"""
    try:
        return _find_control(main_hwnd, control_id)
    except RuntimeError:
        return None


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
        # QL 系列自绘 Edit 不向跨进程 GetWindowText/UIA 暴露实际文本，
        # 但 WM_GETTEXTLENGTH 可正确返回字符数。
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


def _press_virtual_key(vk_code: int) -> None:
    win32api.keybd_event(vk_code, 0, 0, 0)
    win32api.keybd_event(vk_code, 0, win32con.KEYEVENTF_KEYUP, 0)


# 常见异体字归一：客户端下拉项与 Excel 录入间混用（如 按帐户/按账户），
# 两侧同规则归一后即可精确匹配。注意 maketrans 键须为单字符。
_VARIANT_CHAR_MAP = str.maketrans({
    "帐": "账",
})


def _normalize_text_key(value: str) -> str:
    """下拉项匹配用的归一化文本（全角/半角、异体字归一）。"""
    normalized = unicodedata.normalize("NFKC", str(value)).replace("\u3000", " ")
    return normalized.translate(_VARIANT_CHAR_MAP).strip()


def _read_combo_items(combo: int) -> list[str]:
    """跨进程读取组合框下拉项。

    CB_GETCOUNT/CB_GETLBTEXT/CB_GETLBTEXTLEN 由系统跨进程封送，
    无需远程内存即可读取标准/自绘组合框的项文本。
    """
    user32 = ctypes.windll.user32
    count = win32gui.SendMessage(combo, CB_GETCOUNT, 0, 0)
    if count <= 0:
        return []
    items: list[str] = []
    for index in range(count):
        length = win32gui.SendMessage(combo, CB_GETLBTEXTLEN, index, 0)
        if length <= 0:
            items.append("")
            continue
        buf = ctypes.create_unicode_buffer(length + 1)
        result = user32.SendMessageW(combo, CB_GETLBTEXT, index, buf)
        items.append(buf.value if result >= 0 else "")
    return items


def _expand_combo_and_pick(combo: int, index: int) -> None:
    """点右侧箭头展开下拉后，用键盘移动到 index 项并回车确认。"""
    # QL 系自绘组合框点击中心只聚焦编辑区，必须点右侧箭头展开。
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


def _verify_combo_cursel(combo: int, control_id: int, value: str, index: int) -> None:
    if not _wait_until(
        lambda: win32gui.SendMessage(combo, CB_GETCURSEL, 0, 0) == index,
        timeout=1.0,
    ):
        actual_index = win32gui.SendMessage(combo, CB_GETCURSEL, 0, 0)
        raise RuntimeError(
            f"auto_id={control_id} 下拉选择失败: "
            f"期望={value!r}({index}), 实际索引={actual_index}"
        )


def _select_combo_by_option(main_hwnd: int, control_id: int, options, value: str) -> None:
    """按 OPTIONS 清单的索引做键盘选择（同快速下单报价方式）。"""
    combo = _find_control(main_hwnd, control_id)
    normalized = str(value).strip()
    if normalized not in options:
        raise RuntimeError(
            f"auto_id={control_id} 下拉项清单中没有 {value!r}，"
            f"可选: {'/'.join(options)}"
        )
    index = list(options).index(normalized)
    if not win32gui.IsWindowEnabled(combo):
        raise RuntimeError(f"auto_id={control_id} 下拉框不可用")

    _expand_combo_and_pick(combo, index)
    _verify_combo_cursel(combo, control_id, value, index)


def _select_combo_from_dropdown(main_hwnd: int, control_id: int, value: str) -> bool:
    """运行时读取下拉项并按文本精确匹配选择。

    返回 True 表示已选中；返回 False 表示读不到下拉项（count<=0），
    交由调用方决定是否走“直接写入编辑区”兜底。
    下拉项里没有目标文本时直接报错并列出全部可选项。
    """
    combo = _find_control(main_hwnd, control_id)
    if not win32gui.IsWindowEnabled(combo):
        raise RuntimeError(f"auto_id={control_id} 下拉框不可用，无法选择 {value!r}")
    items = _read_combo_items(combo)
    if not items:
        return False
    target = _normalize_text_key(value)
    matches = [
        index for index, item in enumerate(items)
        if _normalize_text_key(item) == target
    ]
    if not matches:
        raise RuntimeError(
            f"auto_id={control_id} 下拉项中没有 {value!r}，"
            f"可选: {'/'.join(items)}"
        )
    index = matches[0]
    _expand_combo_and_pick(combo, index)
    _verify_combo_cursel(combo, control_id, value, index)
    print(f"[OK] 已从下拉项选中 {value!r}（共 {len(items)} 项，索引={index}）")
    return True


def _set_combo_field(main_hwnd: int, field: dict, value: str) -> None:
    """组合框统一入口：options > 运行时下拉项匹配 > 直接写入兜底。"""
    control_id = field["auto_id"]
    if field["options"]:
        _select_combo_by_option(main_hwnd, control_id, field["options"], value)
        return
    if _select_combo_from_dropdown(main_hwnd, control_id, value):
        return
    # 读不到下拉项：退回直接写入编辑区（保留旧行为）
    print(
        f"[INFO] auto_id={control_id} 读不到下拉项清单，"
        f"改用直接写入: {value!r}"
    )
    _set_combo_typed(main_hwnd, control_id, value)


def _set_combo_typed(main_hwnd: int, control_id: int, value: str) -> None:
    """OPTIONS 未配置的组合框：直接写编辑区并按长度校验。"""
    hwnd = _find_control(main_hwnd, control_id)
    if not win32gui.IsWindowEnabled(hwnd):
        raise RuntimeError(f"auto_id={control_id} 组合框不可用，无法写入 {value!r}")
    expected = str(value).strip()
    write_result = win32gui.SendMessage(hwnd, win32con.WM_SETTEXT, 0, expected)

    def matches() -> bool:
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
            f"auto_id={control_id} 组合框写入后长度校验失败: "
            f"期望长度={len(expected)}, 实际长度={actual_length}。"
            "若该组合框只允许从下拉项中选择，请在脚本 FIELDS 的 options 里"
            "补充实机下拉项顺序清单。"
        )


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


def _press_dialog_confirm(wrapper) -> None:
    """点击弹窗中的“确定/是”类按钮；找不到按钮时回退回车。

    协议行权参数提交后的确认框是 确定(Y)/取消(N) 双按钮询问框，
    回车可能落在默认的“取消”上或不起作用，必须显式点“确定”。
    """
    for pattern in (r"^确定", r"^是", r"^OK", r"^Yes"):
        try:
            button = wrapper.child_window(
                title_re=pattern, control_type="Button"
            )
            button.wait("visible ready", timeout=0.8)
            try:
                button.invoke()
            except Exception:
                button.click_input()
            return
        except Exception:
            continue
    # 没有可识别按钮：回退回车（保持旧行为）
    wrapper.set_focus()
    wrapper.type_keys("{ENTER}", with_spaces=False)


def _confirm_new_dialogs(
    main_hwnd: int,
    dialogs_before_click: set[int],
    max_dialogs: int = MAX_DIALOGS,
    first_timeout: float = 3.0,
    next_timeout: float = 1.0,
    context: str = "本次提交",
) -> int:
    """只确认本次点击后在同进程新出现的顶层窗口（与下单脚本同策略）。"""
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
                _press_dialog_confirm(wrapper)
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


def _click_button_verified(main_hwnd: int, control_id: int, name: str) -> None:
    hwnd = _find_control(main_hwnd, control_id)
    if not win32gui.IsWindowEnabled(hwnd):
        raise RuntimeError(f"按钮“{name}”(auto_id={control_id}) 当前不可用")
    _mouse_click(hwnd)
    print(f"[OK] 已点击“{name}” (auto_id={control_id})")


# ====================== 业务流程 ======================

def execute_row(main_hwnd: int, entry: dict) -> None:
    """填写一行参数；任一步失败都在点击确认按钮前抛错。"""
    values: dict = entry["values"]

    # 按面板联动规则跳过隐藏/无关字段（Excel 填了也忽略并提示）：
    #   1. 操作类别=取消自动行权 -> 面板不显示 策略类别/策略值；
    #   2. 设置方式=按账户       -> 证券类别/证券代码/合约类别/合约代码 不用填；
    #      设置方式=按合约       -> 证券类别/证券代码/合约类别 不用填；
    #      设置方式=按标的证券   -> 合约代码 不用填。
    # （设置方式匹配走异体字归一，按帐户/按账户 均可命中）
    skip_columns: set[str] = set()
    if _normalize_text_key(values.get("操作类别", "")) == "取消自动行权":
        skip_columns.update(("策略类别", "策略值"))

    mode = _normalize_text_key(values.get("设置方式", ""))
    if mode == "按账户":
        skip_columns.update(("证券类别", "证券代码", "合约类别", "合约代码"))
    elif mode == "按合约":
        skip_columns.update(("证券类别", "证券代码", "合约类别"))
    elif mode == "按标的证券":
        skip_columns.add("合约代码")

    ignored = sorted(column for column in skip_columns if column in values)
    if ignored:
        print(
            f"[INFO] Excel 第 {entry['excel_row']} 行按联动规则"
            f"（操作类别={values.get('操作类别', '')!r}, "
            f"设置方式={values.get('设置方式', '')!r}）无需填写的列已忽略: "
            f"{', '.join(ignored)}"
        )

    if RESET_BEFORE_ROW:
        reset_hwnd = _find_control_optional(main_hwnd, RESET_ID)
        if reset_hwnd is not None and win32gui.IsWindowEnabled(reset_hwnd):
            pid = _process_id(main_hwnd)
            dialogs_before = _visible_process_windows(pid, main_hwnd)
            _mouse_click(reset_hwnd)
            print("[OK] 已点击“重置”清空表单")
            time.sleep(0.3)
            _confirm_new_dialogs(
                main_hwnd,
                dialogs_before,
                first_timeout=1.5,
                next_timeout=1.0,
                context="重置表单",
            )
        else:
            print("[INFO] “重置”按钮不可用，跳过清空（将逐字段覆盖）")

    for field in FIELDS:
        column = field["column"]
        if column not in values or column in skip_columns:
            continue
        value = values[column]
        control_id = field["auto_id"]
        hwnd = _find_control_optional(main_hwnd, control_id)
        if hwnd is None:
            raise RuntimeError(
                f"Excel 第 {entry['excel_row']} 行字段“{column}”="
                f"{value!r} 无法填写：面板上找不到可见控件 "
                f"auto_id={control_id}（该控件可能需要滚动或切换"
                "“设置方式”后才显示）"
            )
        if field["kind"] == "combo":
            _set_combo_field(main_hwnd, field, value)
        else:
            _set_edit_verified(main_hwnd, control_id, value)
        print(f"[OK] {column} 已填写并校验: {value} (auto_id={control_id})")
        time.sleep(0.15)

    # 全部字段写入成功后才提交
    pid = _process_id(main_hwnd)
    dialogs_before = _visible_process_windows(pid, main_hwnd)
    _click_button_verified(main_hwnd, CONFIRM_ID, "确认")
    print(f"[INFO] 等待交易客户端响应 {SUBMIT_DIALOG_DELAY:g} 秒...")
    time.sleep(SUBMIT_DIALOG_DELAY)
    confirmed = _confirm_new_dialogs(
        main_hwnd,
        dialogs_before,
        max_dialogs=MAX_DIALOGS,
        first_timeout=SUBMIT_DIALOG_TIMEOUT,
        next_timeout=SUBMIT_DIALOG_TIMEOUT,
        context="协议行权参数提交",
    )
    print(
        f"[OK] Excel 第 {entry['excel_row']} 行协议行权参数设置完成，"
        f"确认弹窗={confirmed} 个"
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

    rows = load_rows(excel_path)
    print(f"[OK] Excel 校验通过，{MENU_NAME}共 {len(rows)} 行")
    countdown(COUNTDOWN)

    hwnd = find_window(WINDOW_KEY)
    win = activate_window(hwnd)
    switch_panel(win, PANEL_PATH)
    time.sleep(0.6)

    for index, entry in enumerate(rows, start=1):
        summary = ", ".join(
            f"{column}={value}" for column, value in entry["values"].items()
        )
        print(
            f"\n=== [{index}/{len(rows)}] Excel第{entry['excel_row']}行 {summary} ==="
        )
        try:
            execute_row(win.handle, entry)
        except Exception as exc:
            raise RuntimeError(
                f"Excel 第 {entry['excel_row']} 行执行失败，已停止后续提交: {exc}"
            ) from exc
        time.sleep(INTERVAL)

    print(f"\n=== {MENU_NAME}全部完成: {len(rows)} 行 ===")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        raise SystemExit(0)
    except Exception as error:
        print(f"\n[错误] {type(error).__name__}: {error}")
        raise SystemExit(1)
