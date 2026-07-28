# -*- coding: utf-8 -*-
"""超级策略四类一键开仓的共享驱动。"""

from __future__ import annotations

import os
import re
import time

from PIL import Image
import win32api
import win32con
import win32gui
import win32process

from core.clients import get_client, get_default_client_id
from core.tactics_panel import (
    TacticsPanelError,
    capture_window_image,
    click_tactics_item,
    dpi_unaware,
    ocr_image_items,
)
from core.window import countdown, find_window
from core.workspace import WORKSPACE_SUPER, ensure_workspace


ACTION_PANEL_CONTROL_ID = 128
ACTION_PANEL_CLASS = "AfxWnd140u"
ADD_UNDERLYING_TEXT = "加入标的"
OPEN_POSITION_TEXT = "一键开仓"
LOGIN_REQUIRED_EXIT_CODE = 3
TRADING_TIME_BLOCKED_EXIT_CODE = 4
LOGIN_PATTERNS = (
    "未登录",
    "尚未登录",
    "请先登录",
    "请登录",
    "登录交易",
    "交易登录",
    "登录账号",
)
TRADING_TIME_PATTERNS = (
    "交易时间错误",
    "系统停市期间禁止委托",
    "非交易时间",
    "非交易时段",
    "不在交易时间",
)
BATCH_RESULT_PATTERNS = (
    "期权合约批量下单处理完毕",
    "批量下单处理完毕",
)
ORDER_CONFIRM_PATTERNS = (
    "您确定要下单吗",
    "确定要下单吗",
)
class SuperStrategyError(RuntimeError):
    """超级策略流程无法安全继续。"""


class LoginRequiredError(SuperStrategyError):
    """客户端需要用户先完成交易登录。"""


class TradingTimeBlockedError(SuperStrategyError):
    """委托因当前不在允许交易的时段而被客户端拒绝。"""


def _enum_descendants(hwnd: int) -> list[int]:
    found: list[int] = []
    win32gui.EnumChildWindows(int(hwnd), lambda child, _: found.append(child), None)
    return found


def _control_id(hwnd: int) -> int | None:
    try:
        return int(win32gui.GetDlgCtrlID(hwnd))
    except Exception:
        return None


def get_action_panel(main_hwnd: int) -> int:
    """定位承载“加入标的/一键开仓”的自绘策略面板。"""
    matches = []
    for hwnd in _enum_descendants(main_hwnd):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                continue
            if win32gui.GetClassName(hwnd) != ACTION_PANEL_CLASS:
                continue
            if _control_id(hwnd) != ACTION_PANEL_CONTROL_ID:
                continue
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if right - left >= 500 and bottom - top >= 150:
                matches.append(hwnd)
        except Exception:
            continue
    if len(matches) != 1:
        raise SuperStrategyError(
            f"策略操作面板定位不唯一（control_id=128, 匹配数={len(matches)}）"
        )
    return matches[0]


def _action_strip(main_hwnd: int, action_panel: int):
    image = capture_window_image(main_hwnd)
    with dpi_unaware():
        main_left, main_top, _, _ = win32gui.GetWindowRect(main_hwnd)
        left, top, right, bottom = win32gui.GetWindowRect(action_panel)
    # 按钮位于策略面板底沿，部分绘制会向下溢出数像素。
    screen_top = max(top, bottom - 32)
    x0 = left - main_left
    y0 = screen_top - main_top
    x1 = min(image.width, right - main_left)
    y1 = min(image.height, bottom - main_top + 16)
    if x1 <= x0 or y1 <= y0:
        raise SuperStrategyError("策略操作按钮截图区域无效")
    return image.crop((x0, y0, x1, y1)), (left, screen_top)


def _find_action_text(main_hwnd: int, action_panel: int, target: str) -> dict:
    # 客户端会随主窗口宽度调整按钮位置，不能使用固定面板坐标。
    # 每次只识别高度很小的操作栏，并以 OCR 返回的屏幕坐标作为真实点击点。
    image, origin = _action_strip(main_hwnd, action_panel)
    items = ocr_image_items(
        image,
        screen_origin=origin,
        min_conf=0.70,
        enlarge=3.0,
    )
    matches = [item for item in items if item["text"].strip() == target]
    if len(matches) != 1:
        # 密集操作栏偶尔会在第一次文字检测中漏掉单个按钮；同一区域换一个
        # 缩放倍率重试一次。两次都必须精确识别，绝不使用模糊文字点击。
        retry_items = ocr_image_items(
            image,
            screen_origin=origin,
            min_conf=0.70,
            enlarge=4.0,
        )
        retry_matches = [
            item for item in retry_items if item["text"].strip() == target
        ]
        if len(retry_matches) == 1:
            return retry_matches[0]
        items = retry_items
        matches = retry_matches
    if len(matches) != 1:
        seen = [item["text"] for item in items]
        raise SuperStrategyError(
            f"未唯一识别到操作按钮 {target!r}（匹配数={len(matches)}, OCR={seen}）"
        )
    return matches[0]


def _set_foreground_with_attached_input(hwnd: int) -> None:
    """把当前线程临时并入前台/目标输入队列后设置前台窗口。"""
    current_thread = int(win32api.GetCurrentThreadId())
    foreground = win32gui.GetForegroundWindow()
    thread_ids = []
    for window in (foreground, int(hwnd)):
        if not window:
            continue
        thread_id = int(win32process.GetWindowThreadProcessId(window)[0])
        if thread_id != current_thread and thread_id not in thread_ids:
            win32process.AttachThreadInput(current_thread, thread_id, True)
            thread_ids.append(thread_id)
    try:
        win32gui.BringWindowToTop(int(hwnd))
        win32gui.SetForegroundWindow(int(hwnd))
    finally:
        for thread_id in reversed(thread_ids):
            win32process.AttachThreadInput(current_thread, thread_id, False)


def _activate_main_window(main_hwnd: int) -> None:
    """恢复并置前交易客户端；真实鼠标点击前必须确认前台窗口。"""
    if win32gui.IsIconic(int(main_hwnd)):
        win32gui.ShowWindow(int(main_hwnd), win32con.SW_RESTORE)
        time.sleep(0.3)
    flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    activation_error = None
    try:
        win32gui.SetWindowPos(
            int(main_hwnd), win32con.HWND_TOPMOST, 0, 0, 0, 0, flags
        )
        try:
            win32gui.SetForegroundWindow(int(main_hwnd))
            win32gui.BringWindowToTop(int(main_hwnd))
        except Exception as exc:
            # Windows 的前台锁可能拒绝后台子进程调用 SetForegroundWindow。
            # 先记下异常，随后使用输入线程附加方式重试，不在未知 DPI 下
            # 额外点击标题栏或窗口内容。
            activation_error = exc

        time.sleep(0.15)
        foreground = win32gui.GetForegroundWindow()
        foreground_root = win32gui.GetAncestor(foreground, win32con.GA_ROOT)
        if foreground_root != int(main_hwnd):
            try:
                _set_foreground_with_attached_input(main_hwnd)
            except Exception as exc:
                raise SuperStrategyError(
                    "交易客户端未能切换到前台，已拒绝发送真实鼠标点击"
                ) from (activation_error or exc)
            time.sleep(0.15)
    finally:
        win32gui.SetWindowPos(
            int(main_hwnd), win32con.HWND_NOTOPMOST, 0, 0, 0, 0, flags
        )

    foreground = win32gui.GetForegroundWindow()
    foreground_root = win32gui.GetAncestor(foreground, win32con.GA_ROOT)
    if foreground_root != int(main_hwnd):
        raise SuperStrategyError(
            "交易客户端未能切换到前台，已拒绝发送真实鼠标点击"
        )


def _real_mouse_click(main_hwnd: int, target_hwnd: int, screen_x: float,
                      screen_y: float, *, delay: float = 0.8) -> None:
    """在 OCR 校验点执行真实鼠标点击，不再向自绘面板伪造窗口消息。"""
    _activate_main_window(main_hwnd)
    x, y = int(round(screen_x)), int(round(screen_y))
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(int(target_hwnd))
        if not (left <= x < right and top <= y < bottom):
            raise SuperStrategyError(
                f"真实点击坐标越界: point={(x, y)}, panel={(left, top, right, bottom)}"
            )
        point_hwnd = win32gui.WindowFromPoint((x, y))
        if point_hwnd != int(target_hwnd) and not win32gui.IsChild(
            int(target_hwnd), point_hwnd
        ):
            raise SuperStrategyError(
                "真实点击位置已被其他窗口遮挡，已拒绝点击"
            )
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(delay)


def click_action(main_hwnd: int, target: str, *, delay: float = 0.8) -> dict:
    """OCR 定位自绘操作按钮，并在校验后的屏幕坐标执行真实鼠标点击。"""
    panel = get_action_panel(main_hwnd)
    hit = _find_action_text(main_hwnd, panel, target)
    _real_mouse_click(
        main_hwnd,
        panel,
        hit["cx"],
        hit["cy"],
        delay=delay,
    )
    print(
        f"[OK] 已点击 {target}: 置信度={hit['score']:.3f}, "
        f"OCR={hit['ocr_elapsed']:.2f}s"
    )
    return hit


def _process_id(hwnd: int) -> int:
    return int(win32process.GetWindowThreadProcessId(int(hwnd))[1])


def _visible_process_surfaces(pid: int, main_hwnd: int) -> set[int]:
    """返回可能承载提示的可见顶层窗口和主窗口内嵌子窗口。"""
    result = set()

    def callback(hwnd, _):
        try:
            if hwnd == int(main_hwnd) or not win32gui.IsWindowVisible(hwnd):
                return
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid != pid:
                return
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if right - left >= 80 and bottom - top >= 40:
                result.add(hwnd)
        except Exception:
            return

    win32gui.EnumWindows(callback, None)
    for hwnd in _enum_descendants(main_hwnd):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                continue
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if right - left >= 80 and bottom - top >= 40:
                result.add(hwnd)
        except Exception:
            continue
    return result


def _window_text_dump(hwnd: int) -> str:
    texts = []
    try:
        title = win32gui.GetWindowText(hwnd)
        if title:
            texts.append(title)
    except Exception:
        pass
    for child in _enum_descendants(hwnd):
        try:
            text = win32gui.GetWindowText(child)
            if text:
                texts.append(text)
        except Exception:
            continue
    return "\n".join(texts)


def _is_login_prompt(text: str) -> bool:
    compact = "".join(str(text or "").split())
    return any("".join(pattern.split()) in compact for pattern in LOGIN_PATTERNS)


def _is_trading_time_prompt(text: str) -> bool:
    compact = "".join(str(text or "").split())
    return any(
        "".join(pattern.split()) in compact
        for pattern in TRADING_TIME_PATTERNS
    )


def _is_order_confirmation(text: str) -> bool:
    compact = "".join(str(text or "").split())
    return any(
        "".join(pattern.split()) in compact
        for pattern in ORDER_CONFIRM_PATTERNS
    )


def _exact_child_buttons(hwnd: int, text: str) -> list[int]:
    matches = []
    for child in _enum_descendants(hwnd):
        try:
            if not win32gui.IsWindowVisible(child):
                continue
            if win32gui.GetClassName(child) != "Button":
                continue
            if win32gui.GetWindowText(child).strip() == text:
                matches.append(child)
        except Exception:
            continue
    return matches


def _confirm_order_dialog(dialog_hwnd: int, combined_text: str) -> None:
    """仅对已核验的“一键开仓下单确认框”执行真实鼠标确定。"""
    title = win32gui.GetWindowText(int(dialog_hwnd)).strip()
    if "下单" not in title or not _is_order_confirmation(combined_text):
        raise SuperStrategyError("弹窗不是可确认的一键开仓下单确认框")

    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    cancels = _exact_child_buttons(dialog_hwnd, "取消")
    if len(confirms) != 1 or len(cancels) != 1:
        raise SuperStrategyError(
            "下单确认框按钮结构异常，已拒绝自动确认："
            f"确定={len(confirms)}, 取消={len(cancels)}"
        )

    button = confirms[0]
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(button)
    _real_mouse_click(
        dialog_hwnd,
        button,
        (left + right) / 2,
        (top + bottom) / 2,
        delay=0.3,
    )
    print("[OK] 已确认经过文字和按钮双重校验的一键开仓下单确认框")


def _parse_batch_result(text: str) -> tuple[int, int] | None:
    compact = "".join(str(text or "").split())
    if not any(
        "".join(pattern.split()) in compact
        for pattern in BATCH_RESULT_PATTERNS
    ):
        return None
    match = re.search(r"成功(\d+)笔[，,]?失败(\d+)笔", compact)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def _handle_batch_result(text: str, dialog_hwnd: int | None = None) -> dict | None:
    result = _parse_batch_result(text)
    if result is None:
        return None
    success_count, failure_count = result
    if failure_count or success_count <= 0:
        if dialog_hwnd is not None:
            _dismiss_prompt_dialog(dialog_hwnd)
        raise SuperStrategyError(
            f"批量下单未全部成功：成功{success_count}笔，失败{failure_count}笔；"
            "客户端提示框已关闭"
        )
    return {
        "open_confirmed": True,
        "success_count": success_count,
        "failure_count": failure_count,
    }


def _confirm_success_result_dialog(dialog_hwnd: int, text: str) -> dict:
    """确认成功批量结果并关闭提示框；失败结果由上层保留。"""
    result = _handle_batch_result(text)
    if result is None:
        raise SuperStrategyError("弹窗不包含可确认的批量下单成功结果")
    title = win32gui.GetWindowText(int(dialog_hwnd)).strip()
    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    if "提示" not in title or len(confirms) != 1:
        raise SuperStrategyError(
            "批量下单成功框结构异常，已拒绝自动确认："
            f"标题={title!r}, 确定={len(confirms)}"
        )

    button = confirms[0]
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(button)
    _real_mouse_click(
        dialog_hwnd,
        button,
        (left + right) / 2,
        (top + bottom) / 2,
        delay=0.3,
    )
    deadline = time.monotonic() + 1.5
    while time.monotonic() < deadline:
        if not win32gui.IsWindow(int(dialog_hwnd)) or not win32gui.IsWindowVisible(
            int(dialog_hwnd)
        ):
            print(
                "[OK] 已确认并关闭批量下单成功框："
                f"成功{result['success_count']}笔，失败0笔"
            )
            return result
        time.sleep(0.05)
    raise SuperStrategyError("已点击批量下单成功框的确定按钮，但弹窗未关闭")


def _ocr_window_text(hwnd: int) -> str:
    """仅在常规窗口文本不足时，用 OCR 读取自绘提示内容。"""
    try:
        image = capture_window_image(hwnd)
        if image.width > 1200:
            ratio = 1200 / image.width
            image = image.resize(
                (1200, max(1, int(image.height * ratio))),
                Image.Resampling.LANCZOS,
            )
        items = ocr_image_items(image, min_conf=0.45, enlarge=1.0)
        return "\n".join(item["text"] for item in items)
    except Exception as exc:
        print(f"[WARN] 提示窗口 OCR 读取失败: hwnd={hwnd}, {exc}")
        return ""


def _dismiss_prompt_dialog(dialog_hwnd: int) -> None:
    """关闭已经分类的提示框；优先真实点击唯一的“确定”按钮。"""
    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    if len(confirms) == 1:
        button = confirms[0]
        with dpi_unaware():
            left, top, right, bottom = win32gui.GetWindowRect(button)
        _real_mouse_click(
            dialog_hwnd,
            button,
            (left + right) / 2,
            (top + bottom) / 2,
            delay=0.3,
        )
    else:
        # 登录类提示在部分客户端上没有标准“确定”按钮，只允许关闭已经被
        # 明确分类的窗口；未知窗口不会进入这里。
        win32gui.PostMessage(int(dialog_hwnd), win32con.WM_CLOSE, 0, 0)

    deadline = time.monotonic() + 1.5
    while time.monotonic() < deadline:
        if not win32gui.IsWindow(int(dialog_hwnd)) or not win32gui.IsWindowVisible(
            int(dialog_hwnd)
        ):
            print("[OK] 已关闭识别完成的客户端提示框")
            return
        time.sleep(0.05)
    raise SuperStrategyError("已处理识别完成的客户端提示框，但弹窗未关闭")


def _raise_for_known_prompt(text: str, dialog_hwnd: int | None = None) -> None:
    if _is_login_prompt(text):
        if dialog_hwnd is not None:
            _dismiss_prompt_dialog(dialog_hwnd)
        raise LoginRequiredError(
            "当前交易客户端尚未登录，请完成交易登录后重新运行超级策略。"
        )
    if _is_trading_time_prompt(text):
        if dialog_hwnd is not None:
            _dismiss_prompt_dialog(dialog_hwnd)
        raise TradingTimeBlockedError(
            "当前不在允许交易的时段，客户端已拒绝委托；请在交易时段重新运行。"
        )


def wait_after_open(main_hwnd: int, before_windows: set[int],
                    timeout: float = 3.0) -> dict:
    """确认明确的下单确认框，再分类登录、非交易时段和批量结果。"""
    pid = _process_id(main_hwnd)
    deadline = time.monotonic() + timeout
    unknown: dict[int, str] = {}
    handled_confirmations: set[int] = set()
    while time.monotonic() < deadline:
        current = _visible_process_surfaces(pid, main_hwnd)
        for hwnd in current - before_windows:
            if hwnd in handled_confirmations:
                continue
            text = _window_text_dump(hwnd)
            _raise_for_known_prompt(text, hwnd)
            batch_result = _handle_batch_result(text, hwnd)
            if batch_result is not None:
                return _confirm_success_result_dialog(hwnd, text)
            ocr_text = _ocr_window_text(hwnd)
            combined = f"{text}\n{ocr_text}"
            _raise_for_known_prompt(combined, hwnd)
            batch_result = _handle_batch_result(combined, hwnd)
            if batch_result is not None:
                return _confirm_success_result_dialog(hwnd, combined)
            if ocr_text:
                text = f"{text}\n{ocr_text}".strip()
            if _is_order_confirmation(text):
                _confirm_order_dialog(hwnd, text)
                handled_confirmations.add(hwnd)
                unknown.pop(hwnd, None)
                # 确认后客户端需要时间生成下一层登录、停市或结果弹窗。
                deadline = time.monotonic() + timeout
                time.sleep(0.2)
                break
            unknown[hwnd] = text
        else:
            if unknown:
                # 给连续弹窗留出短暂稳定时间，但绝不确认未知弹窗。
                time.sleep(0.4)
                break
            time.sleep(0.1)
            continue

        # 命中并确认了明确的下单确认框，继续等待下一层弹窗。
        continue

    if unknown:
        close_errors = []
        for hwnd in list(unknown):
            try:
                _dismiss_prompt_dialog(hwnd)
            except Exception as exc:
                close_errors.append(f"hwnd={hwnd}: {exc}")
        summaries = [
            text.replace("\n", " / ") or f"hwnd={hwnd}"
            for hwnd, text in unknown.items()
        ]
        close_note = (
            "；关闭异常=" + " | ".join(close_errors)
            if close_errors else "；弹窗已关闭"
        )
        raise SuperStrategyError(
            "一键开仓后出现未识别弹窗：" + "；".join(summaries) + close_note
        )
    # 少数客户端把提示直接绘制在主窗口内部，不会新增 HWND；最后对主窗口
    # 做一次 OCR 兜底。已登录且正常执行时不会命中这些明确的错误短语。
    main_text = _ocr_window_text(main_hwnd)
    _raise_for_known_prompt(main_text)
    batch_result = _handle_batch_result(main_text)
    if batch_result is not None:
        raise SuperStrategyError(
            "已从主窗口 OCR 识别到批量下单成功，但未定位到独立结果框，"
            "无法安全点击确定"
        )
    raise SuperStrategyError(
        "真实鼠标点击后未检测到批量下单结果或错误提示，不能确认一键开仓已执行"
    )


def run_strategy(target: str, *, add_underlying: bool | None = None,
                 execute_open: bool = True) -> dict:
    """执行：自动进入超级策略 → OCR 选菜单 → 可选加入标的 → 一键开仓。"""
    if target not in {"牛市认购", "牛市认沽", "熊市认购", "熊市认沽"}:
        raise ValueError(f"不支持的超级策略: {target!r}")
    client_id = os.environ.get("GUI_CLIENT_ID") or get_default_client_id()
    client = get_client(client_id)
    if not client:
        raise SuperStrategyError(f"客户端档案不存在: {client_id!r}")
    countdown_sec = max(0, int(os.environ.get("GUI_COUNTDOWN") or 3))
    countdown(countdown_sec)
    main_hwnd = find_window(client.get("window_key") or client.get("name") or "")
    _activate_main_window(main_hwnd)
    print("[OK] 已将交易客户端切换到前台")
    ensure_workspace(main_hwnd, WORKSPACE_SUPER, client)

    print(f"[INFO] 超级策略目标: {target}")
    menu_hit = click_tactics_item(main_hwnd, target)

    if add_underlying is None:
        add_underlying = os.environ.get(
            "GUI_SUPER_ADD_UNDERLYING", "False"
        ).strip().lower() == "true"
    if add_underlying:
        click_action(main_hwnd, ADD_UNDERLYING_TEXT)
    else:
        print("[INFO] 未启用“加入标的”，跳过该步骤")

    result = {
        "target": target,
        "add_underlying": bool(add_underlying),
        "menu_hit": menu_hit,
        "open_triggered": False,
    }
    if not execute_open:
        print("[INFO] 安全验证模式：未点击一键开仓")
        return result

    pid = _process_id(main_hwnd)
    before = _visible_process_surfaces(pid, main_hwnd)
    click_action(main_hwnd, OPEN_POSITION_TEXT)
    result.update(wait_after_open(main_hwnd, before))
    result["open_triggered"] = bool(result.get("open_confirmed"))
    print(f"[OK] {target} 一键开仓动作已触发")
    return result


def main(target: str) -> None:
    try:
        run_strategy(target)
    except LoginRequiredError as exc:
        print(f"[LOGIN_REQUIRED] {exc}")
        raise SystemExit(LOGIN_REQUIRED_EXIT_CODE)
    except TradingTimeBlockedError as exc:
        print(f"[TRADING_TIME_BLOCKED] {exc}")
        raise SystemExit(TRADING_TIME_BLOCKED_EXIT_CODE)
    except (SuperStrategyError, TacticsPanelError) as exc:
        print(f"[错误] {exc}")
        raise SystemExit(1)
