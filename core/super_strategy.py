# -*- coding: utf-8 -*-
"""超级策略六类一键开仓的共享驱动。"""

from __future__ import annotations

import ctypes
import json
import os
import re
import tempfile
import time

import win32api
import win32con
import win32gui
import win32process

from core.clients import (
    DEFAULT_SUPER_STRATEGY_UNDERLYING,
    SUPER_STRATEGY_UNDERLYINGS,
    get_client,
    get_default_client_id,
)
from core.tactics_panel import (
    TacticsPanelError,
    capture_window_image,
    click_tactics_item,
    dpi_unaware,
    ocr_image_items,
    ocr_single_line,
    select_super_underlying,
)
from core.window import countdown, find_window
from core.workspace import WORKSPACE_SUPER, ensure_workspace
# 复用组合/拆分申报的 Win32 控件原语（下拉映射、控件查找、弹窗处理等）
from core import combination_order as _combo_mod


ACTION_PANEL_CONTROL_ID = 128
ACTION_PANEL_CLASS = "AfxWnd140u"
SUPER_STRATEGY_TARGETS = frozenset({
    "牛市认购",
    "牛市认沽",
    "熊市认购",
    "熊市认沽",
    "卖出跨式",
    "卖宽跨式",
})
ADD_UNDERLYING_TEXT = "加入标的"
OPEN_POSITION_TEXT = "一键开仓"
COMBINATION_DECLARE_TEXT = "组合申报"
ACTION_CACHE_VERSION = 1
ACTION_CACHE_TARGETS = (
    ADD_UNDERLYING_TEXT,
    OPEN_POSITION_TEXT,
    COMBINATION_DECLARE_TEXT,
)
LOGIN_REQUIRED_EXIT_CODE = 3
TRADING_TIME_BLOCKED_EXIT_CODE = 4
RESULT_QUIET_PERIOD = 1.2
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
DIALOG_CLASS = "#32770"
AFFIRMATIVE_BUTTON_IDS = (win32con.IDOK, win32con.IDYES, 5051)


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


def _action_cache_path() -> str:
    override = os.environ.get("GUI_SUPER_ACTION_CACHE")
    if override:
        return os.path.abspath(override)
    root = os.environ.get("LOCALAPPDATA") or os.environ.get("TEMP") or os.getcwd()
    return os.path.join(root, "qianlong_auto", "super_strategy_action_cache.json")


def _load_action_cache() -> dict:
    try:
        with open(_action_cache_path(), "r", encoding="utf-8") as stream:
            payload = json.load(stream)
        if payload.get("version") != ACTION_CACHE_VERSION:
            return {"version": ACTION_CACHE_VERSION, "entries": {}}
        if not isinstance(payload.get("entries"), dict):
            return {"version": ACTION_CACHE_VERSION, "entries": {}}
        return payload
    except (OSError, ValueError, TypeError, AttributeError):
        return {"version": ACTION_CACHE_VERSION, "entries": {}}


def _save_action_cache(payload: dict) -> None:
    path = _action_cache_path()
    directory = os.path.dirname(path) or os.getcwd()
    os.makedirs(directory, exist_ok=True)
    fd, temporary = tempfile.mkstemp(
        prefix="super_strategy_action_", suffix=".json.tmp", dir=directory
    )
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as stream:
            json.dump(payload, stream, ensure_ascii=False, indent=2)
        os.replace(temporary, path)
    except Exception:
        try:
            os.unlink(temporary)
        except OSError:
            pass
        raise


def _action_cache_key(action_panel: int, image) -> str:
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(int(action_panel))
    client_id = os.environ.get("GUI_CLIENT_ID") or get_default_client_id()
    return (
        f"{client_id}|panel={right-left}x{bottom-top}|"
        f"strip={image.width}x{image.height}"
    )


def _unique_action_hits(items: list[dict]) -> dict[str, dict]:
    result = {}
    for target in ACTION_CACHE_TARGETS:
        matches = [item for item in items if item["text"].strip() == target]
        if len(matches) == 1:
            result[target] = matches[0]
    return result


def _cache_detected_action_hits(cache_key: str, origin: tuple[int, int],
                                image, hits: dict[str, dict]) -> None:
    if not hits:
        return
    payload = _load_action_cache()
    entries = payload.setdefault("entries", {})
    entry = entries.setdefault(cache_key, {})
    for target, hit in hits.items():
        entry[target] = {
            "left": float(hit["left"]) - origin[0],
            "top": float(hit["top"]) - origin[1],
            "right": float(hit["right"]) - origin[0],
            "bottom": float(hit["bottom"]) - origin[1],
        }
    entry["updated_at"] = time.time()
    entry["strip_size"] = [image.width, image.height]
    try:
        _save_action_cache(payload)
    except OSError as exc:
        print(f"[WARN] 超级策略按钮位置缓存写入失败，继续使用本次 OCR 结果: {exc}")


def _cached_action_hits(image, origin: tuple[int, int], cache_key: str,
                        targets: tuple[str, ...]) -> dict[str, dict] | None:
    entry = _load_action_cache().get("entries", {}).get(cache_key)
    if not isinstance(entry, dict):
        return None
    verified = {}
    for target in targets:
        box = entry.get(target)
        if not isinstance(box, dict):
            return None
        try:
            left = float(box["left"])
            top = float(box["top"])
            right = float(box["right"])
            bottom = float(box["bottom"])
        except (KeyError, TypeError, ValueError):
            return None
        padding_x = max(4.0, (right - left) * 0.10)
        padding_y = max(3.0, (bottom - top) * 0.25)
        crop_box = (
            max(0, int(left - padding_x)),
            max(0, int(top - padding_y)),
            min(image.width, int(right + padding_x + 1)),
            min(image.height, int(bottom + padding_y + 1)),
        )
        if crop_box[2] <= crop_box[0] or crop_box[3] <= crop_box[1]:
            return None
        recognized = ocr_single_line(
            image.crop(crop_box), min_conf=0.70, enlarge=3.0
        )
        if recognized is None or recognized["text"].strip() != target:
            return None
        verified[target] = {
            "text": target,
            "score": recognized["score"],
            "left": origin[0] + left,
            "top": origin[1] + top,
            "right": origin[0] + right,
            "bottom": origin[1] + bottom,
            "cx": origin[0] + (left + right) / 2,
            "cy": origin[1] + (top + bottom) / 2,
            "ocr_elapsed": recognized["ocr_elapsed"],
            "mode": "cached_recognition_only",
        }
    return verified


def _find_action_texts(main_hwnd: int, action_panel: int,
                       targets) -> dict[str, dict]:
    """一次截图定位多个操作按钮；缓存只在逐按钮 OCR 复核后使用。"""
    targets = tuple(dict.fromkeys(str(target) for target in targets))
    if not targets:
        return {}
    unsupported = [target for target in targets if target not in ACTION_CACHE_TARGETS]
    if unsupported:
        raise SuperStrategyError(f"不支持的策略操作按钮: {unsupported}")

    # 客户端会随主窗口宽度调整按钮位置，不能直接使用固定屏幕坐标。
    # 首次完整检测后缓存相对位置；以后仅裁出单个文字块做识别复核。
    image, origin = _action_strip(main_hwnd, action_panel)
    cache_key = _action_cache_key(action_panel, image)
    cached = _cached_action_hits(image, origin, cache_key, targets)
    if cached is not None:
        total = sum(hit["ocr_elapsed"] for hit in cached.values())
        print(
            f"[OK] 已用缓存位置快速复核操作按钮: {', '.join(targets)}，"
            f"OCR={total:.2f}s"
        )
        return cached

    detected = {}
    last_items = []
    for enlarge in (3.0, 4.0):
        items = ocr_image_items(
            image,
            screen_origin=origin,
            min_conf=0.70,
            enlarge=enlarge,
            use_cls=False,
        )
        last_items = items
        detected.update(_unique_action_hits(items))
        if all(target in detected for target in targets):
            break
    if not all(target in detected for target in targets):
        seen = [item["text"] for item in last_items]
        missing = [target for target in targets if target not in detected]
        raise SuperStrategyError(
            f"未唯一识别到操作按钮 {missing!r}（OCR={seen}）"
        )
    _cache_detected_action_hits(cache_key, origin, image, detected)
    for hit in detected.values():
        hit["mode"] = "full_detection"
    return {target: detected[target] for target in targets}


def _find_action_text(main_hwnd: int, action_panel: int, target: str) -> dict:
    return _find_action_texts(main_hwnd, action_panel, (target,))[target]


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


def click_action(main_hwnd: int, target: str, *, delay: float = 0.8,
                 panel: int | None = None, hit: dict | None = None) -> dict:
    """OCR 定位自绘操作按钮，并在校验后的屏幕坐标执行真实鼠标点击。"""
    panel = get_action_panel(main_hwnd) if panel is None else int(panel)
    hit = _find_action_text(main_hwnd, panel, target) if hit is None else hit
    if hit.get("text", "").strip() != target:
        raise SuperStrategyError(
            f"预识别按钮与点击目标不一致: 目标={target!r}, 识别={hit.get('text')!r}"
        )
    _real_mouse_click(
        main_hwnd,
        panel,
        hit["cx"],
        hit["cy"],
        delay=delay,
    )
    print(
        f"[OK] 已点击 {target}: 置信度={hit['score']:.3f}, "
        f"OCR={hit['ocr_elapsed']:.2f}s, 模式={hit.get('mode', 'full_detection')}"
    )
    return hit


def _process_id(hwnd: int) -> int:
    return int(win32process.GetWindowThreadProcessId(int(hwnd))[1])


def _visible_process_surfaces(pid: int, main_hwnd: int) -> set[int]:
    """返回目标进程内可见的原生对话框，不扫描普通业务子面板。"""
    result = set()

    def callback(hwnd, _):
        try:
            if hwnd == int(main_hwnd) or not win32gui.IsWindowVisible(hwnd):
                return True
            _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
            if window_pid != pid:
                return True
            if win32gui.GetClassName(hwnd) == DIALOG_CLASS:
                result.add(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(callback, None)
    for hwnd in _enum_descendants(main_hwnd):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                continue
            if win32gui.GetClassName(hwnd) == DIALOG_CLASS:
                result.add(hwnd)
        except Exception:
            continue
    return result


def _native_window_text(hwnd: int, max_chars: int = 32768) -> str:
    """通过原生 WM_GETTEXT 读取跨进程窗口/控件文字。"""
    try:
        length = int(
            win32gui.SendMessage(
                int(hwnd), win32con.WM_GETTEXTLENGTH, 0, 0
            )
        )
        size = max(1, min(length + 1, int(max_chars)))
        buffer = ctypes.create_unicode_buffer(size)
        win32gui.SendMessage(
            int(hwnd), win32con.WM_GETTEXT, size, buffer
        )
        text = buffer.value.strip()
        if text:
            return text
    except Exception:
        pass
    try:
        return (win32gui.GetWindowText(int(hwnd)) or "").strip()
    except Exception:
        return ""


def _window_text_dump(hwnd: int) -> str:
    """只用 Win32 读取弹窗标题及所有原生子控件文字。"""
    texts: list[str] = []
    title = _native_window_text(hwnd)
    if title:
        texts.append(title)
    for child in _enum_descendants(hwnd):
        text = _native_window_text(child)
        if text and text not in texts:
            texts.append(text)
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
            if _native_window_text(child) == text:
                matches.append(child)
        except Exception:
            continue
    return matches


def _button_by_id(dialog_hwnd: int, control_ids: tuple[int, ...]) -> int | None:
    for control_id in control_ids:
        try:
            button = win32gui.GetDlgItem(int(dialog_hwnd), int(control_id))
        except Exception:
            button = 0
        if not button:
            continue
        try:
            if (
                win32gui.IsWindowVisible(button)
                and win32gui.IsWindowEnabled(button)
                and win32gui.GetClassName(button) == "Button"
            ):
                return int(button)
        except Exception:
            continue
    return None


def _dialog_signature(dialog_hwnd: int) -> tuple:
    """生成轻量原生状态签名，用于验证弹窗关闭或内容切换。"""
    controls = []
    for child in _enum_descendants(dialog_hwnd):
        try:
            if not win32gui.IsWindowVisible(child):
                continue
            controls.append((
                _control_id(child),
                win32gui.GetClassName(child),
                _native_window_text(child),
            ))
        except Exception:
            continue
    return (_native_window_text(dialog_hwnd), tuple(controls))


def _dialog_is_visible(dialog_hwnd: int) -> bool:
    try:
        return bool(
            win32gui.IsWindow(int(dialog_hwnd))
            and win32gui.IsWindowVisible(int(dialog_hwnd))
        )
    except Exception:
        return False


def _wait_dialog_transition(dialog_hwnd: int, previous_signature: tuple,
                            timeout: float = 1.5, *,
                            require_closed: bool = False) -> bool:
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not _dialog_is_visible(dialog_hwnd):
            return True
        if (
            not require_closed
            and _dialog_signature(dialog_hwnd) != previous_signature
        ):
            return True
        time.sleep(0.05)
    return False


def _click_dialog_button(dialog_hwnd: int, button_hwnd: int, *,
                         attempts: int = 3, timeout: float = 1.5,
                         require_closed: bool = False) -> None:
    """异步发送 BM_CLICK，并确认弹窗关闭或切换到下一状态。"""
    for attempt in range(1, attempts + 1):
        if not _dialog_is_visible(dialog_hwnd):
            return
        signature = _dialog_signature(dialog_hwnd)
        win32gui.PostMessage(int(button_hwnd), win32con.BM_CLICK, 0, 0)
        if _wait_dialog_transition(
            dialog_hwnd,
            signature,
            timeout=timeout,
            require_closed=require_closed,
        ):
            return
        print(
            f"[WARN] Win32 弹窗按钮点击后没有状态变化，正在重试 "
            f"({attempt}/{attempts})"
        )
    raise SuperStrategyError("Win32 弹窗按钮点击后窗口未关闭且内容未变化")


def _confirm_order_dialog(dialog_hwnd: int, combined_text: str) -> None:
    """仅对原生文字和按钮结构均已核验的下单确认框执行 BM_CLICK。"""
    title = _native_window_text(dialog_hwnd)
    if "下单" not in title or not _is_order_confirmation(combined_text):
        raise SuperStrategyError("弹窗不是可确认的一键开仓下单确认框")

    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    cancels = _exact_child_buttons(dialog_hwnd, "取消")
    if len(confirms) != 1 or len(cancels) != 1:
        raise SuperStrategyError(
            "下单确认框按钮结构异常，已拒绝自动确认："
            f"确定={len(confirms)}, 取消={len(cancels)}"
        )

    _click_dialog_button(dialog_hwnd, confirms[0])
    print("[OK] 已用 Win32 确认经过文字和按钮双重校验的下单确认框")


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
    """用 Win32 确认成功批量结果并验证提示框关闭。"""
    result = _handle_batch_result(text)
    if result is None:
        raise SuperStrategyError("弹窗不包含可确认的批量下单成功结果")
    title = _native_window_text(dialog_hwnd)
    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    if "提示" not in title or len(confirms) != 1:
        raise SuperStrategyError(
            "批量下单成功框结构异常，已拒绝自动确认："
            f"标题={title!r}, 确定={len(confirms)}"
        )

    _click_dialog_button(dialog_hwnd, confirms[0], require_closed=True)
    print(
        "[OK] 已用 Win32 确认并关闭批量下单成功框："
        f"成功{result['success_count']}笔，失败0笔"
    )
    return result


def _dismiss_prompt_dialog(dialog_hwnd: int) -> None:
    """只用 Win32 关闭已分类、失败或未知提示框，并验证无残留。"""
    confirms = _exact_child_buttons(dialog_hwnd, "确定")
    button = confirms[0] if len(confirms) == 1 else _button_by_id(
        dialog_hwnd, AFFIRMATIVE_BUTTON_IDS
    )
    if button:
        try:
            _click_dialog_button(
                dialog_hwnd,
                button,
                attempts=2,
                timeout=0.8,
                require_closed=True,
            )
        except SuperStrategyError:
            # 部分错误框的默认按钮只改变焦点，继续用精确句柄关闭窗口。
            win32gui.PostMessage(int(dialog_hwnd), win32con.WM_CLOSE, 0, 0)
    else:
        win32gui.PostMessage(int(dialog_hwnd), win32con.WM_CLOSE, 0, 0)

    deadline = time.monotonic() + 1.5
    while time.monotonic() < deadline:
        if not _dialog_is_visible(dialog_hwnd):
            print("[OK] 已用 Win32 关闭客户端提示框")
            return
        time.sleep(0.05)
    raise SuperStrategyError("已用 Win32 处理客户端提示框，但弹窗仍未关闭")


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


def _raise_collected_result_errors(events: list[dict], *,
                                   result_labels: tuple[str, ...],
                                   expected_count: int) -> None:
    failed = [event for event in events if event.get("error") is not None]
    if not failed:
        return
    summaries = []
    for index, event in enumerate(events):
        error = event.get("error")
        if error is None:
            continue
        label = (
            result_labels[index]
            if index < len(result_labels)
            else f"第{index + 1}次操作"
        )
        if "success_count" in event and "failure_count" in event:
            summaries.append(
                f"{label}结果：成功{event['success_count']}笔，"
                f"失败{event['failure_count']}笔"
            )
        else:
            summaries.append(f"{label}：{error}")
    if len(events) < expected_count:
        summaries.append(
            f"结果弹窗链不完整（期望{expected_count}个，实际{len(events)}个）"
        )
    message = "；".join(summaries) + "；所有已出现的客户端弹窗均已关闭"
    error_types = [type(event["error"]) for event in failed]
    if LoginRequiredError in error_types:
        raise LoginRequiredError(message)
    if TradingTimeBlockedError in error_types:
        raise TradingTimeBlockedError(message)
    raise SuperStrategyError(message)


def _wait_after_action(main_hwnd: int, before_windows: set[int], *,
                       action_name: str, timeout: float = 3.0,
                       quiet_period: float = RESULT_QUIET_PERIOD,
                       minimum_success_dialogs: int = 1,
                       result_labels: tuple[str, ...] = ()) -> dict:
    """逐个处理一次业务动作的弹窗，连续静默后才允许报告成功。"""
    pid = _process_id(main_hwnd)
    deadline = time.monotonic() + timeout
    unknown: dict[int, str] = {}
    success_results: list[dict] = []
    result_events: list[dict] = []
    quiet_since: float | None = None
    while time.monotonic() < deadline:
        current = _visible_process_surfaces(pid, main_hwnd)
        pending = current - before_windows
        for hwnd in pending:
            quiet_since = None
            text = _window_text_dump(hwnd)
            try:
                _raise_for_known_prompt(text, hwnd)
            except (LoginRequiredError, TradingTimeBlockedError) as exc:
                result_events.append({"error": exc})
                unknown.pop(hwnd, None)
                # 即使第一个结果已经失败，第二个操作结果仍可能延迟出现。
                deadline = time.monotonic() + max(timeout, quiet_period + 0.2)
                time.sleep(0.05)
                break
            try:
                batch_result = _handle_batch_result(text, hwnd)
            except SuperStrategyError as exc:
                parsed = _parse_batch_result(text)
                event = {"error": exc}
                if parsed is not None:
                    event.update({
                        "success_count": parsed[0],
                        "failure_count": parsed[1],
                    })
                result_events.append(event)
                unknown.pop(hwnd, None)
                # 关闭失败框后继续监听，避免遗漏延迟生成的第二个结果框。
                deadline = time.monotonic() + max(timeout, quiet_period + 0.2)
                time.sleep(0.05)
                break
            if batch_result is not None:
                latest = _confirm_success_result_dialog(hwnd, text)
                success_results.append(latest)
                result_events.append({"result": latest})
                unknown.pop(hwnd, None)
                # 交易客户端可能在首个结果框关闭后异步创建下一层结果框。
                deadline = time.monotonic() + max(timeout, quiet_period + 0.2)
                time.sleep(0.05)
                break
            if _is_order_confirmation(text):
                _confirm_order_dialog(hwnd, text)
                unknown.pop(hwnd, None)
                # 确认后客户端需要时间生成下一层登录、停市或结果弹窗。
                deadline = time.monotonic() + timeout
                time.sleep(0.2)
                break
            unknown[hwnd] = text
        else:
            if result_events and not pending:
                now = time.monotonic()
                if quiet_since is None:
                    quiet_since = now
                if (
                    len(result_events) >= minimum_success_dialogs
                    and now - quiet_since >= quiet_period
                ):
                    print(
                        f"[OK] {action_name}后续弹窗已全部处理，"
                        f"连续 {quiet_period:.1f}s 无新弹窗"
                    )
                    _raise_collected_result_errors(
                        result_events,
                        result_labels=result_labels,
                        expected_count=minimum_success_dialogs,
                    )
                    result = dict(success_results[0])
                    result["result_dialog_count"] = len(result_events)
                    result["followup_results"] = success_results[1:]
                    return result
                time.sleep(0.05)
                continue
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
        message = f"{action_name}后出现未识别弹窗：" + "；".join(summaries) + close_note
        if result_events:
            try:
                _raise_collected_result_errors(
                    result_events,
                    result_labels=result_labels,
                    expected_count=minimum_success_dialogs,
                )
            except SuperStrategyError as exc:
                message = f"{exc}；{message}"
        raise SuperStrategyError(message)
    if result_events:
        _raise_collected_result_errors(
            result_events,
            result_labels=result_labels,
            expected_count=minimum_success_dialogs,
        )
        if len(result_events) < minimum_success_dialogs:
            raise SuperStrategyError(
                f"{action_name}结果弹窗不完整：期望至少"
                f"{minimum_success_dialogs}个，实际{len(result_events)}个"
            )
        raise SuperStrategyError(
            f"已处理{action_name}成功框，但在等待后续弹窗静默时超时"
        )
    raise SuperStrategyError(
        f"{action_name}后未检测到原生 Win32 弹窗，不能确认操作已执行"
    )


def wait_after_open(main_hwnd: int, before_windows: set[int],
                    timeout: float = 3.0,
                    quiet_period: float = RESULT_QUIET_PERIOD, *,
                    expect_add_result: bool = False) -> dict:
    """兼容入口：处理一键开仓的原生弹窗链。"""
    result = _wait_after_action(
        main_hwnd,
        before_windows,
        action_name=OPEN_POSITION_TEXT,
        timeout=timeout,
        quiet_period=quiet_period,
        minimum_success_dialogs=2 if expect_add_result else 1,
        result_labels=(
            (OPEN_POSITION_TEXT, ADD_UNDERLYING_TEXT)
            if expect_add_result else (OPEN_POSITION_TEXT,)
        ),
    )
    if expect_add_result:
        result["add_result"] = result["followup_results"][0]
    return result


def _find_combination_dialog(main_hwnd: int, timeout: float = 3.0) -> int:
    """等待并返回目标进程内新出现的“组合申报”原生对话框。"""
    pid = _process_id(main_hwnd)
    deadline = time.monotonic() + max(0.1, float(timeout))
    while time.monotonic() < deadline:
        for hwnd in _visible_process_surfaces(pid, main_hwnd):
            if "组合申报" in _native_window_text(hwnd):
                return int(hwnd)
        time.sleep(0.05)
    raise SuperStrategyError("点击“组合申报”后未检测到组合申报弹窗")


def _is_combo_dialog_alive(dlg) -> bool:
    """组合申报主对话框是否仍有效可复用。

    部分客户端（如华宝实测）在组合提交成功后会自动关闭申报对话框，旧句柄
    随之失效，继续复用会在填表时报“未找到控件(auto_id=9059)”。且 Windows
    句柄可能被新窗口复用，因此除可见性外还必须校验市场下拉框(9059)存在，
    才能判定仍可安全复用。
    """
    if not dlg:
        return False
    try:
        hwnd = int(dlg)
        if not win32gui.IsWindow(hwnd) or not win32gui.IsWindowVisible(hwnd):
            return False
        return bool(win32gui.GetDlgItem(hwnd, COMBO_DLG_MARKET_CID))
    except Exception:
        return False


def main_combination_declare() -> None:
    try:
        # 参数来自 GUI/任务中心写入的环境变量；市场/策略为逗号分隔的多选列表，
        # 合约一/合约二默认留空，由 run_combination_declare 按持仓派生（从对话框下拉候选项推导配对）。
        markets_raw = os.environ.get("GUI_COMBO_MARKET") or ""
        strategies_raw = os.environ.get("GUI_COMBO_STRATEGY") or ""
        markets = [m for m in markets_raw.split(",") if m] or [COMBO_MARKETS[0]]
        strategies = [s for s in strategies_raw.split(",") if s] or [COMBO_STRATEGIES[0]]
        qty = int(os.environ.get("GUI_ORDER_QTY") or 1)
        execute = os.environ.get("GUI_COMBO_EXECUTE", "1") != "0"
        results = run_combination_declare_all(
            markets=markets,
            strategies=strategies,
            qty=qty,
            execute=execute,
        )
        if not results:
            raise SuperStrategyError("批量组合申报全部失败，未生成任何成功记录")
    except LoginRequiredError as exc:
        print(f"[LOGIN_REQUIRED] {exc}")
        raise SystemExit(LOGIN_REQUIRED_EXIT_CODE)
    except TradingTimeBlockedError as exc:
        print(f"[TRADING_TIME_BLOCKED] {exc}")
        raise SystemExit(TRADING_TIME_BLOCKED_EXIT_CODE)
    except (SuperStrategyError, TacticsPanelError) as exc:
        print(f"[错误] {exc}")
        raise SystemExit(1)


# ====================== 组合申报（填表申报） ======================
# 超级策略 -> 组合申报 页签打开的对话框控件 ID（与 行情交易组合申报页 同族）
COMBO_DLG_MARKET_CID = 9059
COMBO_DLG_STRATEGY_CID = 9040
COMBO_DLG_CONTRACT1_CID = 9067
COMBO_DLG_CONTRACT2_CID = 9068
COMBO_DLG_QTY_CID = 9057
COMBO_DLG_SUBMIT_CID = 1
COMBO_DLG_CLEAR_CID = 2366

# 市场/策略为固定枚举（国泰海通实测抓取）
COMBO_MARKETS = ("上证", "深证")
COMBO_STRATEGIES = (
    "认购牛市价差策略",
    "认购熊市价差策略",
    "宽跨式空头策略",
    "跨式空头策略",
    "认沽牛市价差策略",
    "认沽熊市价差策略",
)

# 合约代码解析：如 "50ETF购8月2850" -> 标的=50ETF, 类型=购, 到期=8月, 行权价=2850
# under 允许数字/字母/汉字，但排除 购/沽（否则会被吞进前缀）
_CONTRACT_RE = re.compile(
    r"^(?P<under>(?:(?!购|沽)[0-9A-Za-z一-龥])+?)"
    r"(?P<type>[购沽])(?P<month>\d+月)(?P<strike>\d+)$"
)


def parse_contract(code: str) -> dict | None:
    """解析期权合约代码；无法解析返回 None。"""
    m = _CONTRACT_RE.match((code or "").strip())
    if not m:
        return None
    return {
        "under": m.group("under"),
        "type": m.group("type"),       # 购=call / 沽=put
        "month": m.group("month"),
        "strike": int(m.group("strike")),
    }


def _pair_matches(strategy: str, c1: str, c2: str) -> bool:
    """判断 (c1, c2) 是否构成所选策略要求的两条腿。"""
    p1, p2 = parse_contract(c1), parse_contract(c2)
    if not p1 or not p2:
        return False
    if p1["under"] != p2["under"] or p1["month"] != p2["month"]:
        return False

    if strategy in (
        "认购牛市价差策略", "认购熊市价差策略",
        "认沽牛市价差策略", "认沽熊市价差策略",
    ):
        # 同类型两腿、不同行权价
        want = "购" if strategy.startswith("认购") else "沽"
        if p1["type"] != want or p2["type"] != want:
            return False
        return p1["strike"] != p2["strike"]
    if strategy == "跨式空头策略":
        # 一购一沽、同行权价
        return p1["type"] != p2["type"] and p1["strike"] == p2["strike"]
    if strategy == "宽跨式空头策略":
        # 一购一沽、不同行权价
        return p1["type"] != p2["type"] and p1["strike"] != p2["strike"]
    return False


def derive_contract_pair(strategy: str, c1_opts, c2_opts):
    """从 合约一/合约二 下拉候选项中推导符合策略的配对。

    优先以 合约二 候选项为锚点（券商已按策略过滤），在 合约一 中找配对；
    否则退而在 合约一 候选项内部两两配对。找不到抛出 SuperStrategyError。
    """
    c1_opts = list(c1_opts)
    c2_opts = list(c2_opts)
    valid = []
    for c2 in c2_opts:
        for c1 in c1_opts:
            if c1 == c2:
                continue
            if _pair_matches(strategy, c1, c2):
                valid.append((c1, c2))
    if valid:
        # 价差类：合约一取较低行权价（多头腿），合约二保持锚点
        if strategy in ("认购牛市价差策略", "认沽牛市价差策略"):
            valid.sort(key=lambda p: parse_contract(p[0])["strike"])
        return valid[0]
    # 两腿都在 合约一 下拉内的情况
    for i, c1 in enumerate(c1_opts):
        for c2 in c1_opts[i + 1:]:
            if _pair_matches(strategy, c1, c2):
                return (c1, c2)
    raise SuperStrategyError(
        f"未从下拉候选项中找到符合策略'{strategy}'的合约一/合约二配对"
    )


def _combo_dlg_control(dlg, cid):
    ctrl = _combo_mod.find_visible_child(int(dlg), cid)
    if not ctrl:
        raise SuperStrategyError(f"组合申报对话框未找到控件(auto_id={cid})")
    return ctrl


def _read_combo_items(dlg, cid):
    """读取下拉框候选项；跨进程回退 UIA 遍历。选过合约一后合约二候选项会变，
    必须清缓存后重读，否则拿到旧列表。"""
    ctrl = _combo_dlg_control(dlg, cid)
    _combo_mod._combo_maps.pop(ctrl, None)
    _combo_mod._combo_sources.pop(ctrl, None)
    return _combo_mod.build_combo_map(ctrl)


def _set_combo_qty(dlg, qty):
    edit = _combo_dlg_control(dlg, COMBO_DLG_QTY_CID)
    val = str(int(qty))
    win32gui.SendMessage(edit, win32con.WM_SETTEXT, 0, val)
    if _combo_mod.get_edit_text(edit) != val:
        # 回退：先全选清除再逐字符输入，兼容自绘编辑框
        try:
            _combo_mod._send_timeout(edit, 0x00B1, 0, -1)   # EM_SETSEL 0,-1
            _combo_mod._send_timeout(edit, 0x0303, 0, 0)    # WM_CLEAR
            for ch in val:
                _combo_mod._send_timeout(edit, 0x0102, ord(ch), 1)  # WM_CHAR
        except Exception:
            pass
    if _combo_mod.get_edit_text(edit) != val:
        raise SuperStrategyError(f"组合数量设置失败: 期望={val}")
    print(f"[OK] 组合数量已设为 {val}")


def _close_combo_leftover_dialogs(main_hwnd):
    """关闭目标进程内残留的 组合申报/警告/确认/提示 模态对话框，
    避免其禁用操作面板子窗口导致真实点击被拒。"""
    pid = _process_id(main_hwnd)
    targets = []

    def cb(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if _combo_mod._dlg_pid(hwnd) != pid:
                return True
            if win32gui.GetClassName(hwnd) != "#32770":
                return True
            title = win32gui.GetWindowText(hwnd)
            if any(k in title for k in ("组合申报", "警告", "确认", "提示")):
                targets.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(cb, None)
    for d in targets:
        print(f"[INFO] 关闭残留对话框: {win32gui.GetWindowText(d)!r}")
        win32gui.PostMessage(d, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.3)


def _wait_combo_submit_dialog(main_hwnd, base_dlg, timeout=3.0):
    """等待点击申报后新出现的确认/提示弹窗（排除基础组合申报对话框）。"""
    pid = _process_id(main_hwnd)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        for h in _combo_mod.find_all_dialogs(pid):
            if h == int(base_dlg) or not win32gui.IsWindowVisible(h):
                continue
            title = win32gui.GetWindowText(h)
            if any(k in title for k in ("组合申报", "确认", "提示", "警告")):
                return h
        time.sleep(0.1)
    return 0


def _submit_combination(dlg, main_hwnd, cancel_only=False):
    """点击申报并处置后续弹窗。

    cancel_only=True 用于安全验证：点击申报后取消确认弹窗，绝不真正下单。
    否则填数量（若确认框含数量）、点确定、处理后续警告，返回是否提交成功。
    """
    btn = _combo_dlg_control(dlg, COMBO_DLG_SUBMIT_CID)
    win32gui.PostMessage(btn, win32con.BM_CLICK, 0, 0)
    print("[OK] 已点击申报（异步）")
    time.sleep(0.5)

    new_dlg = _wait_combo_submit_dialog(main_hwnd, dlg)
    if not new_dlg:
        print("[WARN] 点击申报后未检测到确认弹窗")
        return False

    if cancel_only:
        _combo_mod.cancel_dialog(new_dlg)
        print("[OK] 已取消申报（安全验证，未真正提交）")
        return False

    # 真实提交：确认框若含数量编辑框则先填数量
    if _combo_mod.has_control(new_dlg, COMBO_DLG_QTY_CID):
        _set_combo_qty(new_dlg, _combo_mod.get_edit_text(
            _combo_mod.find_visible_child(int(dlg), COMBO_DLG_QTY_CID)) or 1)
    ok = _combo_mod._find_dialog_button(new_dlg, affirmative=True)
    if not ok:
        print("[ERROR] 未找到申报确认按钮")
        return False
    win32gui.PostMessage(ok, win32con.BM_CLICK, 0, 0)
    if not _combo_mod.confirm_dialog(main_hwnd, new_dlg, attempts=3):
        print("[ERROR] 申报确认后弹窗未正常结束")
        return False
    if not _combo_mod.handle_dialogs(main_hwnd):
        print("[ERROR] 申报后续弹窗未正常结束")
        return False
    print("[OK] 组合申报已提交")
    return True


def _fill_combo_dialog(dlg, main_hwnd, *, market=None, strategy=None,
                        contract1=None, contract2=None, qty=1,
                        execute=True, close_after=False) -> dict:
    """在已打开的组合申报对话框内选市场/策略/合约、填数量，并选择性地提交。"""
    market = market or COMBO_MARKETS[0]
    if not _combo_mod.combo_select(
        _combo_dlg_control(dlg, COMBO_DLG_MARKET_CID), market
    ):
        print(f"[WARN] 市场选择失败: {market}")
    else:
        print(f"[OK] 已选择市场: {market}")

    strategy = strategy or COMBO_STRATEGIES[0]
    if not _combo_mod.combo_select(
        _combo_dlg_control(dlg, COMBO_DLG_STRATEGY_CID), strategy
    ):
        print(f"[WARN] 策略选择失败: {strategy}")
    else:
        print(f"[OK] 已选择策略: {strategy}")
    time.sleep(0.4)  # 等合约一下拉联动刷新

    c1_opts = list(_read_combo_items(dlg, COMBO_DLG_CONTRACT1_CID).keys())
    c2_opts = list(_read_combo_items(dlg, COMBO_DLG_CONTRACT2_CID).keys())
    print(f"[INFO] 合约一候选项({len(c1_opts)}): {c1_opts}")
    print(f"[INFO] 合约二候选项({len(c2_opts)}): {c2_opts}")

    if contract1 and contract2:
        c1, c2 = contract1, contract2
        if c1 not in c1_opts:
            raise SuperStrategyError(f"指定合约一不在候选项: {c1}")
    else:
        c1, c2 = derive_contract_pair(strategy, c1_opts, c2_opts)
    print(f"[INFO] 推导合约一={c1} 合约二={c2}")

    if not _combo_mod.combo_select(
        _combo_dlg_control(dlg, COMBO_DLG_CONTRACT1_CID), c1
    ):
        raise SuperStrategyError(f"合约一选择失败: {c1}")
    print(f"[OK] 已选择合约一: {c1}")
    time.sleep(0.3)

    # 选过合约一后合约二候选项可能变化，重读并重新推导
    c2_opts2 = list(_read_combo_items(dlg, COMBO_DLG_CONTRACT2_CID).keys())
    if c2 not in c2_opts2:
        print("[WARN] 选完合约一后合约二候选项变化，重新推导")
        c1, c2 = derive_contract_pair(strategy, c1_opts, c2_opts2)
        if not _combo_mod.combo_select(
            _combo_dlg_control(dlg, COMBO_DLG_CONTRACT1_CID), c1
        ):
            raise SuperStrategyError(f"合约一重选失败: {c1}")
        print(f"[OK] 已重新选择合约一: {c1}")
        time.sleep(0.3)
        c2_opts2 = list(_read_combo_items(dlg, COMBO_DLG_CONTRACT2_CID).keys())
        if c2 not in c2_opts2:
            raise SuperStrategyError(
                f"重新推导的合约二不在候选项: {c2}（可用: {c2_opts2}）"
            )

    if not _combo_mod.combo_select(
        _combo_dlg_control(dlg, COMBO_DLG_CONTRACT2_CID), c2
    ):
        raise SuperStrategyError(f"合约二选择失败: {c2}")
    print(f"[OK] 已选择合约二: {c2}")

    _set_combo_qty(dlg, qty)

    result = {
        "target": COMBINATION_DECLARE_TEXT,
        "dialog_hwnd": dlg,
        "main_hwnd": main_hwnd,
        "market": market,
        "strategy": strategy,
        "contract1": c1,
        "contract2": c2,
        "qty": qty,
        "submitted": False,
    }
    if execute:
        result["submitted"] = _submit_combination(dlg, main_hwnd)
    elif close_after:
        _close_combo_leftover_dialogs(main_hwnd)
        result["closed"] = True
    return result


def run_combination_declare(*, market=None, strategy=None, qty=1,
                            contract1=None, contract2=None,
                            execute=True, cleanup_leftover=True,
                            close_after=False, do_setup=True,
                            main_hwnd=None, combo_dlg=None) -> dict:
    """在超级策略界面填表并提交组合申报（单个 市场×策略 组合）。

    参数：
      market/strategy : 市场(上证/深证)、策略；为空则用各自首项。
      qty             : 组合数量。
      contract1/2     : 指定合约一/合约二（从下拉候选项选）；留空则按持仓派生
                         （从对话框下拉候选项按合约属性推导配对）。
      execute         : True 点击申报并提交；False 仅填表校验（安全）。
      close_after     : 填表后关闭对话框（不提交时用于清理现场，避免遗留模态框）。
      do_setup        : True 时执行倒计时/找窗/激活/切工作区（批量模式仅首次）。
      main_hwnd       : 已定位的主窗口句柄（do_setup=False 时复用，避免重复找窗）。
      combo_dlg       : 已有组合申报对话框句柄（批量模式复用，跳过关闭重开）。
    """
    client_id = os.environ.get("GUI_CLIENT_ID") or get_default_client_id()
    client = get_client(client_id)
    if not client:
        raise SuperStrategyError(f"客户端档案不存在: {client_id!r}")
    if do_setup:
        countdown_sec = max(0, int(os.environ.get("GUI_COUNTDOWN") or 3))
        countdown(countdown_sec)
        main_hwnd = find_window(client.get("window_key") or client.get("name") or "")
        _activate_main_window(main_hwnd)
        print("[OK] 已将交易客户端切换到前台")
        ensure_workspace(main_hwnd, WORKSPACE_SUPER, client)
    elif main_hwnd is None:
        main_hwnd = find_window(client.get("window_key") or client.get("name") or "")

    if combo_dlg is not None and _is_combo_dialog_alive(combo_dlg):
        # 复用已有对话框：确保前台激活后直接填表，跳过关闭/重开
        dlg = combo_dlg
        _activate_main_window(main_hwnd)
        print(f"[OK] 复用已有组合申报对话框: hwnd={dlg}")
    else:
        if combo_dlg is not None:
            # 客户端可能在提交成功后自动关闭申报对话框（华宝实测），需重新打开
            print(f"[WARN] 已有组合申报对话框已失效(hwnd={combo_dlg})，重新打开")
        if cleanup_leftover:
            _close_combo_leftover_dialogs(main_hwnd)
        print(f"[INFO] 准备点击: {COMBINATION_DECLARE_TEXT}")
        action_panel = get_action_panel(main_hwnd)
        click_action(
            main_hwnd,
            COMBINATION_DECLARE_TEXT,
            panel=action_panel,
            delay=0.8,
        )
        dlg = _find_combination_dialog(main_hwnd)
        print(f"[OK] 组合申报弹窗已出现: hwnd={dlg}")

    return _fill_combo_dialog(
        dlg, main_hwnd,
        market=market, strategy=strategy,
        contract1=contract1, contract2=contract2,
        qty=qty, execute=execute, close_after=close_after,
    )


def run_combination_declare_all(*, markets, strategies, qty=1,
                                execute=True, cleanup_leftover=True):
    """批量组合申报：对 (市场 × 策略) 的每个组合依次填表/提交。

    仅在首次做完整环境准备（倒计时/找窗/激活/切工作区 + 打开组合申报对话框），
    后续组合复用同一个对话框句柄，直接更改字段并提交，避免关闭对话框再重新打开。
    某一组合失败时记录错误并继续；全部结束后统一清理残留对话框。
    """
    import itertools

    market_list = list(markets) or [COMBO_MARKETS[0]]
    strategy_list = list(strategies) or [COMBO_STRATEGIES[0]]
    pairs = list(itertools.product(market_list, strategy_list))
    if not pairs:
        raise SuperStrategyError("未选择任何市场/策略组合")

    results = []
    errors = []
    main_hwnd = None
    combo_dlg = None  # 已打开的组合申报对话框句柄，成功打开后由后续迭代复用
    first = True      # 首个组合做完整环境准备；此后复用对话框，避免关闭重开

    for market, strategy in pairs:
        try:
            res = run_combination_declare(
                market=market, strategy=strategy, qty=qty,
                execute=execute,
                cleanup_leftover=cleanup_leftover if first else False,
                close_after=False, do_setup=first,
                main_hwnd=main_hwnd, combo_dlg=combo_dlg,
            )
            results.append(res)
            first = False
            main_hwnd = res.get("main_hwnd") or main_hwnd
            combo_dlg = res.get("dialog_hwnd") or combo_dlg
        except SuperStrategyError as exc:
            msg = f"组合失败 市场={market} 策略={strategy}: {exc}"
            print(f"[ERROR] {msg}")
            errors.append(msg)
            if first:
                # 首个组合失败：若对话框仍打开则找回复用，否则下一次自动重新完整准备
                try:
                    if main_hwnd is None:
                        client_id = os.environ.get("GUI_CLIENT_ID") or get_default_client_id()
                        client = get_client(client_id)
                        if client:
                            main_hwnd = find_window(
                                client.get("window_key") or client.get("name") or ""
                            )
                    if main_hwnd:
                        combo_dlg = _find_combination_dialog(main_hwnd, timeout=0.5)
                        first = False
                except Exception:
                    combo_dlg = None
            else:
                # 复用中的对话框如果仍可见，后续继续复用；否则标记为 None，下次自动打开
                try:
                    if not (
                        win32gui.IsWindow(combo_dlg)
                        and win32gui.IsWindowVisible(combo_dlg)
                    ):
                        combo_dlg = None
                except Exception:
                    combo_dlg = None

    # 全部结束后关闭对话框
    if main_hwnd:
        try:
            _close_combo_leftover_dialogs(main_hwnd)
        except Exception as exc:
            print(f"[WARN] 清理组合申报对话框失败: {exc}")

    if errors:
        print(f"[汇总] {len(errors)} 个组合未成功: " + "; ".join(errors))
    else:
        print(f"[OK] 批量组合申报完成: 共 {len(results)} 个组合全部处理")
    return results


def run_strategy(target: str, *, underlying: str | None = None,
                 add_underlying: bool | None = None,
                 execute_open: bool = True) -> dict:
    """执行：选择ETF标的 → 选策略 → 可选加入标的 → 一键开仓。"""
    if target not in SUPER_STRATEGY_TARGETS:
        raise ValueError(f"不支持的超级策略: {target!r}")
    if underlying is None:
        underlying = (
            os.environ.get("GUI_SUPER_UNDERLYING")
            or DEFAULT_SUPER_STRATEGY_UNDERLYING
        ).strip()
    if underlying not in SUPER_STRATEGY_UNDERLYINGS:
        raise ValueError(f"不支持的超级策略标的: {underlying!r}")
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

    print(f"[INFO] 超级策略标的: {underlying}")
    underlying_hit = select_super_underlying(main_hwnd, underlying)
    print(f"[INFO] 超级策略目标: {target}")
    menu_hit = click_tactics_item(main_hwnd, target)

    if add_underlying is None:
        add_underlying = os.environ.get(
            "GUI_SUPER_ADD_UNDERLYING", "False"
        ).strip().lower() == "true"
    action_targets = []
    if add_underlying:
        action_targets.append(ADD_UNDERLYING_TEXT)
    if execute_open:
        action_targets.append(OPEN_POSITION_TEXT)
    action_panel = get_action_panel(main_hwnd) if action_targets else None
    action_hits = (
        _find_action_texts(main_hwnd, action_panel, action_targets)
        if action_panel is not None else {}
    )
    if add_underlying:
        click_action(
            main_hwnd,
            ADD_UNDERLYING_TEXT,
            panel=action_panel,
            hit=action_hits[ADD_UNDERLYING_TEXT],
        )
    else:
        print("[INFO] 未启用“加入标的”，跳过该步骤")

    result = {
        "target": target,
        "underlying": underlying,
        "add_underlying": bool(add_underlying),
        "add_result": None,
        "underlying_hit": underlying_hit,
        "menu_hit": menu_hit,
        "open_triggered": False,
    }
    if not execute_open:
        print("[INFO] 安全验证模式：未点击一键开仓")
        return result

    pid = _process_id(main_hwnd)
    before = _visible_process_surfaces(pid, main_hwnd)
    click_action(
        main_hwnd,
        OPEN_POSITION_TEXT,
        panel=action_panel,
        hit=action_hits[OPEN_POSITION_TEXT],
    )
    result.update(wait_after_open(
        main_hwnd,
        before,
        expect_add_result=bool(add_underlying),
    ))
    if result.get("add_result") is not None:
        add_summary = result["add_result"]
        print(
            "[OK] 已分别处理两次操作结果："
            f"一键开仓成功{result['success_count']}笔，"
            f"加入标的成功{add_summary['success_count']}笔"
        )
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
