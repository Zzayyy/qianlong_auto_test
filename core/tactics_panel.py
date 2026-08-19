# -*- coding: utf-8 -*-
"""超级策略自绘左菜单的快速 OCR 读取与安全点击。"""

from __future__ import annotations

import ctypes
from contextlib import contextmanager
import difflib
import os
import time
import unicodedata

import numpy as np
from PIL import Image
import win32api
import win32con
import win32gui
import win32process
import win32ui

from core.clients import SUPER_STRATEGY_UNDERLYINGS


TACTICS_PANEL_CONTROL_ID = 103
TACTICS_SCROLLBAR_CONTROL_ID = 104
TACTICS_PANEL_CLASS = "AfxWnd140u"
TACTICS_PANEL_TEXT = "TacticsPanel"
UNDERLYING_CATEGORY_TEXT = "ETF期权"
UNDERLYING_CATEGORY_CONTROL_ID = 100
UNDERLYING_SELECTOR_CONTROL_ID = 101
UNDERLYING_CONTROL_CLASS = "Static"
UNDERLYING_POPUP_CLASS = "CQlComboDropList"
SUPER_TARGETS = {
    "牛市认购",
    "牛市认沽",
    "熊市认购",
    "熊市认沽",
    "卖出跨式",
    "卖宽跨式",
}
# 菜单复位到顶部后，六个正式目标所在的逻辑像素单元格。位置只用于裁出
# 单个文字块；是否可点击仍必须由 OCR 对该块进行精确文字校验。
FORMAL_TARGET_CELLS = {
    "牛市认购": (5, 200, 80, 242),
    "牛市认沽": (82, 200, 158, 242),
    "熊市认购": (5, 380, 80, 422),
    "熊市认沽": (82, 380, 158, 422),
    "卖出跨式": (5, 540, 80, 582),
    "卖宽跨式": (82, 540, 158, 582),
}


class TacticsPanelError(RuntimeError):
    """超级策略左菜单无法被唯一、安全地读取。"""


@contextmanager
def dpi_unaware():
    """临时使用 DPI-unaware 坐标，使 PrintWindow 与控件矩形保持同一尺度。"""
    user32 = ctypes.windll.user32
    setter = getattr(user32, "SetThreadDpiAwarenessContext", None)
    if setter is None:  # pragma: no cover - 旧版 Windows
        yield
        return
    setter.restype = ctypes.c_void_p
    setter.argtypes = [ctypes.c_void_p]
    old = setter(ctypes.c_void_p(-1))  # DPI_AWARENESS_CONTEXT_UNAWARE
    try:
        yield
    finally:
        if old:
            setter(ctypes.c_void_p(old))


def _normalize(value: str) -> str:
    return (
        unicodedata.normalize("NFKC", str(value or ""))
        .replace("\u3000", " ")
        .strip()
        .casefold()
    )


def _as_hwnd(window) -> int:
    if isinstance(window, int):
        return int(window)
    return int(window.handle)


def _enum_descendants(hwnd: int) -> list[int]:
    found: list[int] = []
    win32gui.EnumChildWindows(int(hwnd), lambda child, _: found.append(child), None)
    return found


def _control_id(hwnd: int) -> int | None:
    try:
        return int(win32gui.GetDlgCtrlID(hwnd))
    except Exception:
        return None


def _find_children(hwnd: int, *, class_name: str | None = None,
                   control_id: int | None = None, visible: bool = True) -> list[int]:
    result = []
    for child in _enum_descendants(hwnd):
        try:
            if visible and not win32gui.IsWindowVisible(child):
                continue
            if class_name and win32gui.GetClassName(child) != class_name:
                continue
            if control_id is not None and _control_id(child) != control_id:
                continue
            result.append(child)
        except Exception:
            continue
    return result


def get_tactics_panel(window) -> int:
    """返回唯一、可见且尺寸合理的 TacticsPanel HWND。"""
    main_hwnd = _as_hwnd(window)
    candidates = _find_children(
        main_hwnd,
        class_name=TACTICS_PANEL_CLASS,
        control_id=TACTICS_PANEL_CONTROL_ID,
        visible=True,
    )
    valid = []
    for hwnd in candidates:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        if (
            win32gui.GetWindowText(hwnd) == TACTICS_PANEL_TEXT
            and 100 <= right - left <= 500
            and bottom - top >= 200
        ):
            valid.append(hwnd)
    if len(valid) != 1:
        raise TacticsPanelError(
            f"TacticsPanel 定位不唯一（匹配数={len(valid)}）；请先进入超级策略界面"
        )
    return valid[0]


def get_tactics_scrollbar(panel_hwnd: int) -> int:
    matches = _find_children(
        panel_hwnd,
        class_name="QLScrollBar",
        control_id=TACTICS_SCROLLBAR_CONTROL_ID,
        visible=True,
    )
    if len(matches) != 1:
        raise TacticsPanelError(f"超级策略滚动条定位不唯一（匹配数={len(matches)}）")
    return matches[0]


def _get_rapid_ocr():
    """进程内只初始化一次 OCR 模型。"""
    instance = getattr(_get_rapid_ocr, "_instance", None)
    if instance is None:
        try:
            from rapidocr import RapidOCR
        except Exception as exc:  # pragma: no cover - 部署依赖错误
            raise TacticsPanelError(
                f"rapidocr 导入失败: {exc}；请安装 rapidocr 与 onnxruntime"
            ) from exc
        instance = RapidOCR()
        _get_rapid_ocr._instance = instance
    return instance


def capture_window_image(hwnd: int) -> Image.Image:
    """用 PrintWindow 截取 HWND；不依赖前台焦点，也不怕窗口被遮挡。"""
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(int(hwnd))
        width, height = right - left, bottom - top
        if width <= 0 or height <= 0:
            raise TacticsPanelError(f"窗口矩形无效: {(left, top, right, bottom)}")
        window_dc = win32gui.GetWindowDC(int(hwnd))
        source_dc = win32ui.CreateDCFromHandle(window_dc)
        memory_dc = source_dc.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        try:
            bitmap.CreateCompatibleBitmap(source_dc, width, height)
            memory_dc.SelectObject(bitmap)
            if not ctypes.windll.user32.PrintWindow(
                int(hwnd), memory_dc.GetSafeHdc(), 2
            ):
                raise TacticsPanelError("主窗口 PrintWindow 截图失败")
            info = bitmap.GetInfo()
            bits = bitmap.GetBitmapBits(True)
            return Image.frombuffer(
                "RGB",
                (info["bmWidth"], info["bmHeight"]),
                bits,
                "raw",
                "BGRX",
                0,
                1,
            ).copy()
        finally:
            win32gui.DeleteObject(bitmap.GetHandle())
            memory_dc.DeleteDC()
            source_dc.DeleteDC()
            win32gui.ReleaseDC(int(hwnd), window_dc)


def crop_child_from_main(main_image: Image.Image, main_hwnd: int, child_hwnd: int,
                         *, right_crop: int = 0, top_limit: int | None = None,
                         top_offset: int = 0,
                         bottom_extra: int = 0) -> tuple[Image.Image, tuple[int, int]]:
    """按 Win32 矩形从主窗口截图裁出子控件，并返回屏幕坐标原点。"""
    with dpi_unaware():
        main_left, main_top, _, _ = win32gui.GetWindowRect(int(main_hwnd))
        left, top, right, bottom = win32gui.GetWindowRect(int(child_hwnd))
    width = max(1, right - left - max(0, right_crop))
    top_offset = max(0, int(top_offset))
    height = max(1, bottom - top - top_offset + max(0, bottom_extra))
    if top_limit is not None:
        height = min(height, max(1, int(top_limit)))
    x0, y0 = left - main_left, top - main_top + top_offset
    x1 = min(main_image.width, x0 + width)
    y1 = min(main_image.height, y0 + height)
    if x0 < 0 or y0 < 0 or x1 <= x0 or y1 <= y0:
        raise TacticsPanelError(
            f"子控件裁剪越界: main={main_image.size}, child={(x0, y0, x1, y1)}"
        )
    return main_image.crop((x0, y0, x1, y1)), (left, top + top_offset)


def _parse_ocr_output(output) -> list[tuple[object, str, float]]:
    if output is None:
        return []
    if hasattr(output, "boxes"):
        if output.boxes is None:
            return []
        return list(zip(output.boxes, output.txts, output.scores))
    if isinstance(output, tuple) and len(output) == 3:
        return list(zip(output[0], output[1], output[2]))
    if isinstance(output, tuple) and len(output) == 2:
        return list(output[0] or [])
    return list(output)


def ocr_image_items(image: Image.Image, *, screen_origin: tuple[int, int] = (0, 0),
                    min_conf: float = 0.80, enlarge: float = 3.0,
                    use_cls: bool = True) -> list[dict]:
    """OCR 一张小区域截图并把结果换算为屏幕坐标。"""
    if enlarge and enlarge != 1.0:
        image = image.resize(
            (max(1, int(image.width * enlarge)), max(1, int(image.height * enlarge))),
            Image.Resampling.LANCZOS,
        )
    started = time.perf_counter()
    output = _get_rapid_ocr()(
        np.asarray(image), use_det=True, use_cls=bool(use_cls), use_rec=True
    )
    elapsed = time.perf_counter() - started
    scale = float(enlarge or 1.0)
    result = []
    for box, text, score in _parse_ocr_output(output):
        if float(score) < min_conf:
            continue
        xs = [float(point[0]) / scale for point in box]
        ys = [float(point[1]) / scale for point in box]
        left, right = min(xs), max(xs)
        top, bottom = min(ys), max(ys)
        result.append({
            "text": str(text),
            "score": float(score),
            "left": screen_origin[0] + left,
            "top": screen_origin[1] + top,
            "right": screen_origin[0] + right,
            "bottom": screen_origin[1] + bottom,
            "cx": screen_origin[0] + (left + right) / 2,
            "cy": screen_origin[1] + (top + bottom) / 2,
            "ocr_elapsed": elapsed,
        })
    return result


def ocr_single_line(image: Image.Image, *, min_conf: float = 0.80,
                    enlarge: float = 3.0) -> dict | None:
    """识别一个已经裁好的横向文字块，不再运行耗时的文字检测。"""
    if enlarge and enlarge != 1.0:
        image = image.resize(
            (max(1, int(image.width * enlarge)), max(1, int(image.height * enlarge))),
            Image.Resampling.LANCZOS,
        )
    started = time.perf_counter()
    output = _get_rapid_ocr()(
        np.asarray(image), use_det=False, use_cls=False, use_rec=True
    )
    elapsed = time.perf_counter() - started
    texts = list(getattr(output, "txts", ()) or ())
    scores = list(getattr(output, "scores", ()) or ())
    if len(texts) != 1 or len(scores) != 1 or float(scores[0]) < min_conf:
        return None
    return {
        "text": str(texts[0]),
        "score": float(scores[0]),
        "ocr_elapsed": elapsed,
    }


def _click_client(hwnd: int, x: int, y: int, *, delay: float = 0.25) -> None:
    if not win32gui.IsWindow(int(hwnd)) or not win32gui.IsWindowEnabled(int(hwnd)):
        raise TacticsPanelError(f"目标控件不可点击: hwnd={hwnd}")
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetClientRect(int(hwnd))
        if not (left <= x < right and top <= y < bottom):
            raise TacticsPanelError(
                f"点击坐标越界: hwnd={hwnd}, point={(x, y)}, rect={(left, top, right, bottom)}"
            )
        lparam = win32api.MAKELONG(int(x), int(y))
        win32gui.SendMessage(
            int(hwnd), win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam
        )
        win32gui.SendMessage(int(hwnd), win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(delay)


def reset_tactics_scroll(panel_hwnd: int, clicks: int = 32) -> None:
    """点击自绘 QLScrollBar 的向上箭头，确定性回到菜单顶部。"""
    scrollbar = get_tactics_scrollbar(panel_hwnd)
    with dpi_unaware():
        _, _, right, bottom = win32gui.GetClientRect(scrollbar)
    x, y = max(1, right // 2), max(1, min(bottom - 1, 8))
    lparam = win32api.MAKELONG(x, y)
    for _ in range(max(0, int(clicks))):
        win32gui.SendMessage(
            scrollbar, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam
        )
        win32gui.SendMessage(scrollbar, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.2)


def _page_down(panel_hwnd: int, clicks: int = 18) -> None:
    scrollbar = get_tactics_scrollbar(panel_hwnd)
    with dpi_unaware():
        _, _, right, bottom = win32gui.GetClientRect(scrollbar)
    x, y = max(1, right // 2), max(1, bottom - 8)
    lparam = win32api.MAKELONG(x, y)
    for _ in range(max(1, int(clicks))):
        win32gui.PostMessage(
            scrollbar, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, lparam
        )
        win32gui.PostMessage(scrollbar, win32con.WM_LBUTTONUP, 0, lparam)
    time.sleep(0.35)


def _ocr_current_panel(main_hwnd: int, panel_hwnd: int, *,
                       top_limit: int | None = None,
                       top_offset: int = 0,
                       min_conf: float = 0.80,
                       enlarge: float = 2.5) -> tuple[list[dict], Image.Image]:
    main_image = capture_window_image(main_hwnd)
    panel_image, origin = crop_child_from_main(
        main_image,
        main_hwnd,
        panel_hwnd,
        right_crop=20,
        top_limit=top_limit,
        top_offset=top_offset,
    )
    return (
        ocr_image_items(
            panel_image,
            screen_origin=origin,
            min_conf=min_conf,
            enlarge=enlarge,
        ),
        panel_image,
    )


def get_underlying_controls(window) -> tuple[int, int]:
    """定位固定在策略列表上方的类别和标的 Static 控件。"""
    main_hwnd = _as_hwnd(window)
    controls = []
    for control_id in (
        UNDERLYING_CATEGORY_CONTROL_ID,
        UNDERLYING_SELECTOR_CONTROL_ID,
    ):
        matches = _find_children(
            main_hwnd,
            class_name=UNDERLYING_CONTROL_CLASS,
            control_id=control_id,
            visible=True,
        )
        valid = []
        for hwnd in matches:
            with dpi_unaware():
                left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            if 100 <= right - left <= 300 and 15 <= bottom - top <= 40:
                valid.append(hwnd)
        if len(valid) != 1:
            raise TacticsPanelError(
                f"超级策略标的控件定位不唯一（control_id={control_id}, "
                f"匹配数={len(valid)}）"
            )
        controls.append(valid[0])
    category_hwnd, selector_hwnd = controls
    if win32gui.GetParent(category_hwnd) != win32gui.GetParent(selector_hwnd):
        raise TacticsPanelError("超级策略类别与标的控件不属于同一页面")
    return category_hwnd, selector_hwnd


def _recognize_underlying_control(main_hwnd: int, control_hwnd: int, *,
                                  min_conf: float, enlarge: float) -> dict:
    main_image = capture_window_image(main_hwnd)
    control_image, _ = crop_child_from_main(
        main_image, main_hwnd, control_hwnd, right_crop=16
    )
    hit = ocr_single_line(control_image, min_conf=min_conf, enlarge=enlarge)
    if hit is None:
        artifact = _save_failure(control_image, "标的固定选择框")
        raise TacticsPanelError(f"超级策略标的固定选择框OCR失败；截图={artifact}")
    return hit


def _real_click_target(target_hwnd: int, screen_x: float, screen_y: float, *,
                       delay: float) -> None:
    """在已校验的固定控件或下拉弹窗内执行真实鼠标点击。"""
    x, y = int(round(screen_x)), int(round(screen_y))
    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(int(target_hwnd))
        if not (left <= x < right and top <= y < bottom):
            raise TacticsPanelError(
                f"真实点击坐标越界: point={(x, y)}, rect={(left, top, right, bottom)}"
            )
        point_hwnd = win32gui.WindowFromPoint((x, y))
        if point_hwnd != int(target_hwnd) and not win32gui.IsChild(
            int(target_hwnd), point_hwnd
        ):
            raise TacticsPanelError("超级策略标的点击位置被其他窗口遮挡")
        win32api.SetCursorPos((x, y))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(delay)


def _underlying_popups(main_hwnd: int) -> list[int]:
    pid = int(win32process.GetWindowThreadProcessId(int(main_hwnd))[1])
    matches = []

    def callback(hwnd, _):
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            if win32gui.GetClassName(hwnd) != UNDERLYING_POPUP_CLASS:
                return True
            if int(win32process.GetWindowThreadProcessId(hwnd)[1]) != pid:
                return True
            if win32gui.GetWindow(hwnd, win32con.GW_OWNER) != int(main_hwnd):
                return True
            matches.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(callback, None)
    return matches


def _open_underlying_popup(main_hwnd: int, selector_hwnd: int,
                           timeout: float = 1.5) -> int:
    existing = _underlying_popups(main_hwnd)
    if len(existing) > 1:
        raise TacticsPanelError(f"ETF下拉弹窗不唯一（匹配数={len(existing)}）")
    if len(existing) == 1:
        return existing[0]

    with dpi_unaware():
        left, top, right, bottom = win32gui.GetWindowRect(selector_hwnd)
    _real_click_target(selector_hwnd, right - 6, (top + bottom) / 2, delay=0.2)
    deadline = time.monotonic() + max(0.1, float(timeout))
    while time.monotonic() < deadline:
        matches = _underlying_popups(main_hwnd)
        if len(matches) == 1:
            return matches[0]
        if len(matches) > 1:
            raise TacticsPanelError(f"ETF下拉弹窗不唯一（匹配数={len(matches)}）")
        time.sleep(0.05)
    raise TacticsPanelError("点击固定标的选择框后未出现ETF下拉弹窗")


def _close_underlying_popup(popup_hwnd: int) -> None:
    if win32gui.IsWindow(popup_hwnd):
        win32gui.PostMessage(popup_hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(0.15)


def _wait_underlying_popup_closed(popup_hwnd: int,
                                  timeout: float = 1.0) -> None:
    """等待标的下拉弹窗销毁或隐藏，不再截图复核主窗口。"""
    deadline = time.monotonic() + max(0.1, float(timeout))
    while time.monotonic() < deadline:
        if not win32gui.IsWindow(popup_hwnd):
            return
        if not win32gui.IsWindowVisible(popup_hwnd):
            return
        time.sleep(0.05)
    raise TacticsPanelError("点击ETF标的后下拉弹窗未关闭")


def select_super_underlying(window, target: str, *, min_conf: float = 0.80,
                            enlarge: float = 2.5,
                            delay: float = 0.5) -> dict:
    """读取固定选择框，展开独立弹窗并精确选择ETF标的。"""
    if target not in SUPER_STRATEGY_UNDERLYINGS:
        raise ValueError(f"不支持的超级策略标的: {target!r}")

    main_hwnd = _as_hwnd(window)
    category_hwnd, selector_hwnd = get_underlying_controls(main_hwnd)
    category = _recognize_underlying_control(
        main_hwnd, category_hwnd, min_conf=min_conf, enlarge=enlarge
    )
    if _normalize(category["text"]) != _normalize(UNDERLYING_CATEGORY_TEXT):
        raise TacticsPanelError(
            f"超级策略类别校验失败：期望={UNDERLYING_CATEGORY_TEXT!r}，"
            f"实际={category['text']!r}"
        )
    selected = _recognize_underlying_control(
        main_hwnd, selector_hwnd, min_conf=min_conf, enlarge=enlarge
    )

    if _normalize(selected["text"]) == _normalize(target):
        selected = dict(selected)
        selected["mode"] = "already_selected"
        print(f"[OK] 超级策略标的已是 {target!r}，无需切换")
        return selected

    popup_hwnd = _open_underlying_popup(main_hwnd, selector_hwnd)
    try:
        popup_image = capture_window_image(popup_hwnd)
        with dpi_unaware():
            popup_left, popup_top, _, _ = win32gui.GetWindowRect(popup_hwnd)
        list_items = ocr_image_items(
            popup_image,
            screen_origin=(popup_left, popup_top),
            min_conf=min_conf,
            enlarge=enlarge,
        )
        list_hits = [
            item for item in list_items
            if _normalize(item["text"]) == _normalize(target)
        ]
        if len(list_hits) != 1:
            artifact = _save_failure(popup_image, f"标的_{target}")
            raise TacticsPanelError(
                f"ETF下拉列表未唯一识别到 {target!r}"
                f"（匹配数={len(list_hits)}）；截图={artifact}"
            )

        hit = list_hits[0]
        _real_click_target(popup_hwnd, hit["cx"], hit["cy"], delay=delay)
        _wait_underlying_popup_closed(popup_hwnd)
    except Exception:
        _close_underlying_popup(popup_hwnd)
        raise

    result = dict(hit)
    result["text"] = target
    result["mode"] = "selected_from_dropdown"
    print(
        f"[OK] OCR 选择超级策略标的 {target!r}，下拉弹窗已关闭: "
        f"置信度={hit['score']:.3f}"
    )
    return result


def list_tactics_items(window, *, min_conf: float = 0.80, enlarge: float = 2.5,
                       scroll_pages: int = 0, reset_to_top: bool = True) -> list[dict]:
    """读取左菜单；截图来自主窗口，因此窗口被遮挡时也能工作。"""
    main_hwnd = _as_hwnd(window)
    panel_hwnd = get_tactics_panel(main_hwnd)
    if reset_to_top:
        reset_tactics_scroll(panel_hwnd)
    seen = set()
    result = []
    for page in range(max(0, int(scroll_pages)) + 1):
        items, _ = _ocr_current_panel(
            main_hwnd, panel_hwnd, min_conf=min_conf, enlarge=enlarge
        )
        for item in items:
            key = _normalize(item["text"])
            if key and key not in seen:
                seen.add(key)
                result.append(item)
        if page < scroll_pages:
            _page_down(panel_hwnd)
    return result


def _match_target(items: list[dict], target: str, *, fuzzy_threshold: float) -> dict | None:
    normalized = _normalize(target)
    exact = [item for item in items if _normalize(item["text"]) == normalized]
    if len(exact) == 1:
        return exact[0]
    if len(exact) > 1:
        raise TacticsPanelError(f"OCR 重复识别到 {target!r}，拒绝选择")

    ranked = []
    for item in items:
        ratio = difflib.SequenceMatcher(
            None, normalized, _normalize(item["text"])
        ).ratio()
        if ratio >= fuzzy_threshold:
            ranked.append((ratio, item))
    ranked.sort(key=lambda pair: pair[0], reverse=True)
    if not ranked:
        return None
    if len(ranked) > 1 and abs(ranked[0][0] - ranked[1][0]) < 0.03:
        raise TacticsPanelError(f"{target!r} 存在多个相近 OCR 结果，拒绝选择")
    return ranked[0][1]


def _scan_formal_target(main_hwnd: int, panel_hwnd: int, target: str,
                        min_conf: float = 0.80
                        ) -> tuple[dict | None, Image.Image, tuple[int, int]]:
    """识别单个已知菜单格，跳过耗时的整页文字检测，同时返回 panel 截图。

    OCR 与后续高亮判定共用同一张 panel_image，避免重复 PrintWindow。
    """
    rect = FORMAL_TARGET_CELLS[target]
    main_image = capture_window_image(main_hwnd)
    panel_image, origin = crop_child_from_main(
        main_image, main_hwnd, panel_hwnd, right_crop=20
    )
    left, top, right, bottom = rect
    if right > panel_image.width or bottom > panel_image.height:
        return None, panel_image, origin
    cell = panel_image.crop(rect).resize(
        ((right - left) * 3, (bottom - top) * 3), Image.Resampling.LANCZOS
    )
    started = time.perf_counter()
    output = _get_rapid_ocr()(
        np.asarray(cell), use_det=False, use_cls=False, use_rec=True
    )
    elapsed = time.perf_counter() - started
    texts = list(getattr(output, "txts", ()) or ())
    scores = list(getattr(output, "scores", ()) or ())
    if len(texts) != 1 or len(scores) != 1:
        return None, panel_image, origin
    if _normalize(texts[0]) != _normalize(target) or float(scores[0]) < min_conf:
        print(
            f"[WARN] 快速菜单格校验未通过: 期望={target!r}, "
            f"识别={texts[0]!r}, 置信度={float(scores[0]):.3f}"
        )
        return None, panel_image, origin
    return {
        "text": str(texts[0]),
        "score": float(scores[0]),
        "left": origin[0] + left,
        "top": origin[1] + top,
        "right": origin[0] + right,
        "bottom": origin[1] + bottom,
        "cx": origin[0] + (left + right) / 2,
        "cy": origin[1] + (top + bottom) / 2,
        "ocr_elapsed": elapsed,
        "mode": "recognition_only",
    }, panel_image, origin


# 选中态高亮判定：客户端选中菜单时，整行被高亮颜色填充（钱龙是饱和蓝、
# 部分客户端是其它色相），导致 cell 内"亮像素(brightness > 130)"占比明显
# 高于其他 cell。当目标 cell 的占比远高于其他 5 个 cell 时即为已选中态。
# 阈值目前是全局统一值（实测选中≈0.72、未选中≈0.07，区分度足够），暂不
# 按客户端参数化；若未来某客户端改用"暗色反白"之类非亮色高亮再单独配置。
_FORMAL_HIGHLIGHT_BRIGHTNESS = 130
_FORMAL_HIGHLIGHT_MIN_PCT = 0.30
_FORMAL_HIGHLIGHT_DELTA_PCT = 0.15


def _cell_bright_pct(cell_pixels: np.ndarray) -> float:
    """统计单元格里亮像素占比；空 cell 返回 0。"""
    rgb = cell_pixels[..., :3].astype(np.int32)
    if rgb.size == 0:
        return 0.0
    brightness = rgb.max(axis=-1)
    return float((brightness > _FORMAL_HIGHLIGHT_BRIGHTNESS).mean())


def _detect_selected_formal_target(panel_image: Image.Image
                                   ) -> tuple[str | None, dict[str, float]]:
    """根据 6 个 FORMAL_TARGET_CELLS 的亮像素占比，判定当前是否已有选中项。

    唯一显著高于其他 cell 的项才返回，避免两个 cell 都被高亮的客户端误判。
    调用方应确保 panel_image 完整覆盖 FORMAL_TARGET_CELLS（_scan_formal_target
    越界时已返回 hit=None），因此这里对越界一律按"无选中"处理。
    """
    pct_map: dict[str, float] = {}
    for name, rect in FORMAL_TARGET_CELLS.items():
        left, top, right, bottom = rect
        if right > panel_image.width or bottom > panel_image.height:
            return None, {}
        crop = panel_image.crop(rect).convert("RGB")
        pct_map[name] = _cell_bright_pct(np.asarray(crop))
    ranked = sorted(pct_map.items(), key=lambda kv: kv[1], reverse=True)
    top_target, top_pct = ranked[0]
    second_pct = ranked[1][1] if len(ranked) > 1 else 0.0
    if (
        top_pct >= _FORMAL_HIGHLIGHT_MIN_PCT
        and top_pct - second_pct >= _FORMAL_HIGHLIGHT_DELTA_PCT
    ):
        return top_target, pct_map
    return None, pct_map


def _save_failure(image: Image.Image, target: str) -> str:
    directory = os.environ.get("GUI_SUPER_ARTIFACT_DIR") or os.path.join(
        os.environ.get("TEMP") or os.getcwd(), "qianlong_auto_super_strategy"
    )
    os.makedirs(directory, exist_ok=True)
    safe_target = "".join(ch for ch in target if ch.isalnum()) or "unknown"
    path = os.path.join(
        directory,
        f"菜单识别失败_{safe_target}_{time.strftime('%Y%m%d_%H%M%S')}.png",
    )
    image.save(path)
    return path


def click_tactics_item(window, target: str, *, min_conf: float = 0.80,
                       enlarge: float = 2.5, fuzzy_threshold: float = 0.88,
                       scroll_pages: int = 0, delay: float = 0.5) -> dict:
    """OCR 唯一命中目标后，以面板消息点击；不会使用裸屏幕坐标。"""
    main_hwnd = _as_hwnd(window)
    panel_hwnd = get_tactics_panel(main_hwnd)
    reset_tactics_scroll(panel_hwnd)
    if target in FORMAL_TARGET_CELLS:
        hit, panel_image, _ = _scan_formal_target(
            main_hwnd, panel_hwnd, target, min_conf=min_conf
        )
        if hit is not None:
            # 客户端选中态由整行高亮颜色表达，若 OCR 通过时目标已处于选中态，
            # 再点击会取消选择；必须先采样背景颜色并跳过点击。
            selected, pct_map = _detect_selected_formal_target(panel_image)
            if selected == target:
                hit = dict(hit)
                hit["mode"] = "already_selected"
                hit["highlight_pct_map"] = pct_map
                print(
                    f"[OK] {target!r} 已处于选中态(亮像素占比={pct_map}), "
                    "跳过点击以免取消"
                )
                return hit
            with dpi_unaware():
                left, top, _, _ = win32gui.GetWindowRect(panel_hwnd)
            _click_client(
                panel_hwnd,
                int(round(hit["cx"] - left)),
                int(round(hit["cy"] - top)),
                delay=delay,
            )
            print(
                f"[OK] OCR 快速校验并点击 {target!r}: "
                f"置信度={hit['score']:.3f}, OCR={hit['ocr_elapsed']:.2f}s"
            )
            return hit

        # 自绘滚动条偶尔会丢失一轮消息；再次复位后才进入整页检测回退。
        reset_tactics_scroll(panel_hwnd)
    last_image = None
    for page in range(max(0, int(scroll_pages)) + 1):
        # 当前六个正式策略都位于顶部 620px 内，缩小 OCR 区域可显著提速。
        formal_target = page == 0 and target in SUPER_TARGETS
        top_offset = 160 if formal_target else 0
        top_limit = 460 if formal_target else None
        items, last_image = _ocr_current_panel(
            main_hwnd,
            panel_hwnd,
            top_limit=top_limit,
            top_offset=top_offset,
            min_conf=min_conf,
            enlarge=enlarge,
        )
        hit = _match_target(items, target, fuzzy_threshold=fuzzy_threshold)
        if hit is not None:
            with dpi_unaware():
                left, top, _, _ = win32gui.GetWindowRect(panel_hwnd)
            x, y = int(round(hit["cx"] - left)), int(round(hit["cy"] - top))
            _click_client(panel_hwnd, x, y, delay=delay)
            print(
                f"[OK] OCR 命中并点击 {hit['text']!r}: "
                f"置信度={hit['score']:.3f}, OCR={hit['ocr_elapsed']:.2f}s"
            )
            return hit
        if page < scroll_pages:
            _page_down(panel_hwnd)

    artifact = _save_failure(last_image, target) if last_image is not None else ""
    raise TacticsPanelError(
        f"超级策略左菜单未唯一识别到 {target!r}；截图={artifact or '无'}"
    )
