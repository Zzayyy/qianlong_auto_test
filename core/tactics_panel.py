# -*- coding: utf-8 -*-
"""超级策略(TacticsPanel)左面板 OCR 定位与点击
================================================
背景:
    项目现有三套面板定位都针对「主导航树 SysTreeView32 (control_id=1223)」:
      1) select_tree_path            -> TVM_GETITEMW 读文本        (win32)
      2) select_tree_path_by_position-> TVM_GETNEXTITEM 按位置     (win32)
      3) switch_panel 的 UIA 回退      -> pywinauto TreeItem        (UIA)
    而「超级策略」界面左侧的策略列表是一个自定义绘制的 Pane
    (title='TacticsPanel', auto_id='103'), 内部策略项既不暴露文本也不暴露
    auto_id, 上面三套方法都拿不到。本模块改用「截图 + RapidOCR」识别左面板
    可见文字, 再换算屏幕坐标点击, 正好补齐第四种定位, 并可直接读出左面板数据。

依赖(项目 requirements 已含):
    rapidocr / onnxruntime / mss / opencv-python / numpy / pywin32 / pywinauto

注意:
    真实 GUI 点击需管理员权限 + 已登录客户端, 且必须在指定客户端做受控人工验证。
"""

from __future__ import annotations

import time
import unicodedata
import difflib

import cv2
import numpy as np
import mss
import win32api
import win32con
import win32gui
from pywinauto import Application, findwindows


TACTICS_PANEL_AUTO_ID = "103"


# ----------------------------------------------------------------------------
# 基础工具(与下单/表格 OCR 脚本保持一致的实现语义)
# ----------------------------------------------------------------------------
def _normalize(value: str) -> str:
    """NFKC 归一化 + 抹平全角空格, 用于中英文菜单文本比对。"""
    return unicodedata.normalize("NFKC", value).replace("\u3000", " ").strip()


def _get_rapid_ocr():
    """初始化 RapidOCR 实例(单例)。"""
    if getattr(_get_rapid_ocr, "_instance", None) is None:
        try:
            from rapidocr import RapidOCR
        except Exception as e:  # pragma: no cover - 依赖缺失提示
            raise RuntimeError(
                f"rapidocr 导入失败: {e}；请执行 pip install rapidocr onnxruntime"
            ) from e
        _get_rapid_ocr._instance = RapidOCR()
    return _get_rapid_ocr._instance


def _grab_region(left: int, top: int, right: int, bottom: int) -> np.ndarray:
    """用 mss 截取屏幕矩形区域, 返回 RGB 的 np.ndarray。"""
    monitor = {
        "left": int(left),
        "top": int(top),
        "width": max(1, int(right - left)),
        "height": max(1, int(bottom - top)),
    }
    with mss.MSS() as sct:
        shot = sct.grab(monitor)
        img = np.array(shot, dtype=np.uint8)
    return cv2.cvtColor(img, cv2.COLOR_BGRA2RGB)


def _enlarge(img: np.ndarray, factor: float = 2.0) -> np.ndarray:
    if factor and factor != 1.0:
        h, w = img.shape[:2]
        return cv2.resize(
            img, (int(w * factor), int(h * factor)), interpolation=cv2.INTER_CUBIC
        )
    return img


def _click_screen(x: int, y: int, delay: float = 0.3):
    """绝对屏幕坐标点击(与现有 click_output_button 的鼠标方案一致)。"""
    win32api.SetCursorPos((int(x), int(y)))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
    time.sleep(delay)


def _find_scrollbar(panel_hwnd: int):
    """在面板下找 QLScrollBar 的 HWND(用于 WM_VSCROLL 滚动)。"""
    found = []
    def cb(h, _):
        try:
            if win32gui.GetClassName(h) == "QLScrollBar":
                found.append(h)
        except Exception:
            pass
    try:
        win32gui.EnumChildWindows(panel_hwnd, cb, None)
    except Exception:
        pass
    return found[0] if found else None


def _scroll_panel(panel, delta: int = -120, times: int = 1):
    """滚动超级策略左面板(向下 delta<0)。

    现场观察: 该面板对鼠标滚轮无响应, 仅 QLScrollBar 可控。
    故优先向 QLScrollBar 发 WM_VSCROLL(SB_PAGEDOWN, 一次翻一页);
    找不到滚动条时再向面板发 WM_MOUSEWHEEL 兜底。
    """
    panel_hwnd = int(panel.handle)
    sb_hwnd = _find_scrollbar(panel_hwnd)
    r = panel.rectangle()
    cx = (r.left + r.right) // 2
    cy = (r.top + r.bottom) // 2
    WM_MOUSEWHEEL = 0x020A
    for _ in range(times):
        if sb_hwnd:
            try:
                win32gui.SendMessage(
                    sb_hwnd, win32con.WM_VSCROLL, win32con.SB_PAGEDOWN, 0
                )
            except Exception:
                pass
        else:
            try:
                wparam = win32api.MAKELONG(0, delta)
                lparam = win32api.MAKELONG(cx, cy)
                win32gui.SendMessage(panel_hwnd, WM_MOUSEWHEEL, wparam, lparam)
            except Exception:
                pass
        time.sleep(0.25)


# ----------------------------------------------------------------------------
# 左面板识别
# ----------------------------------------------------------------------------
def get_tactics_panel(win):
    """返回 TacticsPanel 的 UIA Pane 元素; 找不到抛错。"""
    panel = win.child_window(auto_id=TACTICS_PANEL_AUTO_ID, control_type="Pane")
    panel.wait("ready", timeout=5)
    return panel


def _panel_rect_screen(panel):
    """取面板屏幕矩形, 返回 dict(left,top,right,bottom)。"""
    r = panel.rectangle()
    return {"left": r.left, "top": r.top, "right": r.right, "bottom": r.bottom}


def _ocr_panel_items(panel, min_conf: float = 0.30, enlarge: float = 2.0,
                     right_crop_px: int = 20):
    """对当前可见的面板区域做 OCR, 返回屏幕坐标的项列表。

    每项: {text, score, left, top, right, bottom, cx, cy}

    right_crop_px: 裁掉右侧滚动条区域(现场确认左面板唯一子控件是
        class='QLScrollBar', 位于右缘约 18px), 避免 OCR 误识箭头/滑块。
    """
    rect = _panel_rect_screen(panel)
    right = max(rect["left"] + 1, rect["right"] - right_crop_px)
    img = _grab_region(rect["left"], rect["top"], right, rect["bottom"])
    img = _enlarge(img, enlarge)

    engine = _get_rapid_ocr()
    out = engine(img)
    # 兼容多种返回格式:
    #  - rapidocr_onnxruntime: (boxes, txts, scores) 三元组
    #  - rapidocr 新版: RapidOCROutput 对象(.boxes/.txts/.scores)
    #  - 旧式 (result, elapse): result 为 [[box, text, score], ...]
    if hasattr(out, "boxes"):
        result = list(zip(out.boxes, out.txts, out.scores))
    elif isinstance(out, tuple) and len(out) == 3:
        result = list(zip(out[0], out[1], out[2]))
    elif isinstance(out, tuple) and len(out) == 2 and isinstance(out[0], list):
        result = out[0]
    else:
        result = list(out)
    # result 元素是 (box, text, score)
    items = []
    scale = enlarge or 1.0
    for box, text, score in result:
        xs = [float(p[0]) / scale for p in box]
        ys = [float(p[1]) / scale for p in box]
        left, right = min(xs), max(xs)
        top, bottom = min(ys), max(ys)
        # 图内坐标(已除回原始缩放) -> 屏幕坐标
        sx = rect["left"] + (left + right) / 2
        sy = rect["top"] + (top + bottom) / 2
        items.append({
            "text": text,
            "score": float(score),
            "left": rect["left"] + left,
            "top": rect["top"] + top,
            "right": rect["left"] + right,
            "bottom": rect["top"] + bottom,
            "cx": sx,
            "cy": sy,
        })
    items = [it for it in items if it["score"] >= min_conf]
    return items


def list_tactics_items(win, *, min_conf: float = 0.30, enlarge: float = 2.0,
                       scroll_pages: int = 0):
    """获取左面板当前可见(及向下翻页)的策略项文字列表。

    返回按出现顺序排列的项 dict 列表, 每项含 screen 坐标,
    可直接用于「读取左侧菜单数据」或后续点击。
    """
    panel = get_tactics_panel(win)
    seen: list[dict] = []
    seen_keys = set()
    for page in range(scroll_pages + 1):
        items = _ocr_panel_items(panel, min_conf=min_conf, enlarge=enlarge)
        for it in items:
            key = (it["text"], round(it["cy"]))
            if key in seen_keys:
                continue
            seen_keys.add(key)
            seen.append(it)
        if page < scroll_pages:
            _scroll_panel(panel, delta=-120, times=1)
            time.sleep(0.25)
    return seen


# ----------------------------------------------------------------------------
# 定位 + 点击
# ----------------------------------------------------------------------------
def click_tactics_item(win, target: str, *,
                       fuzzy: bool = True,
                       fuzzy_threshold: float = 0.72,
                       min_conf: float = 0.30,
                       enlarge: float = 2.0,
                       scroll_pages: int = 3,
                       click_offset: tuple = (0, 0),
                       delay: float = 0.3):
    """在超级策略左面板找到目标并点击。

    Args:
        win: 主窗口(pywinauto window)
        target: 目标策略项文字, 如 "垂直价差" / "牛市价差"
        fuzzy: 是否模糊匹配(推荐, 抗 OCR 轻微误识)
        scroll_pages: 目标不在当前视图时, 向下翻页 OCR 的最大次数
        click_offset: 点击中心相对偏移(像素), 用于微调落点
    Returns:
        命中的项 dict(含屏幕坐标); 未命中抛 RuntimeError
    """
    panel = get_tactics_panel(win)
    norm_target = _normalize(target)

    def _match(items):
        best = None
        best_ratio = 0.0
        for it in items:
            norm_text = _normalize(it["text"])
            if fuzzy:
                ratio = difflib.SequenceMatcher(None, norm_target, norm_text).ratio()
                if ratio > best_ratio:
                    best_ratio, best = ratio, it
            else:
                if norm_text == norm_target:
                    return it, 1.0
        if fuzzy and best is not None and best_ratio >= fuzzy_threshold:
            return best, best_ratio
        return None, best_ratio

    for page in range(scroll_pages + 1):
        items = _ocr_panel_items(panel, min_conf=min_conf, enlarge=enlarge)
        hit, ratio = _match(items)
        if hit is not None:
            cx = int(hit["cx"] + click_offset[0])
            cy = int(hit["cy"] + click_offset[1])
            print(
                f"[OK] 命中左面板项 {hit['text']!r} "
                f"(置信={hit['score']:.2f}, 匹配度={ratio:.2f}) -> 点击 ({cx},{cy})"
            )
            _click_screen(cx, cy, delay=delay)
            return hit
        if page < scroll_pages:
            _scroll_panel(panel, delta=-120, times=1)
            time.sleep(0.25)

    raise RuntimeError(
        f"超级策略左面板未找到 {target!r}（已翻 {scroll_pages} 页，"
        "请确认文字、或调大 scroll_pages / 检查面板 auto_id=103）"
    )


def click_tactics_item_by_index(win, index: int, *,
                                item_height: int = 30,
                                top_padding: int = 6,
                                click_offset: tuple = (0, 0),
                                delay: float = 0.3):
    """非 OCR 回退: 按行号点击(假设策略项等高排列)。

    当左面板项是图标/无文字、OCR 读不到时可用。index 从 0 开始。
    """
    panel = get_tactics_panel(win)
    rect = _panel_rect_screen(panel)
    cx = (rect["left"] + rect["right"]) // 2 + click_offset[0]
    cy = rect["top"] + top_padding + item_height * (index + 0.5) + click_offset[1]
    if cy > rect["bottom"]:
        raise RuntimeError(f"第 {index} 行超出面板范围")
    print(f"[OK] 按行号点击左面板第 {index} 项 -> ({cx},{cy})")
    _click_screen(cx, cy, delay=delay)
    return {"cx": cx, "cy": cy, "index": index}


if __name__ == "__main__":  # pragma: no cover - 手动调试入口
    import os
    import sys
    # 直接运行 core/tactics_panel.py 时, 把项目根目录加入路径, 使 `core` 包可导入
    _root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if _root not in sys.path:
        sys.path.insert(0, _root)
    from core.window import find_window, activate_window

    hwnd = find_window("钱龙模拟期权宝")
    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    win.set_focus()

    items = list_tactics_items(win, scroll_pages=2)
    print(f"左面板可见项({len(items)}个):")
    for it in items:
        print(f"  - {it['text']!r}  @({it['cx']:.0f},{it['cy']:.0f})")

    click_tactics_item(win, "买入认购")
