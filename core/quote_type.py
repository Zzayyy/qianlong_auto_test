# -*- coding: utf-8 -*-
"""报价方式（下单面板「报价方式」下拉）的客户端差异与校验。

背景（2026-09-24 XTP 实机）：
  - 各客户端「报价方式」下拉项**不完全相同**。钱龙/国泰/中泰/华宝/广发/东吴/
    东证/联储/粤开 为 9 项，后三项是 市价转限 / 市价FAK / 市价F0K；
    中泰证券期权宝XTP 同样 9 项，但后三项换成
    全额成交或撤销 / 对方方最优转限价 / 市价剩余转撤销。
  - 该控件是「自绘可输入下拉」：UIA 暴露为 Edit，GetWindowText/WM_GETTEXT
    跨进程恒为空，只有 WM_GETTEXTLENGTH 拿得到长度。
    **但 UIA 的 ValuePattern(get_value) 能读到真实文本**（XTP 实机实测），
    因此校验优先用精确文本，读不到才退回长度比较。
  - 只比长度的旧逻辑在 XTP 上有真实风险：`全额成交或撤销` 与
    `市价剩余转撤销` 都是 7 字/14 字节，写入失败也会判为成功。

客户端差异通过 `clients.json` 的 `quote_types` 配置，运行时按 GUI_CLIENT_ID 读取；
未配置的客户端沿用内置的钱龙基准列表，行为与改动前完全一致。
"""

from __future__ import annotations

import os

# 内置基准列表（钱龙/国泰/中泰/华宝/广发/东吴/东证/联储/粤开的实机顺序）
DEFAULT_QUOTE_TYPES = (
    "对手价",
    "挂盘价",
    "涨停价",
    "跌停价",
    "限价",
    "超价",
    "市价转限",
    "市价FAK",
    "市价F0K",
)

# Excel 中常见写法 -> 客户端实际显示值。
# 注意：只做**同义/字形**归一（如 FOK 的字母 O vs 数字 0），
# 不做跨客户端语义映射（如 市价FAK → 对手方最优转限价 属业务语义，
# 是否等价需业务确认，不在本模块擅自映射）。
DEFAULT_QUOTE_TYPE_MAP = {
    "市价FOK": "市价F0K",   # 钱龙系实际控件显示 F0K(数字0)
}


def get_quote_types(client_id: str | None = None) -> tuple:
    """返回当前客户端的报价方式下拉项（按实际显示顺序）。

    优先 clients.json 的 `quote_types`；未配置则返回内置基准列表。
    """
    if client_id:
        try:
            from core.clients import get_client
            client = get_client(client_id)
        except Exception:
            client = None
        configured = (client or {}).get("quote_types")
        if isinstance(configured, list) and configured:
            values = [str(v).strip() for v in configured if str(v).strip()]
            if values:
                return tuple(values)
    return DEFAULT_QUOTE_TYPES


def get_quote_type_map(client_id: str | None = None) -> dict:
    """返回 Excel 填写值 -> 实际显示值 的映射。

    客户端可用 clients.json 的 `quote_type_map` 覆盖/补充内置映射。
    """
    result = dict(DEFAULT_QUOTE_TYPE_MAP)
    if client_id:
        try:
            from core.clients import get_client
            client = get_client(client_id)
        except Exception:
            client = None
        extra = (client or {}).get("quote_type_map")
        if isinstance(extra, dict):
            for key, value in extra.items():
                if value is None:
                    result.pop(str(key), None)
                else:
                    result[str(key)] = str(value)
    return result


def get_quote_index(client_id: str | None = None) -> dict:
    """返回 报价方式显示值 -> 下拉索引（索引兜底按位置选择时使用）。"""
    return {name: idx for idx, name in enumerate(get_quote_types(client_id))}


def env_client_id() -> str | None:
    """当前进程的客户端 id（GUI 运行时有值；独立运行通常为空）。"""
    value = (os.environ.get("GUI_CLIENT_ID") or "").strip()
    return value or None


def read_quote_text(edit_hwnd: int) -> str:
    """读取报价方式输入框当前文本（尽力而为，读不到返回空串）。

    实机要点（2026-09-24 XTP，中泰系自绘「可输入下拉」）：
      - `GetWindowText` / `WM_GETTEXT` 跨进程**恒为空**，长度只有
        `WM_GETTEXTLENGTH` 拿得到；
      - 用 `connect(handle=edit)` 或 `connect(process=pid)` 再
        `window(handle=edit)` 都拿不到 ValuePattern（报方法不存在）；
      - **只有从主窗口出发**（`connect(handle=主窗口)` 后
        `child_window(auto_id=18083)` 或 `descendants(Edit)` 过滤）才能
        读到真实文本 `get_value()`。
    因此这里从该 Edit 的顶层主窗口出发定位。
    """
    expected_id = 18083
    try:
        import win32gui
        expected_id = win32gui.GetDlgCtrlID(edit_hwnd) or expected_id
    except Exception:
        pass

    try:
        from pywinauto import Application
        main_hwnd = _top_level(edit_hwnd)
        if main_hwnd:
            app = Application(backend="uia").connect(handle=main_hwnd, timeout=2)
            win = app.window(handle=main_hwnd)

            # 1) 按 auto_id 精确取（最稳）
            try:
                el = win.child_window(auto_id=str(expected_id),
                                      control_type="Edit")
                if el.exists(timeout=1):
                    text = _value_of(el)
                    if text:
                        return text
            except Exception:
                pass

            # 2) 遍历 Edit，取 auto_id 匹配且可见的第一个
            try:
                for c in win.descendants(control_type="Edit"):
                    try:
                        if str(c.element_info.automation_id) != str(expected_id):
                            continue
                    except Exception:
                        continue
                    try:
                        if not c.is_visible():
                            continue
                    except Exception:
                        pass
                    text = _value_of(c)
                    if text:
                        return text
            except Exception:
                pass
    except Exception:
        pass

    # 3) GetWindowText 兜底（普通 Edit 有效；自绘控件通常为空）
    try:
        import win32gui
        return (win32gui.GetWindowText(edit_hwnd) or "").strip()
    except Exception:
        return ""


def _top_level(hwnd: int) -> int:
    """取该控件的顶层窗口（GA_ROOT），失败退回逐级 GetParent。"""
    try:
        import win32gui
        root = win32gui.GetAncestor(hwnd, 2)  # GA_ROOT
        if root:
            return root
    except Exception:
        pass
    try:
        import win32gui
        cur = hwnd
        while True:
            parent = win32gui.GetParent(cur)
            if not parent:
                return cur
            cur = parent
    except Exception:
        return 0


def _value_of(element) -> str:
    """尽力从 UIA 元素取文本（ValuePattern 优先）。"""
    for getter in ("get_value", "window_text"):
        try:
            text = getattr(element, getter)()
            if text:
                return str(text).strip()
        except Exception:
            continue
    try:
        name = element.element_info.name
        if name:
            return str(name).strip()
    except Exception:
        pass
    return ""


def quote_text_matches(edit_hwnd: int, expected: str) -> bool:
    """校验报价方式当前值是否等于 expected。

    能用 UIA 读到文本时做**精确比对**；确实读不到时才退回长度比较
    （字符数或 GBK 字节数任一匹配），以兼容个别读不到文本的老客户端。
    """
    expected = str(expected).strip()
    actual = read_quote_text(edit_hwnd)
    if actual:
        return actual == expected

    try:
        import win32con
        import win32gui
        n = win32gui.SendMessage(edit_hwnd, win32con.WM_GETTEXTLENGTH, 0, 0)
        return n == len(expected) or n == len(expected.encode("gbk"))
    except Exception:
        return False
