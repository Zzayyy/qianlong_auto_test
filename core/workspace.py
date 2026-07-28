# -*- coding: utf-8 -*-
"""钱龙/国泰海通主界面切换与到位校验。"""

from __future__ import annotations

import os
from pathlib import Path
import sys
import time

# 作为独立脚本运行时，确保项目根目录可用于导入 core 包；正常作为
# core.workspace 导入时不会重复修改 sys.path。
if __package__ in (None, ""):
    project_root = Path(__file__).resolve().parents[1]
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))

import win32con
import win32gui

from core.clients import get_client, get_default_client_id
from core.window import find_window


WORKSPACE_MARKET = "market_trade"
WORKSPACE_SUPER = "super_strategy"

MARKET_CATEGORIES = {"查询", "通知查询", "结算单", "下单", "组合申报"}
SUPER_CATEGORIES = {
    "超级策略",
    # 兼容菜单重构期间保存的定时任务/任务组分类。
    "牛市认购",
    "牛市认沽",
    "熊市认购",
    "熊市认沽",
}


class WorkspaceNavigationError(RuntimeError):
    """无法安全进入或验证目标主界面。"""


def workspace_for_category(category: str) -> str | None:
    if category in MARKET_CATEGORIES:
        return WORKSPACE_MARKET
    if category in SUPER_CATEGORIES:
        return WORKSPACE_SUPER
    return None


def _profile(client: dict | None) -> dict:
    configured = dict((client or {}).get("workspace", {}) or {})
    return {
        "main_class": configured.get("main_class", "QL_OPTION_MAINWND_CLASS"),
        "market_button_id": int(configured.get("market_button_id", 1019)),
        "super_button_id": int(configured.get("super_button_id", 1023)),
        "market_tree_id": int(configured.get("market_tree_id", 1223)),
        "tactics_panel_id": int(configured.get("tactics_panel_id", 103)),
        "tactics_scrollbar_id": int(configured.get("tactics_scrollbar_id", 104)),
    }


def _descendants(hwnd: int) -> list[int]:
    found: list[int] = []
    win32gui.EnumChildWindows(int(hwnd), lambda child, _: found.append(child), None)
    return found


def _control_id(hwnd: int) -> int | None:
    try:
        return int(win32gui.GetDlgCtrlID(hwnd))
    except Exception:
        return None


def _find_controls(main_hwnd: int, *, class_name: str | None = None,
                   control_id: int | None = None, visible: bool = True) -> list[int]:
    result = []
    for hwnd in _descendants(main_hwnd):
        try:
            if visible and not win32gui.IsWindowVisible(hwnd):
                continue
            if class_name and win32gui.GetClassName(hwnd) != class_name:
                continue
            if control_id is not None and _control_id(hwnd) != control_id:
                continue
            result.append(hwnd)
        except Exception:
            continue
    return result


def _single_control(main_hwnd: int, *, class_name: str, control_id: int) -> int:
    matches = _find_controls(
        main_hwnd, class_name=class_name, control_id=control_id, visible=True
    )
    if len(matches) != 1:
        raise WorkspaceNavigationError(
            f"控件定位不唯一: class={class_name}, id={control_id}, "
            f"匹配数={len(matches)}"
        )
    return matches[0]


def is_workspace_ready(main_hwnd: int, target: str,
                       client: dict | None = None) -> bool:
    """只读判断目标界面是否已经到位。"""
    profile = _profile(client)
    try:
        if win32gui.GetClassName(int(main_hwnd)) != profile["main_class"]:
            return False
    except Exception:
        return False

    if target == WORKSPACE_MARKET:
        return bool(_find_controls(
            main_hwnd,
            class_name="SysTreeView32",
            control_id=profile["market_tree_id"],
            visible=True,
        ))

    if target == WORKSPACE_SUPER:
        panels = _find_controls(
            main_hwnd,
            class_name="AfxWnd140u",
            control_id=profile["tactics_panel_id"],
            visible=True,
        )
        for panel in panels:
            try:
                left, top, right, bottom = win32gui.GetWindowRect(panel)
                if right - left > 500 or bottom - top < 200:
                    continue
                if win32gui.GetWindowText(panel) != "TacticsPanel":
                    continue
                scrollbars = _find_controls(
                    panel,
                    class_name="QLScrollBar",
                    control_id=profile["tactics_scrollbar_id"],
                    visible=True,
                )
                if len(scrollbars) == 1:
                    return True
            except Exception:
                continue
        return False

    raise ValueError(f"未知目标界面: {target!r}")


def ensure_workspace(main_hwnd: int, target: str, client: dict | None = None,
                     timeout: float = 5.0) -> dict:
    """用工具栏按钮切换界面，并以目标控件指纹验证到位。"""
    profile = _profile(client)
    # 最小化时自绘菜单的 PrintWindow 通常只能取得标题栏，OCR 无法安全
    # 校验文字；任务开始前恢复窗口，但不强制抢占前台焦点。
    if win32gui.IsIconic(int(main_hwnd)):
        win32gui.ShowWindow(int(main_hwnd), win32con.SW_RESTORE)
        time.sleep(0.3)
    if is_workspace_ready(main_hwnd, target, client):
        return {"changed": False, "workspace": target, "main_hwnd": int(main_hwnd)}

    button_id = (
        profile["market_button_id"]
        if target == WORKSPACE_MARKET
        else profile["super_button_id"]
    )
    button = _single_control(main_hwnd, class_name="Button", control_id=button_id)
    win32gui.SendMessage(button, win32con.BM_CLICK, 0, 0)

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if is_workspace_ready(main_hwnd, target, client):
            print(f"[OK] 已自动切换并验证界面: {target}")
            return {"changed": True, "workspace": target, "main_hwnd": int(main_hwnd)}
        time.sleep(0.1)
    raise WorkspaceNavigationError(
        f"点击工具栏按钮 id={button_id} 后未检测到目标界面 {target}"
    )


def prepare_task_workspace(category: str, client_id: str | None = None) -> dict | None:
    """任务启动前自动切换界面；交易系统设置等分类返回 None。"""
    target = workspace_for_category(category)
    if target is None:
        return None
    client_id = client_id or get_default_client_id()
    client = get_client(client_id) if client_id else None
    if not client:
        raise WorkspaceNavigationError(f"客户端档案不存在: {client_id!r}")
    hwnd = find_window(client.get("window_key") or client.get("name") or "")
    return ensure_workspace(hwnd, target, client)


def main() -> None:
    """子进程入口：按 GUI 环境变量准备界面后退出。"""
    category = os.environ.get("GUI_CATEGORY", "")
    client_id = os.environ.get("GUI_CLIENT_ID") or None
    try:
        result = prepare_task_workspace(category, client_id)
    except (WorkspaceNavigationError, RuntimeError, ValueError) as exc:
        print(f"[WORKSPACE_ERROR] {exc}")
        raise SystemExit(2)

    if result is None:
        print(f"[WORKSPACE_SKIPPED] 分类无需切换界面: {category}")
    else:
        changed = "已切换" if result.get("changed") else "已在目标界面"
        print(f"[WORKSPACE_READY] {changed}: {result.get('workspace')}")


if __name__ == "__main__":
    main()
