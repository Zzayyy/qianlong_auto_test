# -*- coding: utf-8 -*-
"""诊断东证期货期权宝的三个菜单 tab 对应的 TreeView 控件。

运行方式：在东证期货期权宝客户端打开的情况下，直接运行本脚本。
输出：主窗口内所有 SysTreeView32 的 control_id、可见性、位置、节点数等信息。
"""

import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

import win32gui
import win32process

from core.native_tree import get_tree_count, get_tree_root_child_counts, TVM_GETCOUNT


def find_main_window(keyword="东吴证券期权宝"):
    """查找东证主窗口"""
    candidates = []

    def _enum(hwnd, _):
        try:
            title = win32gui.GetWindowText(hwnd) or ""
            if keyword not in title or "GUI自动化工具" in title:
                return
            if not win32gui.IsWindowVisible(hwnd):
                return
            rect = win32gui.GetWindowRect(hwnd)
            if rect[2] - rect[0] < 100 or rect[3] - rect[1] < 100:
                return
            candidates.append((hwnd, title, rect))
        except Exception:
            pass

    win32gui.EnumWindows(_enum, None)
    if not candidates:
        print(f"[错误] 未找到包含 '{keyword}' 的可见窗口")
        sys.exit(1)
    # 取面积最大的
    candidates.sort(key=lambda x: (x[2][2] - x[2][0]) * (x[2][3] - x[2][1]), reverse=True)
    hwnd, title, rect = candidates[0]
    print(f"[INFO] 主窗口: hwnd={hwnd}, title={title!r}, rect={rect}")
    return hwnd


def enum_all_treeviews(parent_hwnd):
    """枚举主窗口内所有 SysTreeView32"""
    results = []

    def _enum(hwnd, _):
        try:
            class_name = win32gui.GetClassName(hwnd)
            if class_name != "SysTreeView32":
                return
            ctrl_id = win32gui.GetDlgCtrlID(hwnd)
            visible = win32gui.IsWindowVisible(hwnd)
            enabled = win32gui.IsWindowEnabled(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            width = rect[2] - rect[0]
            height = rect[3] - rect[1]
            # 相对主窗口的位置
            parent_rect = win32gui.GetWindowRect(parent_hwnd)
            rel_left = rect[0] - parent_rect[0]
            rel_top = rect[1] - parent_rect[1]

            # 尝试获取节点数
            try:
                count = get_tree_count(hwnd)
            except Exception:
                count = -1

            # 尝试获取根节点子项数
            try:
                root_counts = get_tree_root_child_counts(hwnd)
            except Exception:
                root_counts = []

            # 获取父窗口信息
            parent = win32gui.GetParent(hwnd)
            parent_class = win32gui.GetClassName(parent) if parent else ""

            results.append({
                "hwnd": hwnd,
                "ctrl_id": ctrl_id,
                "visible": visible,
                "enabled": enabled,
                "rect": rect,
                "rel_pos": (rel_left, rel_top),
                "size": (width, height),
                "node_count": count,
                "root_child_counts": root_counts,
                "parent_hwnd": parent,
                "parent_class": parent_class,
            })
        except Exception as e:
            print(f"  [WARN] 枚举 TreeView 出错: {e}")

    win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    return results


def main():
    print("=" * 60)
    print("东证期货期权宝 菜单树诊断工具")
    print("=" * 60)

    main_hwnd = find_main_window()
    print()

    # 也枚举所有 Button 控件，看看 tab 按钮
    print("--- 主窗口顶部区域的按钮（可能是 tab 切换按钮） ---")
    parent_rect = win32gui.GetWindowRect(main_hwnd)
    buttons = []

    def _enum_btn(hwnd, _):
        try:
            class_name = win32gui.GetClassName(hwnd)
            if class_name != "Button":
                return
            if not win32gui.IsWindowVisible(hwnd):
                return
            ctrl_id = win32gui.GetDlgCtrlID(hwnd)
            text = win32gui.GetWindowText(hwnd) or ""
            rect = win32gui.GetWindowRect(hwnd)
            rel_top = rect[1] - parent_rect[1]
            # 只看顶部区域的按钮（tab 栏通常在顶部）
            if rel_top < 80:
                buttons.append({
                    "hwnd": hwnd,
                    "ctrl_id": ctrl_id,
                    "text": text,
                    "rect": rect,
                    "rel_pos": (rect[0] - parent_rect[0], rel_top),
                    "size": (rect[2] - rect[0], rect[3] - rect[1]),
                })
        except Exception:
            pass

    win32gui.EnumChildWindows(main_hwnd, _enum_btn, None)
    buttons.sort(key=lambda b: (b["rel_pos"][0], b["rel_pos"][1]))
    for b in buttons:
        print(f"  hwnd={b['hwnd']} id={b['ctrl_id']} text={b['text']!r} "
              f"rel_pos={b['rel_pos']} size={b['size']}")
    print()

    # 枚举所有 TreeView
    print("--- 所有 SysTreeView32 控件 ---")
    trees = enum_all_treeviews(main_hwnd)
    print(f"共找到 {len(trees)} 个 TreeView\n")

    for i, t in enumerate(trees):
        print(f"  [{i}] hwnd={t['hwnd']}")
        print(f"      control_id={t['ctrl_id']}")
        print(f"      visible={t['visible']}  enabled={t['enabled']}")
        print(f"      绝对位置={t['rect']}  相对位置={t['rel_pos']}  尺寸={t['size']}")
        print(f"      节点数={t['node_count']}")
        print(f"      根节点子项数={t['root_child_counts']}")
        print(f"      父窗口 class={t['parent_class']!r}")
        print()

    # 按 control_id 分组
    by_id = {}
    for t in trees:
        by_id.setdefault(t["ctrl_id"], []).append(t)
    print("--- 按 control_id 分组 ---")
    for cid, group in sorted(by_id.items()):
        vis = sum(1 for t in group if t["visible"])
        print(f"  control_id={cid}: {len(group)} 个 (可见 {vis} 个)")
        for t in group:
            print(f"    hwnd={t['hwnd']} visible={t['visible']} "
                  f"nodes={t['node_count']} size={t['size']}")

    print()
    print("=" * 60)
    print("诊断完成。请切换不同的菜单 tab（期货期权/股票期权/股票）后再次运行，")
    print("对比哪些 TreeView 是随 tab 切换变化的。")
    print("=" * 60)


if __name__ == "__main__":
    main()
