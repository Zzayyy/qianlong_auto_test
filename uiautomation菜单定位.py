import uiautomation as auto
import win32gui
import ctypes
import time
from pywinauto import findwindows


def find_treeview(parent_hwnd, control_id=1223):
    hwnd = ctypes.windll.user32.GetDlgItem(parent_hwnd, control_id)
    if hwnd:
        cls = win32gui.GetClassName(hwnd)
        print(f"✓ TreeView: {hex(hwnd)} (ID={control_id}, 类名={cls})")
        return hwnd
    found = []
    def _enum(h, _):
        if win32gui.GetClassName(h) == "SysTreeView32":
            found.append(h)
        return True
    win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    return found[0] if found else 0


def get_tree(tree_hwnd):
    return auto.ControlFromHandle(tree_hwnd)


def dump_tree(node, level=0):
    for child in node.GetChildren():
        name = child.Name or "(无文本)"
        print("  " * level + name)
        dump_tree(child, level + 1)


def select_node(tree, path):
    """按路径选择节点"""
    current = tree
    for depth, target in enumerate(path):
        found = False
        for child in current.GetChildren():
            if child.Name and child.Name.strip() == target.strip():
                current = child
                found = True
                break
        if not found:
            print(f"✗ 第{depth}层未找到: '{target}'")
            return None
        # 展开（非最后一层）
        if depth < len(path) - 1:
            try:
                current.Expand()
            except Exception:
                pass
            time.sleep(0.15)

    # ★ 选中：用 Click 代替 Select ★
    current.SetFocus()
    time.sleep(0.05)
    current.Click(simulateMove=False)
    time.sleep(0.05)
    print(f"✓ 已选中: {' > '.join(path)}  [{current.Name}]")
    return current


def select_and_click(tree, path):
    """选中 + 双击触发业务"""
    current = tree
    for depth, target in enumerate(path):
        found = False
        for child in current.GetChildren():
            if child.Name and child.Name.strip() == target.strip():
                current = child
                found = True
                break
        if not found:
            print(f"✗ 第{depth}层未找到: '{target}'")
            return False
        if depth < len(path) - 1:
            try:
                current.Expand()
            except Exception:
                pass
            time.sleep(0.15)

    current.SetFocus()
    time.sleep(0.1)
    # ★ 双击：用 DoubleClick ★
    current.DoubleClick(simulateMove=False)
    time.sleep(0.1)
    print(f"✓ 已双击: {' > '.join(path)}  [{current.Name}]")
    return True


if __name__ == "__main__":
    elements = findwindows.find_elements(title_re=".*钱龙模拟.*")
    if not elements:
        print("未找到钱龙模拟窗口")
        exit(1)

    target_hwnd = elements[0].handle
    print(f"主窗口: {hex(target_hwnd)}")

    win32gui.SetForegroundWindow(target_hwnd)
    time.sleep(0.3)

    tree_hwnd = find_treeview(target_hwnd, control_id=1223)
    if not tree_hwnd:
        exit(1)

    tree = get_tree(tree_hwnd)

    # 遍历
    print("\n" + "=" * 50)
    print("树结构:")
    print("=" * 50)
    dump_tree(tree)

    # 选择
    print("\n" + "=" * 50)
    print("选择节点:")
    print("=" * 50)
    select_node(tree, ["查询", "资金查询"])

    # 双击（如果单击选中不触发业务，用这个）
    # select_and_click(tree, ["单向下单", "买入平仓"])

    print("\n完成")