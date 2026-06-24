# -*- coding: utf-8 -*-
"""
钱龙期权交易 - 全选撤单自动化
=====================================
功能流程:
    1. 查找"钱龙"主窗口
    2. 激活窗口并切换到"撤单"面板(左侧树节点)
    3. 点击"全选"按钮,选中当前列表里所有委托
    4. 点击"撤单"按钮,触发批量撤单
    5. 等待置顶弹窗出现,按回车确认

组件来源:
    撤单菜单: 期权下单组件.txt L236-238  (TreeItem, title="撤单")
    撤单按钮: 撤单组件.txt   L387-388  (Button, auto_id="1161", title="撤单")
    全选按钮: 撤单组件.txt   L391-392  (Button, auto_id="1165", title="全选")

运行环境:
    pip install pywinauto==0.6.8

使用方法:
    1. 打开钱龙旗舰版,登录交易账号
    2. 运行本脚本(将鼠标快速移到屏幕左上角可紧急停止)
    3. 倒计时内切换到钱龙窗口
"""

from pywinauto import Application, findwindows
import time
import sys


# ====================== 可配置参数 ======================
WINDOW_KEY     = "钱龙模拟期权宝"        # 窗口标题关键字
TREE_ITEM      = "撤单"        # 左侧树节点名称
SELECT_ALL_AID = "1165"        # "全选" 按钮 auto_id
CANCEL_BTN_AID = "1161"        # "撤单" 按钮 auto_id
AUTO_CONFIRM   = True         # 是否自动确认置顶弹窗
EXPECTED_DIALOGS = 1          # 一般 1 个确认弹窗(撤单确认或结果)
COUNTDOWN      = 3            # 操作前倒计时秒数
INTERVAL       = 0.5          # "全选"与"撤单"之间的间隔(秒)
# ========================================================


def find_window(keyword: str) -> int:
    """根据关键字查找窗口,返回第一个匹配的句柄。"""
    elements = findwindows.find_elements(title_re=f".*{keyword}.*")
    if not elements:
        raise RuntimeError(f"未找到包含'{keyword}'的窗口,请确认软件已启动")
    return elements[0].handle


def activate_window(hwnd: int):
    """连接窗口并置前。"""
    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    win.set_focus()
    return win


def switch_panel(win, tree_item: str):
    """点击左侧树节点,切换到指定面板。"""
    tree = win.child_window(auto_id="1223", control_type="Tree")
    tree.wait("ready", timeout=10)
    tree.set_focus()

    # 先滚到顶部
    tree.type_keys("{HOME}", with_spaces=False)
    time.sleep(0.2)

    # 直接查找并 select
    item = tree.child_window(title=tree_item, control_type="TreeItem")
    item.wait("visible", timeout=10)
    item.select()
    time.sleep(0.15)
    # item.click_input()
    print(f"[OK] 已切换到面板: {tree_item}")

def click_button(win, auto_id: str, title: str):
    """按 auto_id 点击面板内的按钮(避开 title 漂移问题)。"""
    btn = win.child_window(auto_id=auto_id, control_type="Button")
    btn.wait("ready", timeout=5)
    btn.click_input()
    print(f"[OK] 已点击按钮: {title} (auto_id={auto_id})")


def press_enter_to_confirm(
    dialog_patterns: list = None,
    timeout: float = 8,
):
    """通过回车键确认弹窗(更稳:不依赖按钮坐标/控件树)。

    钱龙所有确认弹窗的默认按钮都是"确定(Y)",回车即触发。

    Args:
        dialog_patterns: 弹窗标题需要包含的关键字列表(任一匹配即可)
        timeout: 等待秒数
    Returns:
        bool: 是否成功确认
    """
    if dialog_patterns is None:
        # 撤单场景下,可能出现的弹窗标题: "提示"、"撤单"、"期权下单" 等
        dialog_patterns = ["提示", "撤单", "期权下单", "期权撤单", "警告"]

    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if not any(p in title for p in dialog_patterns):
                    continue
                # 弹窗是模态的,先确保焦点在它身上,再按回车
                dlg.set_focus()
                time.sleep(0.1)
                dlg.type_keys("{ENTER}", with_spaces=False)
                print(f"[OK] 回车确认 (hwnd={hwnd}, title='{title}')")
                return True
            except Exception:
                continue
        time.sleep(0.1)
    print(f"[WARN] 等待弹窗超时({timeout}s),匹配: {dialog_patterns}")
    return False


def confirm_all_dialogs(
    max_dialogs: int = 5,
    no_dialog_timeout: float = 2.0,
    per_dialog_timeout: float = 4.0,
):
    """自动确认所有弹窗，直到一段时间内没有新弹窗出现。
    
    Args:
        max_dialogs: 最大弹窗数量上限（防止死循环）
        no_dialog_timeout: 等待新弹窗的超时时间（秒），超过此时间无新弹窗则认为全部处理完毕
        per_dialog_timeout: 单个弹窗的等待超时时间（秒）
    """
    count = 0
    for i in range(1, max_dialogs + 1):
        print(f"[..] 等待第 {i} 个弹窗 (超时{no_dialog_timeout}s无新弹窗则结束)...")
        ok = press_enter_to_confirm(timeout=per_dialog_timeout)
        if ok:
            count += 1
            print(f"[OK] 已确认第 {count} 个弹窗")
            time.sleep(0.4)  # 弹窗间短暂间隔
        else:
            # 超时未出现弹窗，认为全部处理完毕
            print(f"[OK] 无更多弹窗，共确认 {count} 个")
            break
    else:
        print(f"[WARN] 达到最大弹窗数量上限({max_dialogs})")


def countdown(seconds: int):
    """操作前倒计时,让用户有时间切换窗口。"""
    print(f"将在 {seconds} 秒后开始,请把焦点切到钱龙软件...")
    try:
        for i in range(seconds, 0, -1):
            print(f"  {i}...", end="\r")
            time.sleep(1)
    except KeyboardInterrupt:
        raise
    print(" " * 30, end="\r")


def main():
    try:
        print(f"计划: 切换到 '{TREE_ITEM}' 面板 → 全选 → 撤单 → 回车确认 "
              f"{EXPECTED_DIALOGS} 个弹窗")

        countdown(COUNTDOWN)

        hwnd = find_window(WINDOW_KEY)
        print(f"[OK] 已找到窗口,句柄 = {hwnd}")

        win = activate_window(hwnd)
        switch_panel(win, TREE_ITEM)
        time.sleep(0.5)                 # 等待面板内容加载

        # 1) 全选
        click_button(win, SELECT_ALL_AID, "全选")
        time.sleep(INTERVAL)

        # 2) 撤单
        click_button(win, CANCEL_BTN_AID, "撤单")

        # 3) 自动确认所有弹窗(无论1个还是多个)
        if AUTO_CONFIRM:
            confirm_all_dialogs()

        print("\n=== 全选撤单操作完成 ===")

    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
