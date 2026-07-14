# -*- coding: utf-8 -*-
"""
钱龙拆分申报 - 全自动轮询脚本
===========================
功能:
    遍历所有交易所(上证/深证)
    - 无数据的交易所自动跳过
    - 有数据的交易所执行第一个拆分申报

使用方法:
    1. 打开钱龙旗舰版,登录交易账号
    2. 修改下方参数(默认使用全部交易所)
    3. 运行本脚本(将鼠标快速移到屏幕左上角可紧急停止)
"""

from pywinauto import Application, findwindows
from core.window import switch_panel
import time
import sys
import os


# ====================== 可配置参数 ======================
WINDOW_KEY = "钱龙模拟期权宝"        # 窗口标题关键字
ORDER_QTY = int(os.environ.get("GUI_ORDER_QTY", "1"))  # 委托数量(拆分申报弹窗中的数量)，可由 GUI 参数配置覆盖
COUNTDOWN = 3              # 操作前倒计时秒数
# ========================================================


# 所有交易所
EXCHANGES = ["上证", "深证"]


def find_window(keyword: str):
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


def select_exchange(win, exchange: str):
    """选择交易所(上证/深证)。"""
    combo = win.child_window(auto_id="9059", control_type="ComboBox")
    combo.wait("ready", timeout=5)

    try:
        combo.select(exchange)
        print(f"[OK] 已选择交易所: {exchange}")
        return True
    except Exception:
        pass

    try:
        combo.click_input()
        time.sleep(0.4)
        combo.type_keys(exchange, with_spaces=False)
        time.sleep(0.2)
        combo.type_keys("{ENTER}", with_spaces=False)
        print(f"[OK] 交易所(键盘过滤): {exchange}")
        return True
    except Exception as e:
        print(f"[WARN] 交易所选择失败: {e}")
        return False


def check_list_has_data(win, wait_timeout: float = 3) -> bool:
    """
    检查 ListBox(auto_id="1229") 是否有有效数据行。
    判断规则:
        - 0 行: 无数据
        - 1 行: 占位提示"没有查询到相应的数据",视为无数据
        - 2+ 行: 有数据
    """
    list_box = win.child_window(auto_id="1229", control_type="List")
    try:
        list_box.wait("ready", timeout=10)
    except Exception:
        return False

    def _is_data_row(item) -> bool:
        rect = item.rectangle()
        if rect != (0, 0, 0, 0):
            return True
        try:
            if item.children(control_type="Static"):
                return True
        except Exception:
            pass
        return False

    end = time.time() + wait_timeout
    while time.time() < end:
        rows = list_box.children(control_type="ListItem")
        real_rows = [r for r in rows if _is_data_row(r)]
        if len(real_rows) >= 2:
            return True
        time.sleep(0.3)

    return False


def click_first_list_row_and_split(win, wait_timeout: float = 5) -> bool:
    """
    点击 ListBox 第一条数据并执行拆分申报。
    返回 True 表示执行成功, False 表示无数据或执行失败。
    """
    list_box = win.child_window(auto_id="1229", control_type="List")
    list_box.wait("ready", timeout=10)

    def _is_data_row(item) -> bool:
        rect = item.rectangle()
        if rect != (0, 0, 0, 0):
            return True
        try:
            if item.children(control_type="Static"):
                return True
        except Exception:
            pass
        return False

    end = time.time() + wait_timeout
    real_rows = []
    while time.time() < end:
        rows = list_box.children(control_type="ListItem")
        real_rows = [r for r in rows if _is_data_row(r)]
        if len(real_rows) >= 2:
            break
        time.sleep(0.3)

    total_rows = len(real_rows)
    if total_rows < 2:
        return False

    first_row = real_rows[0]
    print(f"[OK] ListBox 共 {total_rows} 条数据,点击第一条选中")
    first_row.click_input()
    time.sleep(0.3)

    # 点击"拆分"按钮
    split_btn = win.child_window(title="拆分", control_type="Button")
    split_btn.wait("ready", timeout=5)
    split_btn.click_input()
    print("[OK] 已点击'拆分'按钮")

    # 填写委托数量
    fill_qty_in_dialog(timeout=8)
    time.sleep(0.3)

    # 处理后续弹窗(确认弹窗、警告弹窗等)
    find_and_enter_dialog("拆分申报", timeout=2)
    ok = find_and_enter_dialog("警告", timeout=1)
    if not ok:
        find_and_enter_dialog("拆分申报", timeout=1)

    return True


def fill_qty_in_dialog(timeout: float = 8):
    """等待"拆分申报"弹窗出现,填写委托数量。"""
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if "拆分申报" not in title:
                    continue
                dlg.set_focus()
                qty_edit = dlg.child_window(auto_id="9057", control_type="Edit")
                qty_edit.wait("ready", timeout=3)
                qty_edit.double_click_input()
                time.sleep(0.1)
                qty_edit.type_keys("^a", with_spaces=False)
                qty_edit.set_edit_text(str(ORDER_QTY))
                print(f"[OK] 拆分申报弹窗: 委托数量已设为 {ORDER_QTY}")
                return True
            except Exception:
                continue
        time.sleep(0.2)

    print(f"[WARN] 等待'拆分申报'弹窗超时({timeout}s)")
    return False


def find_and_enter_dialog(title_keyword: str, timeout: float = 8) -> bool:
    """
    轮询顶层窗口,找到标题含 title_keyword 的弹窗后回车确认。
    """
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if title_keyword not in title:
                    continue
                dlg.set_focus()
                time.sleep(0.1)
                dlg.type_keys("{ENTER}", with_spaces=False)
                print(f"[OK] '{title_keyword}'弹窗已回车确认 (hwnd={hwnd}, title='{title}')")
                return True
            except Exception:
                continue
        time.sleep(0.15)
    print(f"[WARN] 等待'{title_keyword}'弹窗超时({timeout}s)")
    return False


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


def process_exchange(win, exchange: str) -> bool:
    """
    处理单个交易所:
    1. 检查是否有数据
    2. 有则执行拆分申报
    返回 True 表示执行了拆分, False 表示跳过
    """
    print(f"\n--- 处理: {exchange} ---")

    # 检查是否有数据
    if not check_list_has_data(win, wait_timeout=3):
        print(f"[--] 交易所无数据,跳过")
        return False

    # 有数据,执行拆分
    print(f"[OK] 交易所有数据,执行拆分申报")
    if click_first_list_row_and_split(win, wait_timeout=5):
        print(f"[OK] {exchange} 拆分申报完成")
        return True
    else:
        print(f"[--] 拆分申报执行失败")
        return False


def main():
    try:
        countdown(COUNTDOWN)

        hwnd = find_window(WINDOW_KEY)
        print(f"[OK] 已找到窗口,句柄 = {hwnd}")

        win = activate_window(hwnd)

        # 切换到拆分申报面板
        switch_panel(win, r"\组合申报\拆分申报")
        time.sleep(0.5)

        total_executed = 0
        total_skipped = 0

        # 遍历所有交易所
        for exchange in EXCHANGES:
            print(f"\n{'='*50}")
            print(f"开始处理交易所: {exchange}")
            print(f"{'='*50}")

            # 选择交易所
            if not select_exchange(win, exchange):
                print(f"[错误] 交易所选择失败: {exchange}")
                continue

            time.sleep(0.5)

            try:
                if process_exchange(win, exchange):
                    total_executed += 1
                else:
                    total_skipped += 1
            except Exception as e:
                print(f"[错误] 处理异常: {type(e).__name__}: {e}")
                total_skipped += 1

            # 每个交易所处理完后稍作等待,让界面稳定
            time.sleep(0.5)

        print(f"\n{'='*50}")
        print(f"=== 全部完成 ===")
        print(f"执行拆分申报: {total_executed} 个")
        print(f"跳过(无数据): {total_skipped} 个")
        print(f"{'='*50}")

    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
