import os
import sys
import win32gui
import win32con
import win32process
from win32gui import (
    EnumChildWindows,
    GetWindowText,
    IsWindowVisible
)
from core.clients import get_client
from pywinauto import Application,findwindows
import time
# -*- coding: utf-8 -*-
"""窗口操作公共模块：查找、激活、倒计时、面板切换"""

# 必须在 import win32api/pywinauto 之前执行！
if getattr(sys, 'frozen', False):
    # 获取 exe 所在的根目录（即 dist 下的 "钱龙期权交易自动化工具" 文件夹）
    base_dir = os.path.dirname(sys.executable)
    
    # 拼接出 DLL 所在的完整路径
    dll_dir = os.path.join(base_dir, "_internal", "pywin32_system32")
    
    # 检查文件夹是否存在
    if os.path.exists(dll_dir):
        # 方法1：添加到系统 PATH（传统方法，兼容性好）
        os.environ['PATH'] = dll_dir + os.pathsep + os.environ.get('PATH', '')
        
        # 方法2（Python 3.8+ 官方推荐）：注册为 DLL 搜索目录
        try:
            os.add_dll_directory(dll_dir)
            print(f"✓ 成功添加 DLL 搜索路径: {dll_dir}")
        except AttributeError:
            pass  # 低版本 Python 忽略
    else:
        print(f"⚠ 警告：未找到 DLL 文件夹，路径为 {dll_dir}")

# --- 现在可以正常导入 pywinauto 了 ---


import time
import sys
import ctypes
from pywinauto import Application, findwindows, mouse
from pywinauto.timings import Timings



def find_window(keyword: str) -> int:
    """根据关键字查找窗口,返回第一个匹配且非本工具的句柄。

    本自动化工具的窗口标题同样包含客户端关键字（如“钱龙模拟期权宝 - GUI自动化工具”），
    若不排除，find_elements 可能把自身窗口当成目标，导致切错软件。
    排除策略：
      1) 优先使用当前客户端档案中的 window_key（多客户端支持，GUI_CLIENT_ID 指定时）；
      2) 排除 GUI 进程及其子进程（通过 GUI_PID 环境变量 + 当前子进程 PID）；
      3) 排除标题中含本工具标记“GUI自动化工具”的窗口（兜底）。
    """
    # 多客户端支持：若 GUI 指定了客户端，优先使用该客户端的 window_key
    client_id = os.environ.get("GUI_CLIENT_ID")
    if client_id:
        client = get_client(client_id)
        if client and client.get("window_key"):
            keyword = client["window_key"]

    # 收集需要排除的进程 PID
    exclude_pids = {os.getpid()}  # 当前子进程自身
    gui_pid = os.environ.get("GUI_PID")
    if gui_pid and gui_pid.isdigit():
        exclude_pids.add(int(gui_pid))

    elements = findwindows.find_elements(title_re=f".*{keyword}.*")
    for el in elements:
        # 兜底：标题含本工具标记的直接跳过
        try:
            if "GUI自动化工具" in (el.name or ""):
                continue
        except Exception:
            pass
        # 按进程 PID 排除自身
        try:
            _, pid = win32process.GetWindowThreadProcessId(el.handle)
            if pid in exclude_pids:
                continue
        except Exception:
            pass
        return el.handle

    raise RuntimeError(f"未找到包含'{keyword}'的窗口,请确认软件已启动")


def activate_window(hwnd: int):
    """连接窗口并置前。"""
    Timings.fast()
    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    win.set_focus()
    return win


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


def _select_tree_item_by_path(tree, panel_path: str):
    """逐级 select 展开树节点。

    Args:
        tree: Tree 控件
        panel_path: 树形面板路径,如 r"\查询\资金持仓" 或 "\撤单"
    """
    # 规范化路径
    path = panel_path.replace("/", "\\")
    if not path.startswith("\\"):
        path = "\\" + path

    # 解析路径,过滤空字符串
    parts = [p for p in path.split("\\") if p]
    if not parts:
        raise ValueError(f"无效的面板路径: {panel_path}")

    # 逐级 select 展开
    container = tree
    for part in parts:
        item = container.child_window(title=part, control_type="TreeItem", found_index=0)
        tree.set_focus()        # 必须
        item.wait("visible", timeout=3)
        item.select()
        time.sleep(0.15)
        container = item  # 下一级只在当前节点子节点中搜索

    return item


def switch_panel(win, panel_path: str, use_title: bool = False):
    """切换到指定面板。

    Args:
        win: 主窗口
        panel_path: 树形面板路径,如 r"\查询\资金持仓" 或 "撤单"
        use_title: 是否用title定位TreeItem(历史委托/历史成交需要)
    """
    # ── 先激活左侧菜单区域（全局菜单Home功能） ──
    rect = win.rectangle()

    # 根据屏幕分辨率动态计算水平偏移量
    screen_width = ctypes.windll.user32.GetSystemMetrics(0)
    offset_x = int(screen_width * 0.01)
    if offset_x < 50:
        offset_x = 50
    if offset_x > 200:
        offset_x = 200

    # 菜单位置：窗口最左边 + 偏移量，垂直居中
    menu_x = rect.left + offset_x 
    menu_y = rect.top + (rect.bottom - rect.top) // 2

    mouse.move(coords=(menu_x, menu_y))
    time.sleep(0.3)
    mouse.click(coords=(menu_x, menu_y))
    time.sleep(0.5)
    win.type_keys("{HOME}", with_spaces=False)
    time.sleep(0.3)
    # ───────────────────────────────────────────────

    tree = win.child_window(auto_id="1223", control_type="Tree")
    tree.wait("ready", timeout=10)
    tree.set_focus()

    # 先滚到顶部
    tree.type_keys("{HOME}", with_spaces=False)
    time.sleep(0.2)

    # 右击树控件，选择"全部展开"
    tree.set_focus()
    tree.click_input(button="right")
    time.sleep(0.5)
    try:
        # 用键盘方向键导航到"全部展开"
        # 菜单顺序：打开 → 刷新数据 → 在线帮助 → (分隔线) → 全部展开
        win.type_keys("{DOWN}{DOWN}{DOWN}{DOWN}{ENTER}", with_spaces=False, pause=0)
        print("[OK] 已右击选择'全部展开'")
        time.sleep(0.5)
    except Exception as e:
        print(f"[WARN] 右击展开菜单操作失败，跳过: {e}")
    
    # 获取控件的矩形区域
    rect = tree.rectangle()
    print(f"控件位置: left={rect.left}, top={rect.top}, right={rect.right}, bottom={rect.bottom}")

    # 提取最终面板名称
    panel_name = panel_path.rsplit("\\", 1)[-1]

    # 在控件中心位置滚动
    center_x = (rect.left + rect.right) // 2
    center_y = (rect.top + rect.bottom) // 2

    # ── 优化：第一次向下滚动之前，先尝试查找一次 ──
    # 全部展开后目标节点可能已在可见区域内，可直接命中，避免无谓滚动。
    try:
        item = _select_tree_item_by_path(tree, panel_path)
        print(f"[OK] 已切换到'{panel_name}'面板（首次查找命中，无需滚动）")
        return
    except Exception as e:
        print(f"[INFO] 首次查找未命中，开始向下滚动查找 '{panel_name}'...")

    # 第一次向下滚动
    mouse.scroll(coords=(center_x, center_y), wheel_dist=-7)
    time.sleep(0.3)

    # 第一次滚动后查找
    try:
        item = _select_tree_item_by_path(tree, panel_path)
    except Exception as e:
        # 如果第一次滚动后找不到，再滚动一次
        print(f"[WARN] 第一次滚动后未找到目标节点，再滚动一次...")
        mouse.scroll(coords=(center_x, center_y), wheel_dist=-7)
        time.sleep(0.3)
        try:
            item = _select_tree_item_by_path(tree, panel_path)
        except Exception as e2:
            # 如果第二次滚动后还找不到，直接查找最终面板名称
            print(f"[WARN] 第二次滚动后仍未找到，尝试直接查找面板: {panel_name}")
            item = tree.child_window(title=panel_name, control_type="TreeItem")
            item.wait("visible", timeout=5)
            item.select()
            time.sleep(0.15)

    # 最后 click_input 确保触发点击事件
    # item.click_input()
    print(f"[OK] 已切换到'{panel_name}'面板")



def click_output_button(win, button_auto_id: str = "1159") -> bool:
    """用 win32gui 查找并点击"输出"按钮（参考 win32gui找按钮后点击.py）。"""
    try:
        target_hwnd = win.handle
        buttons = []

        def _enum_cb(hwnd, _):
            if GetWindowText(hwnd) == "输出":
                buttons.append(hwnd)

        EnumChildWindows(target_hwnd, _enum_cb, None)

        if not buttons:
            print("[WARN] 未找到文字为'输出'的按钮")
            return False

        # 取第一个可见的按钮
        btn = next((h for h in buttons if IsWindowVisible(h)), None)
        if not btn:
            print("[WARN] 找到'输出'按钮但均不可见")
            return False

        # 发送 BM_CLICK 消息点击按钮（0x00F5 = BM_CLICK）
        print(btn)
        def click_hwnd(hwnd):
            import win32api
            import win32con
            from win32gui import GetWindowRect

            l, t, r, b = GetWindowRect(hwnd)

            x = (l + r) // 2
            y = (t + b) // 2

            win32api.SetCursorPos((x, y))
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
            win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
        click_hwnd(btn)
        print(f"[OK] 已用 win32gui 点击'输出'按钮(hwnd={btn})")
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"[WARN] 点击'输出'按钮失败: {e}")
        return False


def close_settings_dialog(dlg, title: str = "交易系统设置", keep_open: bool = False):
    """关闭（或保留）交易系统设置对话框。

    任务中心顺序执行场景下用于解决“窗口残留”问题：
      - keep_open=True  -> 保留窗口，供下一个同样属于“交易系统设置”分类的脚本复用。
                           这些脚本的 open_settings_dialog 会先复用已打开的窗口，
                           从而避免每个脚本都重复点击打开/关闭窗口。
      - keep_open=False -> 关闭窗口，防止下一个非“交易系统设置”分类的脚本因找不到
                           本类窗口而报错（也是单独运行/任务中心最后一个任务的安全默认）。

    关闭策略：优先 dlg.close()（WM_CLOSE），失败则用 Alt+F4 兜底；最后校验窗口是否已关闭。
    """
    if keep_open:
        print(f"[INFO] 保留'{title}'窗口，供下一个交易系统设置脚本复用")
        return

    if dlg is None:
        return

    try:
        # 策略1：pywinauto close()（WM_CLOSE）
        try:
            dlg.close()
        except Exception:
            pass

        # 校验：若窗口仍存在，用 Alt+F4 再尝试一次
        try:
            dlg.rectangle()
        except Exception:
            print(f"[OK] 已关闭'{title}'窗口")
            return

        try:
            dlg.set_focus()
            dlg.type_keys("%{F4}")
            time.sleep(0.5)
        except Exception:
            pass

        # 再校验一次是否已关闭
        try:
            dlg.rectangle()
            print(f"[WARN] 未能自动关闭'{title}'窗口，请手动关闭")
        except Exception:
            print(f"[OK] 已关闭'{title}'窗口")
    except Exception as e:
        print(f"[WARN] 关闭'{title}'窗口时出错: {e}")