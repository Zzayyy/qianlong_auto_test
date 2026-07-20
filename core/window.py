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
from pywinauto import Application, findwindows
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


def _load_menu_locator():
    """懒加载 win32gui菜单定位 模块（中文文件名，用 importlib 导入）。

    该模块提供 detect_target_bitness / find_treeview / RemoteMem / select_node
    等函数，通过 TVM_* 消息直接定位并选中左侧菜单树节点，不依赖屏幕坐标。
    """
    if getattr(_load_menu_locator, "mod", None) is not None:
        return _load_menu_locator.mod

    import importlib.util
    import os

    here = os.path.dirname(os.path.abspath(__file__))
    loc_path = os.path.abspath(os.path.join(here, "..", "win32gui菜单定位.py"))

    if not os.path.exists(loc_path):
        raise RuntimeError(f"未找到 win32gui菜单定位.py: {loc_path}")

    spec = importlib.util.spec_from_file_location("win32gui_menu_loc", loc_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    _load_menu_locator.mod = mod
    return mod


def _load_uia_locator():
    """懒加载 uiautomation菜单定位 模块（中文文件名，用 importlib 导入）。

    该模块基于 uiautomation，通过 ControlFromHandle + Click 直接选中左侧菜单树
    节点，不需要管理员权限，也不需要向目标进程注入内存（RemoteMem）。
    当 win32gui 方案因权限/位数等原因失败时，作为兜底使用。
    """
    if getattr(_load_uia_locator, "mod", None) is not None:
        return _load_uia_locator.mod

    import importlib.util
    import os

    here = os.path.dirname(os.path.abspath(__file__))
    loc_path = os.path.abspath(os.path.join(here, "..", "uiautomation菜单定位.py"))

    if not os.path.exists(loc_path):
        raise RuntimeError(f"未找到 uiautomation菜单定位.py: {loc_path}")

    spec = importlib.util.spec_from_file_location("uiautomation_menu_loc", loc_path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)

    _load_uia_locator.mod = mod
    return mod


def _parse_panel_path(panel_path: str):
    """把面板路径字符串解析成层级列表。

    "\\查询\\资金持仓" -> ["查询", "资金持仓"]
    """
    path = panel_path.replace("/", "\\")
    if path.startswith("\\"):
        path = path[1:]
    return [p for p in path.split("\\") if p]


def switch_panel(win, panel_path: str, use_title: bool = False):
    """切换到指定面板。

    优先使用 win32gui 方案（参考 win32gui菜单定位.py）：通过 TVM_* 消息直接定位
    并选中左侧菜单树节点，不依赖屏幕坐标。该方案在部分电脑上需要管理员权限，
    且即使给了管理员权限仍可能失败（如位数/权限限制）。

    一旦 win32gui 方案抛出异常，自动回退到 uiautomation 方案
    （参考 uiautomation菜单定位.py）：基于 uiautomation 的 ControlFromHandle + Click，
    不需要管理员权限，兼容性更好。

    Args:
        win: 主窗口（pywinauto 对象或含 handle 属性的对象）
        panel_path: 树形面板路径,如 r"\查询\资金持仓" 或 "撤单"
        use_title: 保留参数（该定位方法不依赖 title，仅用于兼容调用方）
    """
    target_hwnd = getattr(win, "handle", win)

    # 先把目标窗口置前，确保消息/UI 自动化能被正常处理
    try:
        win32gui.SetForegroundWindow(target_hwnd)
    except Exception:
        pass
    time.sleep(0.3)

    # 路径字符串 -> 层级列表
    path_list = _parse_panel_path(panel_path)

    # —— 方案一：win32gui（TVM_* 消息 + RemoteMem），失败则兜底 ——
    try:
        _switch_panel_win32(target_hwnd, path_list)
        panel_name = path_list[-1]
        print(f"[OK] 已切换到'{panel_name}'面板 (win32gui)")
        return
    except Exception as e:
        print(f"[WARN] win32gui 切换面板失败（{e}），回退到 uiautomation 方案...")

    # —— 兜底方案：uiautomation（ControlFromHandle + Click），无需管理员权限 ——
    _switch_panel_uia(target_hwnd, path_list)
    panel_name = path_list[-1]
    print(f"[OK] 已切换到'{panel_name}'面板 (uiautomation)")


def _switch_panel_win32(target_hwnd: int, path_list: list):
    """win32gui 方案：通过 TVM_* 消息直接定位并选中 TreeView 节点。"""
    menu = _load_menu_locator()

    # 检测目标进程位数（32/64 位结构体布局不同，必须使用对应布局）
    target_bits = menu.detect_target_bitness(target_hwnd)

    # 通过控件 ID 直接定位 TreeView 句柄（auto_id=1223）
    tree_hwnd = menu.find_treeview(target_hwnd, control_id=1223)
    if not tree_hwnd:
        raise RuntimeError("未找到左侧 TreeView 控件(auto_id=1223)，无法切换面板")

    # 在目标进程内分配内存，按路径逐级定位并选中节点
    mem = menu.RemoteMem(tree_hwnd, target_bits)
    try:
        h_item = menu.select_node(tree_hwnd, path_list, mem)
        if not h_item:
            raise RuntimeError(f"未找到面板路径: {'\\'.join(path_list)}")
    finally:
        mem.close()


def _switch_panel_uia(target_hwnd: int, path_list: list):
    """uiautomation 兜底方案：ControlFromHandle + Click 选中 TreeView 节点。"""
    uia = _load_uia_locator()

    tree_hwnd = uia.find_treeview(target_hwnd, control_id=1223)
    if not tree_hwnd:
        raise RuntimeError("未找到左侧 TreeView 控件(auto_id=1223)，无法切换面板")

    tree = uia.get_tree(tree_hwnd)
    # select_node 内部按路径逐级展开并最终 Click 选中最末节点
    result = uia.select_node(tree, path_list)
    if not result:
        raise RuntimeError(f"未找到面板路径: {'\\'.join(path_list)}")


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


def close_settings_dialog(dlg, title: str = "交易系统设置", keep_open: bool = False,
                          main_hwnd: int = None, timeout: float = 3.0) -> bool:
    """安全关闭（或保留）交易系统设置对话框。

    任务中心顺序执行场景下用于解决“窗口残留”问题：
      - keep_open=True  -> 保留窗口，供下一个同样属于“交易系统设置”分类的脚本复用。
                           这些脚本的 open_settings_dialog 会先复用已打开的窗口，
                           从而避免每个脚本都重复点击打开/关闭窗口。
      - keep_open=False -> 关闭窗口，防止下一个非“交易系统设置”分类的脚本因找不到
                           本类窗口而报错（也是单独运行/任务中心最后一个任务的安全默认）。

    关闭策略：记录并校验设置窗口句柄，只向该句柄发送 WM_CLOSE，然后等待句柄真正
    销毁。禁止使用依赖当前焦点的 Alt+F4，避免设置窗口关闭后焦点回到主交易窗口时
    误触发整个客户端的退出确认。
    """
    if keep_open:
        print(f"[INFO] 保留'{title}'窗口，供下一个交易系统设置脚本复用")
        return True

    if dlg is None:
        return True

    try:
        dlg_hwnd = int(dlg.handle)
    except Exception as e:
        print(f"[WARN] 无法取得'{title}'窗口句柄: {e}")
        return False

    if not win32gui.IsWindow(dlg_hwnd):
        print(f"[OK] '{title}'窗口已经关闭")
        return True

    # 最重要的安全保护：识别异常时绝不允许关闭主交易窗口。
    if main_hwnd is not None and dlg_hwnd == int(main_hwnd):
        print(
            f"[错误] 拒绝关闭'{title}'：设置窗口句柄与主窗口相同 "
            f"(hwnd={dlg_hwnd})"
        )
        return False

    try:
        class_name = win32gui.GetClassName(dlg_hwnd)
        window_text = win32gui.GetWindowText(dlg_hwnd)
        print(
            f"[INFO] 准备关闭'{title}': hwnd={dlg_hwnd}, "
            f"class={class_name}, title={window_text!r}"
        )
    except Exception:
        pass

    try:
        # PostMessage 只作用于指定句柄，不依赖前台焦点；即使主窗口已重新获得焦点，
        # 也不会像 Alt+F4 一样误触发整个交易客户端退出。
        win32gui.PostMessage(dlg_hwnd, win32con.WM_CLOSE, 0, 0)
    except Exception as e:
        print(f"[WARN] 向'{title}'发送关闭消息失败: {e}")
        return False

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if not win32gui.IsWindow(dlg_hwnd):
            print(f"[OK] 已关闭'{title}'窗口")
            return True
        time.sleep(0.1)

    print(
        f"[WARN] '{title}'窗口在 {timeout:.1f} 秒内未关闭，已停止自动关闭；"
        "不会发送 Alt+F4，以免误关主交易软件"
    )
    return False
