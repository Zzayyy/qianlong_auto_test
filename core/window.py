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
from core.clients import get_client, get_clients
from core.native_tree import (
    NativeTreeAccessError,
    select_tree_path,
    select_tree_path_by_position,
)
from pywinauto import Application
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

    candidates = []

    def _enum_top_level(hwnd, _):
        try:
            title = win32gui.GetWindowText(hwnd) or ""
            if keyword not in title or "GUI自动化工具" in title:
                return
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if pid in exclude_pids:
                return
            class_name = win32gui.GetClassName(hwnd)
            rect = win32gui.GetWindowRect(hwnd)
            has_area = rect[2] > rect[0] and rect[3] > rect[1]
            # 国泰海通/钱龙主窗口类优先，其次才按可见性和面积排序。
            rank = (
                0 if class_name == "QL_OPTION_MAINWND_CLASS" else 1,
                0 if win32gui.IsWindowVisible(hwnd) else 1,
                0 if has_area else 1,
            )
            candidates.append((rank, hwnd, pid, class_name, title))
        except Exception:
            return

    win32gui.EnumWindows(_enum_top_level, None)
    if candidates:
        candidates.sort(key=lambda row: row[0])
        _, hwnd, pid, class_name, title = candidates[0]
        print(
            f"[INFO] 目标窗口: hwnd={hwnd}, pid={pid}, "
            f"class={class_name}, title={title!r}"
        )
        return hwnd

    raise RuntimeError(f"未找到包含'{keyword}'的窗口,请确认软件已启动")


def activate_window(hwnd: int):
    """连接窗口并置前。"""
    Timings.fast()
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.3)
    app = Application(backend="uia").connect(handle=hwnd)
    win = app.window(handle=hwnd)
    win.set_focus()
    return win


def _resolve_client_for_window(win):
    """Resolve the client profile from GUI selection or the live window title.

    GUI-launched tasks always provide ``GUI_CLIENT_ID``.  Directly launched
    legacy scripts do not, so title inference keeps their original usage while
    still selecting the correct Win11 positional fingerprint.
    """
    client_id = os.environ.get("GUI_CLIENT_ID")
    client = get_client(client_id) if client_id else None
    if client is not None:
        return client

    try:
        title = win32gui.GetWindowText(int(win.handle)) or ""
    except Exception:
        title = ""
    matches = [
        candidate for candidate in get_clients()
        if candidate.get("window_key")
        and candidate["window_key"] in title
    ]
    if len(matches) == 1:
        print(
            f"[INFO] 根据目标窗口识别客户端: "
            f"{matches[0].get('name', matches[0].get('id'))}"
        )
        return matches[0]
    return None


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
    panel_name = panel_path.replace("/", "\\").rsplit("\\", 1)[-1]
    native_error = None
    try:
        info = select_tree_path(win.handle, panel_path, control_id=1223)
        print(
            f"[OK] 已切换到'{panel_name}'面板（原生 Win32 TreeView，"
            f"节点={info['node_count']}，目标={info['target_bits']}位）"
        )
        return
    except NativeTreeAccessError as e:
        native_error = e
        print(f"[INFO] 客户端不允许读取菜单文字: {e}")
        print("[INFO] 改用经过节点总数校验的原生位置定位...")
    except Exception as e:
        native_error = e
        print(f"[WARN] 原生 TreeView 定位失败: {e}")

    client = _resolve_client_for_window(win)
    position_profile = (client or {}).get("native_tree_profile")
    position_error = None
    if position_profile:
        for attempt in range(3):
            try:
                info = select_tree_path_by_position(
                    win.handle, panel_path, position_profile, control_id=1223
                )
                print(
                    f"[OK] 已切换到'{panel_name}'面板（原生位置指纹，"
                    f"节点={info['node_count']}，位置={info['position']}）"
                )
                return
            except Exception as error:
                position_error = error
                if attempt < 2:
                    print(
                        f"[INFO] 原生位置定位暂时失败，正在重试"
                        f"（{attempt + 2}/3）: {error}"
                    )
                    time.sleep(0.6)
        print(f"[WARN] 原生位置定位失败: {position_error}")
    else:
        position_error = "当前客户端未配置原生位置指纹"

    print("[INFO] 尝试 UI Automation 兼容路径...")

    try:
        tree = win.child_window(auto_id="1223", control_type="Tree")
        tree.wait("ready", timeout=5)
        tree.set_focus()
        item = _select_tree_item_by_path(tree, panel_path)
        print(f"[OK] 已切换到'{panel_name}'面板（UI Automation）")
        return
    except Exception as uia_error:
        raise RuntimeError(
            f"无法切换到菜单 {panel_path!r}。"
            f"原生文本定位失败: {native_error}；"
            f"原生位置定位失败: {position_error}；"
            f"UIA 定位失败: {uia_error}"
        ) from uia_error



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
