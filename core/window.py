import os
import sys
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
    """根据关键字查找窗口,返回第一个匹配的句柄。"""
    elements = findwindows.find_elements(title_re=f".*{keyword}.*")
    if not elements:
        raise RuntimeError(f"未找到包含'{keyword}'的窗口,请确认软件已启动")
    return elements[0].handle


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


def switch_panel(win, panel_path: str, use_title: bool = False):
    """切换到指定面板。

    Args:
        win: 主窗口
        panel_path: 树形面板路径,如 r"\查询\资金持仓"
        use_title: 是否用title定位TreeItem(历史委托/历史成交需要)
    """
    if use_title:
        # 历史委托/历史成交使用 title 定位
        panel_name = panel_path.rsplit("\\", 1)[-1]
        tree = win.child_window(title=panel_name, control_type="TreeItem")
        tree.set_focus()
        tree.type_keys("{HOME}", with_spaces=False)
        tree.wait("exists", timeout=10)
        item = tree.get_item(panel_path)
        item.select()
        item.click_input()
    else:
        # 标准方式：通过 auto_id="1223" 定位 Tree
        tree = win.child_window(auto_id="1223", control_type="Tree")
        tree.wait("ready", timeout=10)
        item = tree.get_item(panel_path)
        item.select()

    print(f"[OK] 已切换到'{panel_path.rsplit(chr(92), 1)[-1]}'面板")


def click_output_button(win, button_auto_id: str = "1159") -> bool:
    """点击"输出"按钮。"""
    try:
        output_btn = win.child_window(auto_id=button_auto_id, control_type="Button")
        output_btn.wait("ready", timeout=5)
        output_btn.click_input()
        print(f"[OK] 已点击'输出'按钮(auto_id={button_auto_id})")
        time.sleep(0.5)
        return True
    except Exception as e:
        print(f"[WARN] 点击'输出'按钮失败: {e}")
        return False
