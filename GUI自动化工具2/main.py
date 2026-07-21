# -*- coding: utf-8 -*-
"""
钱龙期权交易 - GUI自动化工具 v2 — 启动入口
=========================================
职责:
  - PyInstaller 打包环境检测与「脚本运行器模式」引导
  - 创建主窗口并启动 Tk 事件循环

运行环境:
    pip install pywinauto openpyxl pandas

使用方法:
    1. 打开钱龙旗舰版，登录交易账号
    2. 运行本GUI工具
    3. 通过菜单选择功能分类，选择脚本，配置参数后执行
"""

import sys
import os
import ctypes
import subprocess

import tkinter as tk

from config import IS_FROZEN, BASE_DIR, get_client, load_user_config
from gui.main_window import AutomationGUI


def _is_admin() -> bool:
    """Whether this process can cross the integrity boundary of the client."""
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def _relaunch_as_admin_if_required() -> bool:
    """Relaunch GUI with UAC consent for clients that run elevated.

    Child script-runner processes inherit the GUI token, so elevation is done
    once at GUI startup instead of once per task.
    """
    user_config = load_user_config()
    client = get_client(user_config.get("client")) or {}
    if not client.get("requires_elevation") or _is_admin():
        return False

    if IS_FROZEN:
        executable = sys.executable
        arguments = subprocess.list2cmdline(sys.argv[1:])
    else:
        executable = sys.executable
        arguments = subprocess.list2cmdline(
            [os.path.abspath(__file__), *sys.argv[1:]]
        )

    shell_execute = ctypes.windll.shell32.ShellExecuteW
    shell_execute.restype = ctypes.c_ssize_t
    result = shell_execute(
        None,
        "runas",
        executable,
        arguments,
        os.getcwd(),
        1,  # SW_SHOWNORMAL
    )
    if result <= 32:
        ctypes.windll.user32.MessageBoxW(
            None,
            f"管理员权限启动失败（错误码 {result}）。\n"
            "请右键自动化工具，选择“以管理员身份运行”。",
            "无法启动自动化工具",
            0x00000010,  # MB_ICONERROR
        )
    return True


# ====================== 脚本运行器模式 ======================
# 打包后exe自身充当Python解释器，用于在无Python环境的电脑上执行子脚本
# 用法: main.exe --_run_script <脚本路径>
def _run_script_mode():
    """
    脚本运行器模式：
    当exe收到 --_run_script 参数时，不启动GUI，而是直接执行指定的Python脚本。
    这样子脚本就能复用exe自带的Python环境和所有打包进去的依赖包。
    """
    if '--_run_script' not in sys.argv:
        return

    # 强制设置 stdout/stderr 编码为 UTF-8，避免特殊字符编码错误
    # 同时设置 line_buffering=True，确保 print 输出实时传递到父进程管道
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace', line_buffering=True)
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace', line_buffering=True)

    # 打包环境下，统一添加 pywin32 DLL 搜索路径
    # 这样所有子脚本（包括不导入 core.window 的脚本）都能正常加载 pywinauto
    if IS_FROZEN:
        dll_dir = os.path.join(os.path.dirname(sys.executable), "_internal", "pywin32_system32")
        if os.path.exists(dll_dir):
            os.environ['PATH'] = dll_dir + os.pathsep + os.environ.get('PATH', '')
            try:
                os.add_dll_directory(dll_dir)
            except AttributeError:
                pass

    idx = sys.argv.index('--_run_script')
    if idx + 1 >= len(sys.argv):
        print("错误: --_run_script 后需要指定脚本路径")
        sys.exit(1)

    script_path = sys.argv[idx + 1]

    if not os.path.exists(script_path):
        print(f"错误: 脚本文件不存在: {script_path}")
        sys.exit(1)

    # 设置子脚本环境
    sys.argv = sys.argv[idx + 1:]  # 子脚本看到的 argv[0] 是脚本路径

    # 确保子脚本所在目录在 sys.path 中
    script_dir = os.path.dirname(os.path.abspath(script_path))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)

    # 确保 _internal 目录也在 sys.path 中（打包后依赖包在这里）
    if IS_FROZEN and BASE_DIR not in sys.path:
        sys.path.insert(0, BASE_DIR)

    try:
        import runpy
        runpy.run_path(script_path, run_name='__main__')
    except SystemExit:
        raise
    except Exception as e:
        print(f"执行脚本失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

    sys.exit(0)


def main():
    root = tk.Tk()
    app = AutomationGUI(root)
    root.mainloop()


if __name__ == "__main__":
    # 脚本运行器模式优先：若收到 --_run_script 参数则执行子脚本后直接退出
    _run_script_mode()
    if _relaunch_as_admin_if_required():
        sys.exit(0)
    main()
