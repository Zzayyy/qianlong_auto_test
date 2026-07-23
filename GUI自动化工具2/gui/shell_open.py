# -*- coding: utf-8 -*-
"""在独立 Windows 进程中安全打开文件或目录。"""

from __future__ import annotations

import os
from pathlib import Path
import subprocess


def open_path(path):
    """打开路径，避免在 Tk/Python 主进程中直接调用 ``os.startfile``。

    文件关联操作由独立的 rundll32 进程完成；即使关联程序或 Shell 扩展
    发生原生崩溃，也不会破坏 GUI 主进程的 Python 线程状态。
    """
    target = Path(path).expanduser().resolve()
    if not target.exists():
        raise FileNotFoundError(f"文件或目录不存在: {target}")

    if target.is_dir():
        command = ["explorer.exe", str(target)]
    else:
        command = [
            "rundll32.exe",
            "url.dll,FileProtocolHandler",
            str(target),
        ]

    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    return subprocess.Popen(
        command,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        close_fds=True,
        creationflags=creation_flags,
    )
