# -*- coding: utf-8 -*-
"""Windows 回收站操作，用于安全删除报告批次。"""

from __future__ import annotations

import json
import os
from pathlib import Path
import subprocess


def validate_batch_directory(batch_dir, output_dir, run_id):
    """确认目标是输出目录 ``批次`` 下与批次号同名的直接子目录。"""
    if not batch_dir or not output_dir or not run_id:
        raise ValueError("批次信息不完整，不能删除")

    batch_root = (Path(output_dir).expanduser().resolve() / "批次").resolve()
    target = Path(batch_dir).expanduser().resolve()
    if not target.is_dir():
        raise FileNotFoundError(f"批次目录不存在: {target}")
    if target.parent != batch_root:
        raise ValueError(f"目标不在当前交易系统设置批次目录内: {target}")
    if target.name != str(run_id):
        raise ValueError(
            f"目录名与批次号不一致: {target.name} != {run_id}"
        )
    summary_path = target / "批次汇总.json"
    if not summary_path.is_file():
        raise ValueError(f"目标目录缺少批次汇总文件: {summary_path}")
    try:
        payload = json.loads(summary_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError) as exc:
        raise ValueError(f"批次汇总文件无法读取: {exc}") from exc
    if str(payload.get("run_id", "")) != str(run_id):
        raise ValueError("批次汇总文件中的批次号不一致")
    return target


def move_directory_to_recycle_bin(path):
    """通过独立 PowerShell/.NET 进程把目录移入 Windows 回收站。"""
    target = Path(path).expanduser().resolve()
    if not target.is_dir():
        raise FileNotFoundError(f"批次目录不存在: {target}")

    script = (
        "Add-Type -AssemblyName Microsoft.VisualBasic; "
        "$target = [Environment]::GetEnvironmentVariable("
        "'GUI_REPORT_BATCH_TO_RECYCLE'); "
        "[Microsoft.VisualBasic.FileIO.FileSystem]::DeleteDirectory("
        "$target, "
        "[Microsoft.VisualBasic.FileIO.UIOption]::OnlyErrorDialogs, "
        "[Microsoft.VisualBasic.FileIO.RecycleOption]::SendToRecycleBin)"
    )
    creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    child_env = os.environ.copy()
    child_env["GUI_REPORT_BATCH_TO_RECYCLE"] = str(target)
    completed = subprocess.run(
        [
            "powershell.exe",
            "-NoLogo",
            "-NoProfile",
            "-NonInteractive",
            "-Command",
            script,
        ],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        errors="replace",
        timeout=120,
        creationflags=creation_flags,
        env=child_env,
        check=False,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout or "").strip()
        raise OSError(detail or f"移入回收站失败，退出码 {completed.returncode}")
    if target.exists():
        raise OSError("回收站操作返回成功，但批次目录仍然存在")
