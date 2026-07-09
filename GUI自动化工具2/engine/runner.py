# -*- coding: utf-8 -*-
"""脚本执行引擎：启动子进程、转发输出、支持停止"""

import os
import sys
import time
import threading
import subprocess

from .task import Task


class ScriptRunner:
    """在子进程中执行 Task，并通过回调把输出/结果回传给调用方。

    回调说明（均由调用方注入）：
      - log_level_getter(): 返回当前日志级别（"详细"/"简洁"），决定是否把子脚本
        stdout 实时回显到 UI。
      - on_log(message): 回显到 UI 日志（详细模式下每个 stdout 行都会触发）。
      - on_debug(message): 写入文件日志（始终触发）。
      - on_finish(return_code, elapsed, task): 子进程结束后触发。
      - on_error(exc, task): 执行异常时触发。
    """

    def __init__(self, is_frozen, project_root,
                 log_level_getter=None, on_log=None, on_debug=None,
                 on_finish=None, on_error=None):
        self.is_frozen = is_frozen
        self.project_root = project_root
        self.log_level_getter = log_level_getter or (lambda: "详细")
        self.on_log = on_log or (lambda *a, **k: None)
        self.on_debug = on_debug or (lambda *a, **k: None)
        self.on_finish = on_finish or (lambda *a, **k: None)
        self.on_error = on_error or (lambda *a, **k: None)

        self._process = None
        self._running = False

    @property
    def is_running(self):
        return self._running

    def run(self, task: Task):
        """启动任务（在后台线程中执行，避免阻塞 UI）"""
        if self._running:
            return
        self._running = True
        threading.Thread(target=self._run_thread, args=(task,), daemon=True).start()

    def _run_thread(self, task: Task):
        try:
            env = task.build_env(self.project_root)
            # 子进程用 UTF-8 输出，避免 gbk 编码错误
            env["PYTHONIOENCODING"] = "utf-8"
            env["PYTHONUTF8"] = "1"

            # 构建命令：打包后用 exe 自身充当 Python 解释器，开发环境用系统 Python
            if self.is_frozen:
                cmd = [sys.executable, "--_run_script", task.path]
            else:
                cmd = [sys.executable, "-u", task.path]

            self.on_log(f"目标: {cmd[0]}")
            self.on_debug(f"执行命令: {cmd}")

            self._process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW,
                env=env,
            )

            start = time.time()
            for line in self._process.stdout:
                line = line.rstrip()
                if line:
                    # 简洁模式下不显示子脚本自身的 print 输出，但继续读取管道避免子进程阻塞
                    self.on_debug(line)
                    if self.log_level_getter() == "详细":
                        self.on_log(line)

            return_code = self._process.wait()
            elapsed = time.time() - start
            self.on_finish(return_code, elapsed, task)

        except Exception as e:
            self.on_error(e, task)

        finally:
            self._running = False
            self._process = None

    def stop(self):
        """停止当前正在运行的子进程"""
        if self._process and self._process.poll() is None:
            try:
                self._process.terminate()
                self._process.wait(timeout=3)
            except Exception:
                try:
                    self._process.kill()
                except Exception:
                    pass
