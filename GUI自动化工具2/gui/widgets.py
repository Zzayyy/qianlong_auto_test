# -*- coding: utf-8 -*-
"""通用 GUI 控件：带颜色分级的日志文本框等"""

import tkinter as tk
from tkinter import scrolledtext
from datetime import datetime


class ColoredLogText(scrolledtext.ScrolledText):
    """带颜色分级的日志文本框（成功/错误/警告/信息/分隔/高亮）"""

    def __init__(self, master=None, **kwargs):
        super().__init__(master, **kwargs)
        self._setup_tags()

    def _setup_tags(self):
        """配置日志颜色标签"""
        self.tag_configure("success", foreground="#4ec9b0")
        self.tag_configure("error", foreground="#f44747")
        self.tag_configure("warning", foreground="#dcdcaa")
        self.tag_configure("info", foreground="#569cd6")
        self.tag_configure("separator", foreground="#808080")
        self.tag_configure("highlight", foreground="#ce9178")

    @staticmethod
    def _get_tag(message):
        """根据日志内容动态返回颜色标签"""
        msg_lower = message.lower()
        msg_stripped = message.strip()

        # 成功类（绿色）
        if any(kw in msg_lower for kw in ["成功", "完成", "success", "[ok]", "已找到", "已切换", "已点击", "已设置", "已选择", "主动打开"]):
            return "success"
        # 错误类（红色）
        elif any(kw in msg_lower for kw in ["错误", "失败", "error", "异常", "exception", "未找到", "找不到", "超时"]):
            return "error"
        # 警告类（黄色）
        elif any(kw in msg_lower for kw in ["警告", "warn", "提示", "注意", "[停止]", "请", "倒计时", "秒后"]):
            return "warning"
        # 配置信息类（蓝色）
        elif any(kw in msg_lower for kw in ["开始执行", "切换分类", "导出格式", "自动打开", "路径", "excel文件", "委托数量", "目标:", "窗口", "输出路径", "脚本路径", "txt路径", "xls路径"]):
            return "info"
        # 分隔线（灰色）
        elif msg_stripped.startswith("=") or msg_stripped.startswith("-") * 10:
            return "separator"
        # 高亮数据（橙色）
        elif any(kw in msg_lower for kw in ["[预览]", "字段", "共", "行"]):
            return "highlight"
        return None

    def append(self, message):
        """向日志追加一行（须在 UI 主线程调用）"""
        ts = datetime.now().strftime("%H:%M:%S")
        tag = self._get_tag(message)
        self.config(state=tk.NORMAL)
        if tag:
            self.insert(tk.END, f"[{ts}] {message}\n", tag)
        else:
            self.insert(tk.END, f"[{ts}] {message}\n")
        self.see(tk.END)
        self.config(state=tk.DISABLED)
