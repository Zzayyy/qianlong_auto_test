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
        self._setup_search()
        self._auto_scroll = True  # 是否自动滚动到底部（搜索/查看历史时可暂停）

    def _setup_search(self):
        """搜索高亮标签与搜索状态"""
        self.tag_configure("search_all", background="#fff3bf", foreground="#000000")
        self.tag_configure("search_cur", background="#ffd43b", foreground="#000000")
        self._search_pattern = ""
        self._search_matches = []  # [(start, end), ...]，全部匹配位置
        self._search_idx = -1      # 当前匹配序号（-1 表示未定位）
        self._search_stale = False  # 日志被截断清理后旧索引失效，需重新收集

    # ====================== 自动滚动 ======================
    def set_auto_scroll(self, on, jump_to_end=True):
        """开关自动滚动到底部；on 且 jump_to_end 时立即滚到底。

        jump_to_end=False 仅恢复跟随标志、保持当前视口位置
        （用于「删除空搜索文字」这类只结束搜索、不打扰回看历史的场景）。
        """
        self._auto_scroll = bool(on)
        if on and jump_to_end:
            self.see(tk.END)

    # ====================== 搜索 ======================
    def search_find(self, pattern, forward=True):
        """按关键字搜索并定位，返回 (total, idx)。

        - 搜索词变化时重新统计全部匹配，并从当前视口内第一个匹配开始定位；
        - 同一搜索词连续调用时在上一次匹配基础上向前/向后移动；
        - 无匹配或空关键字返回 (0, -1)。
        """
        pattern = pattern or ""
        if pattern != self._search_pattern or self._search_stale:
            self._search_pattern = pattern
            self._search_matches = self._collect_matches(pattern) if pattern else []
            self._search_idx = -1
            self._search_stale = False
            if self._search_matches:
                self._search_idx = self._nearest_index_after_view_top()
            self._redraw_search_marks()
        if not self._search_matches:
            return 0, -1
        total = len(self._search_matches)
        if forward:
            self._search_idx = (self._search_idx + 1) % total
        else:
            self._search_idx = (self._search_idx - 1) % total
        self._locate_current()
        return total, self._search_idx + 1

    def _collect_matches(self, pattern):
        """收集全部匹配位置（大小写不敏感）"""
        matches = []
        start = "1.0"
        count_var = tk.IntVar()
        while True:
            pos = self.search(
                pattern, start, stopindex=tk.END, count=count_var, nocase=True
            )
            if not pos:
                break
            end = self.index(f"{pos}+{count_var.get()}c")
            matches.append((pos, end))
            start = end
            if len(matches) > 20000:  # 防御性上限，避免极端情况下卡死
                break
        return matches

    def _nearest_index_after_view_top(self):
        """取视口顶部之后（含）的第一个匹配序号，否则从第一个开始"""
        if not self._search_matches:
            return -1
        try:
            first_visible = self.index("@0,0")
        except Exception:
            return 0
        for i, (start, _end) in enumerate(self._search_matches):
            if self.compare(start, ">=", first_visible):
                return i
        return 0

    def _redraw_search_marks(self):
        """重画全部匹配高亮与当前匹配高亮"""
        self.tag_remove("search_all", "1.0", tk.END)
        self.tag_remove("search_cur", "1.0", tk.END)
        for start, end in self._search_matches:
            self.tag_add("search_all", start, end)
        if 0 <= self._search_idx < len(self._search_matches):
            start, end = self._search_matches[self._search_idx]
            self.tag_add("search_cur", start, end)

    def _locate_current(self):
        """滚动到当前匹配并置光标"""
        self._redraw_search_marks()
        if 0 <= self._search_idx < len(self._search_matches):
            start, _end = self._search_matches[self._search_idx]
            self.see(start)
            self.mark_set(tk.INSERT, start)

    def clear_search(self):
        """清除搜索状态与全部高亮"""
        self._search_pattern = ""
        self._search_matches = []
        self._search_idx = -1
        self._search_stale = False
        self.tag_remove("search_all", "1.0", tk.END)
        self.tag_remove("search_cur", "1.0", tk.END)

    def _setup_tags(self):
        """配置日志颜色标签"""
        self.tag_configure("success", foreground="#4ec9b0")
        self.tag_configure("error", foreground="#f44747")
        self.tag_configure("warning", foreground="#dcdcaa")
        self.tag_configure("info", foreground="#569cd6")
        self.tag_configure("separator", foreground="#808080")
        self.tag_configure("highlight", foreground="#ce9178")
        # 交易系统设置的状态标签“执行失败”(core.settings.STATUS_EXECUTION_FAILED)
        # 是状态而非运行异常，用中性金色区分于 error 红。
        self.tag_configure("status_fail", foreground="#d7ba7d")

    @staticmethod
    def _get_tag(message):
        """根据日志内容动态返回颜色标签"""
        msg_lower = message.lower()
        msg_stripped = message.strip()

        # 交易系统设置的状态“执行失败”是状态标签而非运行异常，单独用中性金色，
        # 避免被下方“失败”关键字误判为红色错误。
        if "执行失败" in message:
            return "status_fail"

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
        # 限制日志行数，防止内存泄漏
        MAX_LINES = 5000
        if float(self.index('end-1c')) > MAX_LINES * 2:
            self.delete(1.0, f'end-{MAX_LINES}l')
            # 旧索引因文本截断而失效，下次搜索时重新收集
            self._search_stale = True
        if tag:
            self.insert(tk.END, f"[{ts}] {message}\n", tag)
        else:
            self.insert(tk.END, f"[{ts}] {message}\n")
        if self._auto_scroll:
            self.see(tk.END)
        self.config(state=tk.DISABLED)
