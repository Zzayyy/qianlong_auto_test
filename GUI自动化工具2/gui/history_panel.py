# -*- coding: utf-8 -*-
"""任务历史标签页面板
========================
与 TaskCenter / SchedulerPanel / ComparePanel / SettingsReportPanel 同构：
HistoryPanel(parent, controller)，挂在主窗口右侧 Notebook 的「任务历史」标签页。

数据层（HistoryManager）在 gui/history.py，仅负责 记录增删改 + JSON 持久化；
本面板只负责 统计卡片 / 列表视图 / 筛选分页 / 详情与清空 等纯 UI 交互。

对主窗口的依赖：
  - controller.gui：用于调用 _log 与访问 history（主窗口 self.gui = self）
  - controller（主窗口）：history_page 置顶由外部在新增任务时调用 reset_page()
"""
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox

from gui.history import (
    STATUS_SUCCESS,
    STATUS_FAILED,
    STATUS_ERROR,
    STATUS_STOPPED,
    STATUS_RUNNING,
)


class HistoryPanel:
    """任务历史标签页面板（统计面板 + 列表视图 + 筛选分页 + 详情/清空）"""

    def __init__(self, parent, controller):
        self.parent = parent
        self.controller = controller
        self.gui = controller.gui

        # 分页状态（每页默认 100 条，可在界面切换 50/100/200）
        self.history_page = 1
        self.history_page_size = 100
        self.history_page_size_var = tk.StringVar(value="100")

        self._build_ui()

    # ====================== UI 构建 ======================
    def _build_ui(self):
        parent = self.parent
        # —— 统计面板（成功率 / 平均耗时 / 今日执行 / 历史总计）——
        stats_frame = ttk.Frame(parent)
        stats_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        self.stat_success = self._make_stat_card(stats_frame, "成功率")
        self.stat_avg = self._make_stat_card(stats_frame, "平均耗时")
        self.stat_today = self._make_stat_card(stats_frame, "今日执行")
        self.stat_total = self._make_stat_card(stats_frame, "历史总计")

        # —— 列表视图 ——
        self._build_list(parent)

        self.refresh()

    def _make_stat_card(self, parent, title):
        """生成一个统计卡片（标题 + 大号数值），返回数值标签引用"""
        card = ttk.LabelFrame(parent, text=title, padding=(6, 2), labelanchor="n")
        card.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=3)
        value_label = ttk.Label(
            card, text="—", font=("Microsoft YaHei UI", 15, "bold"),
            foreground="#0078d4", anchor=tk.CENTER
        )
        value_label.pack(fill=tk.X, pady=(2, 2))
        return value_label

    def _build_list(self, parent):
        """构建列表视图：工具条（统计/筛选/清空）+ Treeview"""
        # 工具条
        tool_frame = ttk.Frame(parent)
        tool_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        # 左侧：筛选组（位置固定，不被计数标签宽度影响）
        left_group = ttk.Frame(tool_frame)
        left_group.pack(side=tk.LEFT)

        # 状态筛选
        ttk.Label(left_group, text="状态:").pack(side=tk.LEFT, padx=(0, 2))
        self.history_filter_var = tk.StringVar(value="全部")
        filter_combo = ttk.Combobox(
            left_group, textvariable=self.history_filter_var, state="readonly", width=6
        )
        filter_combo["values"] = ["全部", "成功", "失败", "异常", "已停止", "运行中"]
        filter_combo.pack(side=tk.LEFT)
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self.reset_and_refresh())

        # 时间范围筛选
        ttk.Label(left_group, text="时间:").pack(side=tk.LEFT, padx=(8, 2))
        self.history_range_var = tk.StringVar(value="全部")
        range_combo = ttk.Combobox(
            left_group, textvariable=self.history_range_var, state="readonly", width=6
        )
        range_combo["values"] = ["全部", "今天", "近一周", "近一月"]
        range_combo.pack(side=tk.LEFT)
        range_combo.bind("<<ComboboxSelected>>", lambda e: self.reset_and_refresh())

        # 分类筛选
        ttk.Label(left_group, text="分类:").pack(side=tk.LEFT, padx=(8, 2))
        self.history_category_var = tk.StringVar(value="全部")
        self.history_category_combo = ttk.Combobox(
            left_group, textvariable=self.history_category_var, state="readonly", width=10
        )
        self.history_category_combo["values"] = ["全部"]
        self.history_category_combo.pack(side=tk.LEFT)
        self.history_category_combo.bind("<<ComboboxSelected>>", lambda e: self.reset_and_refresh())

        # 右侧：信息组（每页 + 清空），锚定在右，内部宽度变化不影响左侧
        right_group = ttk.Frame(tool_frame)
        right_group.pack(side=tk.RIGHT)

        ttk.Label(right_group, text="每页:").pack(side=tk.LEFT, padx=(8, 2))
        page_size_combo = ttk.Combobox(
            right_group, textvariable=self.history_page_size_var, state="readonly", width=5
        )
        page_size_combo["values"] = ["50", "100", "200"]
        page_size_combo.pack(side=tk.LEFT)
        page_size_combo.bind("<<ComboboxSelected>>", self.change_page_size)

        # 列表区：独立容器充满中间剩余空间，避免 Treeview 撑满导致分页栏被挤到边上
        list_frame = ttk.Frame(parent)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        columns = ("time", "task", "category", "status", "elapsed")
        self.history_tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=15, selectmode=tk.BROWSE
        )
        self.history_tree.heading("time", text="时间")
        self.history_tree.heading("task", text="任务")
        self.history_tree.heading("category", text="分类")
        self.history_tree.heading("status", text="状态")
        self.history_tree.heading("elapsed", text="耗时")
        self.history_tree.column("time", width=85, stretch=False)
        self.history_tree.column("task", width=120, stretch=True)
        self.history_tree.column("category", width=80, stretch=False)
        self.history_tree.column("status", width=55, stretch=False)
        self.history_tree.column("elapsed", width=55, stretch=False)

        # 状态配色
        self.history_tree.tag_configure("success", foreground="#008000")
        self.history_tree.tag_configure("failed", foreground="#f44747")
        self.history_tree.tag_configure("error", foreground="#f44747")
        self.history_tree.tag_configure("stopped", foreground="#FF8C00")
        self.history_tree.tag_configure("running", foreground="#0000FF")

        v_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=v_scroll.set)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 双击查看详情
        self.history_tree.bind('<Double-Button-1>', self._show_history_detail)

        # —— 底部分页栏 ——
        pager = ttk.Frame(parent)
        pager.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))
        self.history_prev_btn = ttk.Button(
            pager, text="上一页", width=8, command=self.prev_page
        )
        self.history_prev_btn.pack(side=tk.LEFT, padx=2)
        self.history_page_label = ttk.Label(pager, text="第 1 / 1 页", foreground="#444")
        self.history_page_label.pack(side=tk.LEFT, padx=6)
        self.history_next_btn = ttk.Button(
            pager, text="下一页", width=8, command=self.next_page
        )
        self.history_next_btn.pack(side=tk.LEFT, padx=2)

        # 清空记录放右下角，避免被工具栏筛选控件挤掉
        ttk.Button(
            pager, text="清空记录", command=self.clear_history, width=10
        ).pack(side=tk.RIGHT, padx=(0, 2))

    # ====================== 刷新入口 ======================
    def refresh(self):
        """刷新任务历史：统计面板 + 列表视图"""
        self._refresh_stats()
        self._refresh_list()

    def _refresh_stats(self):
        """计算并显示统计卡片：成功率 / 平均耗时 / 今日执行 / 历史总计"""
        records = self.gui.history.records
        total = len(records)
        finished = [r for r in records if r.get("status") in (
            STATUS_SUCCESS, STATUS_FAILED, STATUS_ERROR, STATUS_STOPPED)]
        success = [r for r in finished if r.get("status") == STATUS_SUCCESS]
        today = sum(1 for r in records if self._is_today(r.get("time", "")))
        elapsed_vals = [float(r.get("elapsed", 0) or 0) for r in finished]
        avg = (sum(elapsed_vals) / len(elapsed_vals)) if elapsed_vals else 0.0
        rate = (len(success) / len(finished) * 100) if finished else 0.0

        self.stat_total.config(text=str(total))
        self.stat_today.config(text=str(today))
        self.stat_avg.config(text=self.gui.history.format_elapsed(avg))
        self.stat_success.config(text=f"{rate:.0f}%")
        # 成功率配色
        if not finished:
            self.stat_success.config(foreground="#888888")
        elif rate >= 80:
            self.stat_success.config(foreground="#008000")
        elif rate >= 50:
            self.stat_success.config(foreground="#FF8C00")
        else:
            self.stat_success.config(foreground="#f44747")

    def _refresh_list(self):
        """刷新列表视图（应用状态 + 时间范围筛选，并按页渲染）"""
        tree = self.history_tree
        tree.delete(*tree.get_children())

        # 先按筛选条件过滤（records 最新在前）
        filt = self.history_filter_var.get() if hasattr(self, "history_filter_var") else "全部"
        rng = self.history_range_var.get() if hasattr(self, "history_range_var") else "全部"
        cat = self.history_category_var.get() if hasattr(self, "history_category_var") else "全部"

        # 分类下拉项根据现有记录动态生成（含「全部」）；若当前选择已无对应记录则回退
        cats = sorted({rec.get("category", "") for rec in self.gui.history.records if rec.get("category", "")})
        if hasattr(self, "history_category_combo"):
            self.history_category_combo["values"] = ["全部"] + cats
        if cat != "全部" and cat not in cats:
            self.history_category_var.set("全部")
            cat = "全部"

        filtered = [
            rec for rec in self.gui.history.records
            if (filt == "全部" or rec.get("status", "") == filt)
            and (rng == "全部" or self._in_range(rec.get("time", ""), rng))
            and (cat == "全部" or rec.get("category", "") == cat)
        ]

        total = len(filtered)
        size = self.history_page_size
        total_pages = max(1, (total + size - 1) // size)
        if self.history_page > total_pages:
            self.history_page = total_pages
        if self.history_page < 1:
            self.history_page = 1

        status_tag = {
            STATUS_SUCCESS: "success",
            STATUS_FAILED: "failed",
            STATUS_ERROR: "error",
            STATUS_STOPPED: "stopped",
            STATUS_RUNNING: "running",
        }
        start = (self.history_page - 1) * size
        for rec in filtered[start:start + size]:
            status = rec.get("status", "")
            tag = status_tag.get(status, "")
            elapsed = self.gui.history.format_elapsed(rec.get("elapsed", 0))
            tree.insert(
                "", tk.END,
                iid=str(rec["id"]),
                values=(
                    self._concise_time(rec.get("time", "")),
                    rec.get("task", ""),
                    rec.get("category", ""),
                    status,
                    elapsed,
                ),
                tags=(tag,) if tag else (),
            )

        # 计数 + 页码（合并到分页栏，避免工具栏拥挤）
        self.history_page_label.config(text=f"共 {total} 条 · 第 {self.history_page} / {total_pages} 页")
        self.history_prev_btn.config(
            state="normal" if self.history_page > 1 else "disabled"
        )
        self.history_next_btn.config(
            state="normal" if self.history_page < total_pages else "disabled"
        )

    # ---------- 分页控制 ----------
    def prev_page(self):
        """上一页"""
        if self.history_page > 1:
            self.history_page -= 1
            self._refresh_list()

    def next_page(self):
        """下一页"""
        self.history_page += 1
        self._refresh_list()

    def change_page_size(self, event=None):
        """切换每页条数后回到第 1 页"""
        try:
            self.history_page_size = int(self.history_page_size_var.get())
        except ValueError:
            self.history_page_size = 100
        self.history_page = 1
        self._refresh_list()

    def reset_and_refresh(self):
        """切换筛选 / 时间范围时回到第 1 页再刷新"""
        self.history_page = 1
        self.refresh()

    def reset_page(self):
        """新任务置顶后回到第 1 页（由主窗口在新增记录时调用）"""
        self.history_page = 1
        self.refresh()

    # ---------- 时间格式化工具 ----------
    @staticmethod
    def _concise_time(time_str):
        """把 YYYY-MM-DD HH:MM:SS 转为紧凑展示：
        今天 -> HH:MM；今年 -> MM-DD HH:MM；跨年 -> YYYY-MM-DD"""
        try:
            dt = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return time_str
        now = datetime.now()
        if dt.date() == now.date():
            return dt.strftime("%H:%M")
        if dt.year == now.year:
            return dt.strftime("%m-%d %H:%M")
        return dt.strftime("%Y-%m-%d")

    @staticmethod
    def _relative_time(time_str):
        """相对时间：刚刚 / N分钟前 / N小时前 / N天前；跨月返回空串"""
        try:
            dt = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return ""
        delta = datetime.now() - dt
        sec = delta.total_seconds()
        if sec < 0:
            return ""
        if sec < 60:
            return "刚刚"
        if sec < 3600:
            return f"{int(sec // 60)}分钟前"
        if sec < 86400:
            return f"{int(sec // 3600)}小时前"
        if sec < 86400 * 30:
            return f"{int(sec // 86400)}天前"
        return ""

    @staticmethod
    def _is_today(time_str):
        """判断记录是否为今天"""
        try:
            dt = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return False
        return dt.date() == datetime.now().date()

    @staticmethod
    def _in_range(time_str, rng):
        """判断记录时间是否落在指定范围内（今天 / 近一周 / 近一月）"""
        try:
            dt = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
        except Exception:
            return False
        now = datetime.now()
        if rng == "今天":
            return dt.date() == now.date()
        if rng == "近一周":
            return 0 <= (now - dt).total_seconds() < 7 * 86400
        if rng == "近一月":
            return 0 <= (now - dt).total_seconds() < 30 * 86400
        return True

    # ---------- 详情 / 清空 ----------
    def clear_history(self):
        """清空任务历史（面板内「清空记录」按钮调用）"""
        if not self.gui.history.records:
            return
        if messagebox.askyesno("确认", "确定清空所有任务历史记录？"):
            self.gui.history.clear()
            self.refresh()
            self.gui._log("[历史] 已清空任务历史记录")

    def _show_history_detail(self, event):
        """双击列表查看任务详情"""
        sel = self.history_tree.selection()
        if not sel:
            return
        self.show_detail_by_id(int(sel[0]))

    def show_detail_by_id(self, rec_id):
        """按记录 id 查看任务详情"""
        rec = next((r for r in self.gui.history.records if r["id"] == rec_id), None)
        if not rec:
            return
        detail = rec.get("detail", "")
        rel = self._relative_time(rec.get("time", ""))
        msg = (
            f"时间: {rec.get('time', '')}"
            + (f" ({rel})" if rel else "")
            + f"\n任务: {rec.get('task', '')}\n"
            f"分类: {rec.get('category', '')}\n"
            f"状态: {rec.get('status', '')}\n"
            f"耗时: {self.gui.history.format_elapsed(rec.get('elapsed', 0))}"
        )
        if detail:
            msg += f"\n\n详情: {detail}"
        messagebox.showinfo("任务详情", msg)
