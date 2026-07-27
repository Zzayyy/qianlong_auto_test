# -*- coding: utf-8 -*-
"""交易系统设置报告中心。"""

from __future__ import annotations

import json
from pathlib import Path
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox

from config import get_client_name, get_scripts_config
from gui.recycle_bin import (
    move_directory_to_recycle_bin,
    validate_batch_directory,
)
from gui.shell_open import open_path
from core.settings_report import (
    DEFAULT_BATCH_LIMIT,
    OVERALL_FAIL,
    OVERALL_PASS,
    OVERALL_REVIEW,
    OVERALL_RUNNING,
    discover_batches,
)

# 结论/状态语义色（与模块明细树、Excel 报告保持一致）
COLOR_PASS = "#067647"
COLOR_REVIEW = "#9a6700"
COLOR_FAIL = "#b42318"
COLOR_RUNNING = "#175cd3"
COLOR_NEUTRAL = "#333333"


class SettingsReportPanel:
    """展示设置检查批次，并发起完整的一键检查。"""

    def __init__(self, parent, controller):
        self.parent = parent
        self.controller = controller
        self._batch_lookup = {}
        self._history_batches = []
        self._current_summary = None
        self._running_batch = None
        self._running_run_id = ""
        self._card_labels = {}

        self.batch_var = tk.StringVar()
        self.progress_var = tk.StringVar(value="尚未运行完整检查")
        self.overall_var = tk.StringVar(value="—")
        self.module_count_var = tk.StringVar(value="0")
        self.pass_count_var = tk.StringVar(value="0")
        self.diff_count_var = tk.StringVar(value="0")
        self.review_count_var = tk.StringVar(value="0")

        self._build_ui()
        self.refresh_batches()

    def _build_ui(self):
        client_name = get_client_name(self.controller.client_id) or self.controller.client_id or "未设置"
        header = tk.Frame(self.parent)
        header.pack(fill=tk.X, pady=(0, 8))
        tk.Label(
            header,
            text="交易系统设置 · 差异报告中心",
            anchor=tk.W,
            font=("Microsoft YaHei UI", 15, "bold"),
            foreground="#1f2d3d",
        ).pack(fill=tk.X)


        toolbar = ttk.Frame(self.parent)
        toolbar.pack(fill=tk.X, pady=(0, 6))

        self.run_button = ttk.Button(
            toolbar,
            text="一键执行全部设置",
            command=self.start_full_check,
        )
        self.run_button.pack(side=tk.LEFT)
        self.stop_button = ttk.Button(
            toolbar,
            text="停止",
            command=self.stop,
            state=tk.DISABLED,
        )
        self.stop_button.pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(
            toolbar,
            text="刷新",
            command=self.refresh_batches,
        ).pack(side=tk.RIGHT)

        history = ttk.Frame(self.parent)
        history.pack(fill=tk.X, pady=(0, 6))
        ttk.Label(
            history,
            text=f"历史批次（最近{DEFAULT_BATCH_LIMIT}批）:",
        ).pack(side=tk.LEFT)
        self.batch_combo = ttk.Combobox(
            history,
            textvariable=self.batch_var,
            state="readonly",
            width=42,
        )
        self.batch_combo.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        self.batch_combo.bind("<<ComboboxSelected>>", self._on_batch_selected)

        ttk.Label(
            self.parent,
            textvariable=self.progress_var,
            foreground="#555555",
        ).pack(fill=tk.X, pady=(0, 6))

        cards = ttk.Frame(self.parent)
        cards.pack(fill=tk.X, pady=(0, 6))
        self._make_card(cards, "总体结论", self.overall_var, 0)
        self._make_card(cards, "模块", self.module_count_var, 1)
        self._make_card(cards, "通过项", self.pass_count_var, 2)
        self._make_card(cards, "差异/失败", self.diff_count_var, 3)
        self._make_card(cards, "待复核", self.review_count_var, 4)
        for column in range(5):
            cards.columnconfigure(column, weight=1)

        self.detail_notebook = ttk.Notebook(self.parent)
        self.detail_notebook.pack(fill=tk.BOTH, expand=True)

        module_frame = ttk.Frame(self.detail_notebook)
        problem_frame = ttk.Frame(self.detail_notebook)
        self.detail_notebook.add(module_frame, text="模块汇总")
        self.detail_notebook.add(problem_frame, text="问题明细")

        self.module_tree = self._make_tree(
            module_frame,
            columns=(
                "module", "conclusion", "total", "passed",
                "differences", "unverified", "disabled", "elapsed",
            ),
            headings=(
                "模块", "结论", "检查项", "通过", "差异",
                "待复核", "未启用", "耗时(秒)",
            ),
            widths=(80, 60, 65, 55, 55, 65, 65, 70),
        )
        self.module_tree.bind("<Double-1>", self._open_selected_module_report)

        self.problem_tree = self._make_tree(
            problem_frame,
            columns=("status", "module", "name", "expected", "actual", "detail"),
            headings=("状态", "模块", "检查项", "期望值", "实际值", "说明"),
            widths=(80, 110, 150, 160, 160, 230),
        )
        self.problem_tree.bind("<Double-1>", self._open_selected_problem_report)

        footer = ttk.Frame(self.parent)
        footer.pack(fill=tk.X, pady=(6, 0))
        ttk.Button(
            footer,
            text="打开批次目录",
            command=self.open_batch_dir,
        ).pack(side=tk.LEFT)
        ttk.Button(
            footer,
            text="打开总TXT",
            command=lambda: self._open_current_path("txt_path"),
        ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(
            footer,
            text="打开总Excel",
            command=lambda: self._open_current_path("xlsx_path"),
        ).pack(side=tk.LEFT, padx=(5, 0))
        ttk.Button(
            footer,
            text="删除选中批次",
            command=self.delete_selected_batch,
        ).pack(side=tk.RIGHT)

    def _make_card(self, parent, title, variable, column):
        frame = ttk.LabelFrame(parent, text=title, padding=(5, 3))
        frame.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 3, 0))
        label = tk.Label(
            frame,
            textvariable=variable,
            anchor=tk.CENTER,
            font=("Microsoft YaHei UI", 13, "bold"),
            foreground=COLOR_NEUTRAL,
        )
        label.pack(fill=tk.X, ipady=2)
        self._card_labels[title] = label

    @staticmethod
    def _make_tree(parent, columns, headings, widths):
        wrap = ttk.Frame(parent)
        wrap.pack(fill=tk.BOTH, expand=True)
        tree = ttk.Treeview(wrap, columns=columns, show="headings", height=12)
        ybar = ttk.Scrollbar(wrap, orient=tk.VERTICAL, command=tree.yview)
        xbar = ttk.Scrollbar(wrap, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=ybar.set, xscrollcommand=xbar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        wrap.rowconfigure(0, weight=1)
        wrap.columnconfigure(0, weight=1)
        for column, heading, width in zip(columns, headings, widths):
            tree.heading(column, text=heading)
            tree.column(column, width=width, minwidth=45, stretch=True)
        tree.tag_configure("fail", foreground="#b42318")
        tree.tag_configure("review", foreground="#9a6700")
        tree.tag_configure("pass", foreground="#067647")
        return tree

    def start_full_check(self):
        """按当前客户端支持清单，顺序执行全部交易系统设置模块。"""
        if self.controller.runner.is_running or self.controller.task_center.is_running:
            messagebox.showwarning("报告中心", "当前已有任务正在运行，请结束后再执行")
            return

        scripts = get_scripts_config(self.controller.client_id).get("交易系统设置", [])
        if not scripts:
            messagebox.showwarning("报告中心", "当前客户端没有可执行的交易系统设置模块")
            return

        output_dir = self.controller.settings_output_dir.get().strip()
        if not output_dir:
            messagebox.showwarning("报告中心", "请先设置交易系统设置的输出目录")
            return

        tasks = []
        base_params = self.controller.collect_params()
        for script in scripts:
            params = dict(base_params)
            tasks.append({
                "category": "交易系统设置",
                "script_name": script["name"],
                "script_path": script["path"],
                "query_key": script.get("query_key", ""),
                "params": params,
                "status": self.controller.task_center.ST_PENDING,
            })

        try:
            batch = self.controller.begin_settings_batch(
                "报告中心一键运行", tasks
            )
        except OSError as exc:
            messagebox.showerror("报告中心", f"无法创建报告批次:\n{exc}")
            return

        self._running_batch = batch
        self._running_run_id = batch["run_id"]
        self.run_button.config(state=tk.DISABLED)
        self.stop_button.config(state=tk.NORMAL)
        self.progress_var.set(
            f"批次 {self._running_run_id}：准备执行 0/{len(tasks)}"
        )

        started = self.controller.task_center.start_transient(
            tasks,
            on_complete=self._on_batch_complete,
            on_progress=self._on_batch_progress,
            label="设置完整检查",
            settings_batch=batch,
        )
        if not started:
            self.controller.update_settings_batch(
                batch, tasks, final=True, stopped=True
            )
            self.run_button.config(state=tk.NORMAL)
            self.stop_button.config(state=tk.DISABLED)
            self.progress_var.set("启动失败：当前执行器正忙")
            messagebox.showwarning("报告中心", "当前执行器正忙，未启动完整检查")

    def stop(self):
        self.controller.task_center.stop()

    def _on_batch_progress(self, completed, total, task):
        self.progress_var.set(
            f"批次 {self._running_run_id}：已完成 {completed}/{total}，"
            f"{task.get('script_name', '')} → {task.get('status', '')}"
        )

    def _on_batch_complete(self, task_records, stopped):
        run_id = self._running_run_id
        self.run_button.config(state=tk.NORMAL)
        self.stop_button.config(state=tk.DISABLED)
        summary = (self._running_batch or {}).get("summary")
        if not summary:
            self.progress_var.set(f"批次 {run_id}：汇总报告生成失败")
            messagebox.showerror("报告中心", "汇总报告生成失败，请查看运行日志")
            return

        self.controller._log(
            f"[报告中心] 批次完成: {summary['overall_status']} | "
            f"{summary['batch_dir']}"
        )
        self.progress_var.set(
            f"批次 {run_id}：{summary['overall_status']}，"
            f"已生成总TXT、总Excel和批次汇总JSON | "
            f"{summary['batch_dir']}"
        )
        if summary["problems"]:
            self.detail_notebook.select(1)
        self.controller.show_report_center(auto_clear=False)
        self._running_batch = None

    def on_batch_summary(self, summary, final=False):
        """直接接收本次批次结果，不重新扫描输出目录。"""
        run_id = summary.get("run_id", "")
        batches = [summary]
        batches.extend(
            item
            for item in self._history_batches
            if item.get("run_id", "") != run_id
        )
        self._render_batch_choices(
            batches[:DEFAULT_BATCH_LIMIT],
            select_run_id=run_id,
        )
        if final:
            self.progress_var.set(
                f"{summary.get('source', '设置检查')}批次 {run_id}："
                f"{summary.get('overall_status', '')} | "
                f"{summary.get('batch_dir', '')}"
            )

    def refresh_batches(self, select_run_id=None):
        """仅扫描并读取当前输出目录中最新的 20 个批次。"""
        batches = discover_batches(
            self.controller.settings_output_dir.get().strip(),
            limit=DEFAULT_BATCH_LIMIT,
        )
        self._render_batch_choices(batches, select_run_id=select_run_id)

    def _render_batch_choices(self, batches, select_run_id=None):
        """用已取得的批次数据更新下拉框和当前报告。"""
        self._history_batches = list(batches[:DEFAULT_BATCH_LIMIT])
        self._batch_lookup = {}
        values = []
        for summary in self._history_batches:
            run_id = summary.get("run_id", "")
            client_name = get_client_name(summary.get("client_id", "")) or summary.get("client_id", "")
            source = summary.get("source", "未知来源")
            batch_status = summary.get("batch_status", "")
            display = (
                f"{run_id} | {source} | {batch_status} | "
                f"{summary.get('overall_status', '未知')} | {client_name}"
            )
            values.append(display)
            self._batch_lookup[display] = summary
        self.batch_combo["values"] = values

        if select_run_id:
            selected = next(
                (value for value in values if value.startswith(f"{select_run_id} |")),
                None,
            )
            if selected:
                self.batch_var.set(selected)
                self._load_summary(self._batch_lookup[selected])
                return

        # 普通刷新（打开报告中心、切换输出目录等）不自动展示历史批次，保持清空
        self.batch_var.set("")
        self._clear_summary()

    def _on_batch_selected(self, event=None):
        summary = self._batch_lookup.get(self.batch_var.get())
        if summary:
            self._load_summary(summary)

    def _load_summary(self, summary):
        self._current_summary = summary
        totals = summary.get("totals", {})
        self.overall_var.set(summary.get("overall_status", "—"))
        self.module_count_var.set(str(totals.get("模块数", len(summary.get("modules", [])))))
        self.pass_count_var.set(str(totals.get("通过", 0)))
        diff_or_fail = int(totals.get("差异合计", 0) or 0) + int(
            totals.get("执行失败", 0) or 0
        )
        self.diff_count_var.set(str(diff_or_fail))
        review_count = int(totals.get("未验证", 0) or 0)
        self.review_count_var.set(str(review_count))

        self._color_cards(summary.get("overall_status", ""), diff_or_fail, review_count)

        self.module_tree.delete(*self.module_tree.get_children())
        for index, module in enumerate(summary.get("modules", [])):
            conclusion = module.get("conclusion", "")
            tag = self._status_tag(conclusion)
            self.module_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    module.get("module", ""),
                    conclusion,
                    module.get("total", 0),
                    module.get("passed", 0),
                    module.get("differences", 0),
                    module.get("unverified", 0),
                    module.get("disabled", 0),
                    f"{float(module.get('elapsed', 0) or 0):.1f}",
                ),
                tags=(tag,),
            )

        self.problem_tree.delete(*self.problem_tree.get_children())
        for index, problem in enumerate(summary.get("problems", [])):
            self.problem_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    problem.get("status", ""),
                    problem.get("module", ""),
                    problem.get("name", ""),
                    self._display_value(problem.get("expected", "")),
                    self._display_value(problem.get("actual", "")),
                    problem.get("detail", ""),
                ),
                tags=(self._status_tag(problem.get("status", "")),),
            )

    def clear_view(self):
        """清空当前展示的批次总结（保留历史下拉列表）。"""
        self._clear_summary()

    def _clear_summary(self):
        self._current_summary = None
        for variable in (
            self.overall_var,
            self.module_count_var,
            self.pass_count_var,
            self.diff_count_var,
            self.review_count_var,
        ):
            variable.set("—" if variable is self.overall_var else "0")
        for label in self._card_labels.values():
            label.config(foreground=COLOR_NEUTRAL)
        self._card_labels["通过项"].config(foreground=COLOR_PASS)
        if getattr(self, "batch_combo", None) is not None:
            self.batch_combo.set("")
        self.module_tree.delete(*self.module_tree.get_children())
        self.problem_tree.delete(*self.problem_tree.get_children())

    def _color_cards(self, overall_status, diff_or_fail, review_count):
        """按结论与计数给卡片数值设置语义色。"""
        overall_color = {
            OVERALL_PASS: COLOR_PASS,
            OVERALL_REVIEW: COLOR_REVIEW,
            OVERALL_FAIL: COLOR_FAIL,
            OVERALL_RUNNING: COLOR_RUNNING,
        }.get(overall_status, COLOR_NEUTRAL)
        self._card_labels["总体结论"].config(foreground=overall_color)
        self._card_labels["通过项"].config(foreground=COLOR_PASS)
        self._card_labels["差异/失败"].config(
            foreground=COLOR_FAIL if diff_or_fail else COLOR_PASS
        )
        self._card_labels["待复核"].config(
            foreground=COLOR_REVIEW if review_count else COLOR_PASS
        )

    @staticmethod
    def _status_tag(status):
        if status in (OVERALL_FAIL, "差异", "新增", "冲突", "执行失败"):
            return "fail"
        if status in (OVERALL_REVIEW, "未验证"):
            return "review"
        return "pass"

    @staticmethod
    def _display_value(value):
        if isinstance(value, (dict, list, tuple)):
            return json.dumps(value, ensure_ascii=False, default=str)
        return str(value)

    def open_batch_dir(self):
        if self._current_summary:
            self._open_path(self._current_summary.get("batch_dir", ""))

    def delete_selected_batch(self):
        """二次确认后，将当前批次目录整体移动到 Windows 回收站。"""
        summary = self._current_summary
        if not summary:
            messagebox.showwarning("报告中心", "请先选择一个历史批次")
            return
        if summary.get("batch_status") == "运行中":
            messagebox.showwarning("报告中心", "当前批次仍在运行，不能删除")
            return

        run_id = str(summary.get("run_id", ""))
        try:
            target = validate_batch_directory(
                summary.get("batch_dir", ""),
                self.controller.settings_output_dir.get().strip(),
                run_id,
            )
        except (OSError, ValueError) as exc:
            messagebox.showerror("报告中心", f"批次路径校验失败:\n{exc}")
            return

        confirmed = messagebox.askyesno(
            "删除选中批次",
            "确定要将以下批次移入 Windows 回收站吗？\n\n"
            f"批次号: {run_id}\n"
            f"运行来源: {summary.get('source', '未知')}\n"
            f"生成时间: {summary.get('generated_at', '')}\n"
            f"目录: {target}\n\n"
            "该操作会移动整个批次目录，可从回收站恢复。",
            icon=messagebox.WARNING,
        )
        if not confirmed:
            return

        try:
            move_directory_to_recycle_bin(target)
        except (OSError, subprocess.SubprocessError) as exc:
            messagebox.showerror("报告中心", f"移入回收站失败:\n{exc}")
            return

        self.controller._log(f"[报告中心] 已移入回收站: {target}")
        self.progress_var.set(f"批次 {run_id} 已移入回收站")
        self._current_summary = None
        self.refresh_batches()
        messagebox.showinfo("报告中心", f"批次 {run_id} 已移入回收站")

    def _open_current_path(self, key):
        if self._current_summary:
            self._open_path(self._current_summary.get(key, ""))

    def _open_selected_module_report(self, event=None):
        selection = self.module_tree.selection()
        if not selection or not self._current_summary:
            return
        index = int(selection[0])
        modules = self._current_summary.get("modules", [])
        if index < len(modules):
            self._open_path(modules[index].get("report_path", ""))

    def _open_selected_problem_report(self, event=None):
        selection = self.problem_tree.selection()
        if not selection or not self._current_summary:
            return
        index = int(selection[0])
        problems = self._current_summary.get("problems", [])
        if index < len(problems):
            self._open_path(problems[index].get("report_path", ""))

    @staticmethod
    def _open_path(path):
        if not path:
            return
        target = Path(path)
        if not target.exists():
            messagebox.showwarning("报告中心", f"文件或目录不存在:\n{path}")
            return
        try:
            open_path(target)
        except OSError as exc:
            messagebox.showerror("报告中心", f"无法打开:\n{exc}")
