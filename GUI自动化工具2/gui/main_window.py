# -*- coding: utf-8 -*-
"""GUI自动化主窗口：界面构建、参数配置、日志展示与执行编排"""

import os
import sys
import time
import logging
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    import openpyxl
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

from config import (
    SCRIPTS_CONFIG,
    load_user_config,
    save_user_config,
    get_output_dir,
    set_output_dir,
    get_script_filename,
    IS_FROZEN,
    PROJECT_ROOT,
)
from engine.runner import ScriptRunner
from engine.task import Task
from gui.widgets import ColoredLogText
from gui.task_center import TaskCenter
from gui.history import (
    HistoryManager,
    STATUS_SUCCESS,
    STATUS_FAILED,
    STATUS_ERROR,
    STATUS_STOPPED,
    STATUS_RUNNING,
)


class AutomationGUI:
    """GUI自动化主界面"""

    # 组合申报：全自动脚本只配置委托数量，查询类脚本配置导出参数
    COMBO_AUTO_SCRIPTS = ("1.组合申报_全自动", "2.拆分申报_全自动")

    def __init__(self, root):
        self.root = root
        self.root.title("钱龙期权交易 - GUI自动化工具")
        self.root.geometry("950x750")
        self.root.minsize(850, 650)

        self.is_running = False
        self.current_category = "查询"

        # 自身引用（供任务中心等子模块访问主窗口能力）
        self.gui = self
        # 任务中心顺序执行模式开关
        self._task_mode = False
        self.task_center = None

        # 脚本列表 -> 任务队列 的拖拽状态
        self._drag_script = None
        self._drag_active = False
        self._drag_start_y = 0

        # 状态栏状态
        self._status_running = False
        self.task_start_time = 0.0

        # 加载用户配置
        self.user_config = load_user_config()

        # 参数变量（使用配置中的默认值）
        self.export_format = tk.StringVar(value=self.user_config.get("export_format", "txt"))
        self.auto_open = tk.BooleanVar(value=self.user_config.get("auto_open", False))
        self.log_level = tk.StringVar(value=self.user_config.get("log_level", "详细"))
        self.txt_path = tk.StringVar(value="")
        self.xls_path = tk.StringVar(value="")
        self.order_qty = tk.IntVar(value=1)
        self.countdown_sec = tk.IntVar(value=3)
        self.xlsx_file = tk.StringVar(value="")

        # 期权下单_一键导出 参数
        self.export_target_position = tk.BooleanVar(value=True)  # 持仓
        self.export_target_order = tk.BooleanVar(value=True)     # 委托
        self.export_output_dir = tk.StringVar(value=get_output_dir(self.user_config, "下单"))

        # 交易系统设置 参数（输出路径可自定义）
        self.settings_output_dir = tk.StringVar(value=get_output_dir(self.user_config, "交易系统设置"))

        # 日志目录：打包后放在exe同级目录
        if IS_FROZEN:
            self.log_dir = os.path.join(os.path.dirname(sys.executable), "logs")
        else:
            self.log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "logs")
        os.makedirs(self.log_dir, exist_ok=True)

        self._setup_logging()
        self._setup_runner()

        # 任务历史管理器（持久化到日志目录）
        self.history = HistoryManager(self.log_dir)
        self._current_record_id = None  # 当前正在运行的记录 id

        self._build_ui()
        self.logger.info("GUI自动化工具启动")

    def _setup_runner(self):
        """创建脚本执行引擎，并注入回调"""
        self.runner = ScriptRunner(
            is_frozen=IS_FROZEN,
            project_root=PROJECT_ROOT,
            log_level_getter=lambda: self.log_level.get(),
            on_log=self._log,
            on_debug=self.logger.debug,
            on_finish=self._on_run_finish,
            on_error=self._on_run_error,
        )

    def _setup_logging(self):
        """配置日志系统"""
        self.logger = logging.getLogger("AutomationGUI")
        self.logger.setLevel(logging.DEBUG)
        log_file = os.path.join(
            self.log_dir,
            f"gui_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        fmt = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(fmt)
        self.logger.addHandler(file_handler)

    def _build_ui(self):
        """构建界面"""
        # 菜单栏
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # 功能菜单
        func_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="功能", menu=func_menu)
        func_menu.add_command(label="下单", command=lambda: self._switch_category("下单"))
        func_menu.add_command(label="组合申报", command=lambda: self._switch_category("组合申报"))
        func_menu.add_command(label="查询", command=lambda: self._switch_category("查询"))
        func_menu.add_command(label="交易系统设置", command=lambda: self._switch_category("交易系统设置"))

        # 工具菜单
        tool_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="工具", menu=tool_menu)
        tool_menu.add_command(label="清空日志", command=self._clear_log)
        tool_menu.add_command(label="打开日志目录", command=self._open_log_dir)
        tool_menu.add_separator()
        tool_menu.add_command(label="清空任务历史", command=self._clear_history)
        tool_menu.add_separator()
        tool_menu.add_command(label="退出", command=self.root.quit)

        # 设置菜单
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="设置", menu=settings_menu)
        settings_menu.add_radiobutton(
            label="详细日志",
            variable=self.log_level,
            value="详细",
            command=self._on_log_level_change
        )
        settings_menu.add_radiobutton(
            label="简洁日志",
            variable=self.log_level,
            value="简洁",
            command=self._on_log_level_change
        )
        settings_menu.add_separator()
        settings_menu.add_command(label="日志级别说明", command=self._show_log_level_help)

        # 主框架
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 标题
        ttk.Label(
            main_frame,
            text="钱龙期权交易自动化",
            font=("Microsoft YaHei UI", 16, "bold")
        ).pack(pady=(0, 10))

        # 左右分栏（PanedWindow 保证日志区域始终可见）
        self.paned = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        self.paned.pack(fill=tk.BOTH, expand=True)

        # 左侧面板：脚本列表 + 参数配置
        left_frame = ttk.Frame(self.paned)
        self.paned.add(left_frame, weight=1)

        # 操作按钮 - 固定在底部
        self.btn_frame = ttk.Frame(left_frame)
        self.btn_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

        self.execute_btn = ttk.Button(self.btn_frame, text="执行", command=self._execute_script, width=12)
        self.execute_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(self.btn_frame, text="停止", command=self._stop_script, state=tk.DISABLED, width=10)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        # 当前分类标签
        self.category_label = ttk.Label(
            left_frame,
            text="",
            font=("Microsoft YaHei UI", 12, "bold"),
            foreground="#0078d4"
        )
        self.category_label.pack(side=tk.TOP, anchor=tk.W, pady=(0, 5))

        # 中间内容区域 - 占据剩余空间
        middle_frame = ttk.Frame(left_frame)
        middle_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 10))

        # 脚本列表
        list_frame = ttk.LabelFrame(middle_frame, text="脚本列表", padding="8")
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        self.script_listbox = tk.Listbox(
            list_frame,
            font=("Microsoft YaHei UI", 10),
            selectmode=tk.SINGLE,
            exportselection=False,
            height=12
        )
        self.script_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.script_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.script_listbox.config(yscrollcommand=scrollbar.set)
        self.script_listbox.bind('<Double-Button-1>', lambda e: self._execute_script())
        self.script_listbox.bind('<<ListboxSelect>>', lambda e: (self._update_paths_for_selected_script(), self._update_params_for_selected_script()))
        # 从脚本列表拖入任务队列
        self.script_listbox.bind('<ButtonPress-1>', self._on_list_drag_start)
        self.script_listbox.bind('<B1-Motion>', self._on_list_drag_motion)
        self.script_listbox.bind('<ButtonRelease-1>', self._on_list_drag_end)

        # 参数配置面板
        self.params_frame = ttk.LabelFrame(middle_frame, text="参数配置", padding="8")
        self.params_frame.pack(fill=tk.X, pady=(0, 10))

        # 数据预览面板（仅下单显示）
        self.preview_frame = ttk.LabelFrame(middle_frame, text="数据预览", padding="3")
        # 初始隐藏

        # 预览Treeview + 滚动条
        tree_wrap = ttk.Frame(self.preview_frame)
        tree_wrap.pack(fill=tk.X, anchor=tk.W)

        self.preview_tree = ttk.Treeview(tree_wrap, show="headings", height=5)
        v_scroll = ttk.Scrollbar(tree_wrap, orient=tk.VERTICAL, command=self.preview_tree.yview)
        h_scroll = ttk.Scrollbar(self.preview_frame, orient=tk.HORIZONTAL, command=self.preview_tree.xview)
        self.preview_tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        self.preview_tree.pack(side=tk.LEFT, fill=tk.BOTH)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # 预览信息标签
        self.preview_info = ttk.Label(self.preview_frame, text="选择Excel文件后显示数据预览", foreground="gray")
        self.preview_info.pack(anchor=tk.W, pady=(3, 0))

        # 右侧：使用 Notebook 容纳「运行日志」与「任务历史」
        self.right_notebook = ttk.Notebook(self.paned)
        self.paned.add(self.right_notebook, weight=1)

        # —— 运行日志标签页 ——
        log_frame = ttk.Frame(self.right_notebook, padding="5")
        self.right_notebook.add(log_frame, text="运行日志")

        self.log_text = ColoredLogText(
            log_frame,
            font=("Consolas", 9),
            state=tk.DISABLED,
            wrap=tk.WORD,
            bg="#1e1e1e",
            fg="#d4d4d4",
            insertbackground="#d4d4d4",
            width=30
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # —— 任务历史标签页 ——
        history_frame = ttk.Frame(self.right_notebook, padding="5")
        self.right_notebook.add(history_frame, text="任务历史")
        self._build_history_panel(history_frame)

        # —— 任务中心标签页 ——
        task_frame = ttk.Frame(self.right_notebook, padding="5")
        self.right_notebook.add(task_frame, text="任务中心")
        self.task_center = TaskCenter(task_frame, self)

        # 状态栏
        status_frame = ttk.Frame(self.root)
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Separator(self.root, orient=tk.HORIZONTAL).pack(side=tk.BOTTOM, fill=tk.X)

        self.status_label = ttk.Label(
            status_frame, text="就绪", anchor=tk.W, relief=tk.SUNKEN, padding=(6, 3)
        )
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.level_label = ttk.Label(
            status_frame, text=f"日志: {self.log_level.get()}", anchor=tk.CENTER,
            relief=tk.SUNKEN, width=12, padding=(6, 3)
        )
        self.level_label.pack(side=tk.LEFT, padx=(2, 0))

        self.time_label = ttk.Label(
            status_frame, text="", anchor=tk.E, relief=tk.SUNKEN, width=18, padding=(6, 3)
        )
        self.time_label.pack(side=tk.LEFT, padx=(2, 0))

        # 默认显示下单
        self._switch_category("下单")

        # 左右分栏比例：sash 位置占整体宽度的比例
        #   0.5 = 五五分 | 0.6 = 左6右4 | 0.8 = 左8右2（改这里即可调默认比例）
        self.pane_ratio = 0.5
        # 必须等窗口真正显示（<Map>）后再设置 sash，否则 winfo_width 为 1 -> 比例失效
        self.root.bind("<Map>", self._on_first_map)
        # 全局捕获鼠标释放，用于「从脚本列表拖入任务队列」的落点判定
        self.root.bind("<ButtonRelease-1>", self._on_global_drop)
        # 窗口缩放时按比例保持
        self.paned.bind("<Configure>", lambda e: self._apply_pane_ratio())

    def _on_first_map(self, event):
        """窗口首次显示后应用一次分栏比例，随后解绑"""
        self.root.unbind("<Map>")
        self._apply_pane_ratio()

    def _apply_pane_ratio(self):
        """按固定比例设置左右分隔条位置（不受右侧标签页数量影响）"""
        w = self.paned.winfo_width()
        if w <= 1:
            return  # 窗口尚未绘制，宽度无效，跳过
        target = int(w * self.pane_ratio)
        if getattr(self, "_last_sash", None) == target:
            return  # 位置未变则跳过，避免与拖动/自身触发形成死循环
        self._last_sash = target
        self.paned.sashpos(0, target)

    def _switch_category(self, category):
        """切换分类"""
        self.current_category = category
        self.category_label.config(text=f"当前功能: {category}")
        self.script_listbox.delete(0, tk.END)

        scripts = SCRIPTS_CONFIG.get(category, [])
        for script in scripts:
            self.script_listbox.insert(tk.END, script["name"])

        # 重建参数面板
        self._rebuild_params()

        # 仅下单显示数据预览
        if self.current_category == "下单":
            self.preview_frame.pack(fill=tk.X, pady=(0, 10))
        else:
            self.preview_frame.pack_forget()

        self._log(f"切换分类: {category}")
        self.logger.info(f"切换分类: {category}")

        # 空闲时同步状态栏（运行中不覆盖）
        if not self.is_running:
            self._set_status(f"就绪 - 当前功能: {category}")

    # ====================== 任务历史标签页 ======================
    def _build_history_panel(self, parent):
        """构建任务历史标签页：统计面板 + 列表视图"""
        # —— 统计面板（成功率 / 平均耗时 / 今日执行 / 历史总计）——
        stats_frame = ttk.Frame(parent)
        stats_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        self.stat_success = self._make_stat_card(stats_frame, "成功率")
        self.stat_avg = self._make_stat_card(stats_frame, "平均耗时")
        self.stat_today = self._make_stat_card(stats_frame, "今日执行")
        self.stat_total = self._make_stat_card(stats_frame, "历史总计")

        # —— 列表视图 ——
        self._build_history_list(parent)

        self._refresh_history()

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

    # ---------- 列表视图 ----------
    def _build_history_list(self, parent):
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
            left_group, textvariable=self.history_filter_var, state="readonly", width=8
        )
        filter_combo["values"] = ["全部", "成功", "失败", "异常", "已停止", "运行中"]
        filter_combo.pack(side=tk.LEFT)
        filter_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_history())

        # 时间范围筛选
        ttk.Label(left_group, text="时间:").pack(side=tk.LEFT, padx=(8, 2))
        self.history_range_var = tk.StringVar(value="全部")
        range_combo = ttk.Combobox(
            left_group, textvariable=self.history_range_var, state="readonly", width=8
        )
        range_combo["values"] = ["全部", "今天", "近一周", "近一月"]
        range_combo.pack(side=tk.LEFT)
        range_combo.bind("<<ComboboxSelected>>", lambda e: self._refresh_history())

        # 右侧：信息组（计数 + 清空），锚定在右，内部宽度变化不影响左侧
        right_group = ttk.Frame(tool_frame)
        right_group.pack(side=tk.RIGHT)

        self.history_count_label = ttk.Label(right_group, text="", foreground="gray")
        self.history_count_label.pack(side=tk.LEFT, padx=(2, 0))

        ttk.Button(
            right_group, text="清空记录", command=self._clear_history, width=10
        ).pack(side=tk.LEFT, padx=(8, 0))

        # 列表
        columns = ("time", "task", "category", "status", "elapsed")
        self.history_tree = ttk.Treeview(
            parent, columns=columns, show="headings", height=15, selectmode=tk.BROWSE
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

        v_scroll = ttk.Scrollbar(parent, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=v_scroll.set)
        self.history_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 双击查看详情
        self.history_tree.bind('<Double-Button-1>', self._show_history_detail)

    # ---------- 刷新入口 ----------
    def _refresh_history(self):
        """刷新任务历史：统计面板 + 列表视图"""
        self._refresh_stats()
        self._refresh_history_list()

    def _refresh_stats(self):
        """计算并显示统计卡片：成功率 / 平均耗时 / 今日执行 / 历史总计"""
        records = self.history.records
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
        self.stat_avg.config(text=self.history.format_elapsed(avg))
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

    def _refresh_history_list(self):
        """刷新列表视图（应用状态 + 时间范围筛选）"""
        tree = self.history_tree
        tree.delete(*tree.get_children())
        total = len(self.history.records)
        filt = self.history_filter_var.get() if hasattr(self, "history_filter_var") else "全部"
        rng = self.history_range_var.get() if hasattr(self, "history_range_var") else "全部"

        status_tag = {
            STATUS_SUCCESS: "success",
            STATUS_FAILED: "failed",
            STATUS_ERROR: "error",
            STATUS_STOPPED: "stopped",
            STATUS_RUNNING: "running",
        }
        shown = 0
        for rec in self.history.records:
            status = rec.get("status", "")
            if filt != "全部" and status != filt:
                continue
            if rng != "全部" and not self._in_range(rec.get("time", ""), rng):
                continue
            shown += 1
            tag = status_tag.get(status, "")
            elapsed = self.history.format_elapsed(rec.get("elapsed", 0))
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

        parts = [f"共 {total} 条"]
        if filt != "全部" or rng != "全部":
            parts.append(f"显示 {shown} 条")
        self.history_count_label.config(text=" | ".join(parts))


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

    def _clear_history(self):
        """清空任务历史"""
        if not self.history.records:
            return
        if messagebox.askyesno("确认", "确定清空所有任务历史记录？"):
            self.history.clear()
            self._refresh_history()
            self._log("[历史] 已清空任务历史记录")

    def _show_history_detail(self, event):
        """双击列表查看任务详情"""
        sel = self.history_tree.selection()
        if not sel:
            return
        self._show_history_detail_by_id(int(sel[0]))

    def _show_history_detail_by_id(self, rec_id):
        """按记录 id 查看任务详情（列表 / 时间轴共用）"""
        rec = next((r for r in self.history.records if r["id"] == rec_id), None)
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
            f"耗时: {self.history.format_elapsed(rec.get('elapsed', 0))}"
        )
        if detail:
            msg += f"\n\n详情: {detail}"
        messagebox.showinfo("任务详情", msg)

    def _update_paths_for_selected_script(self):
        """根据当前选中的脚本更新路径显示"""
        script = self._get_selected_script()
        if not script:
            return
        output_dir = get_output_dir(self.user_config, self.current_category)
        filename = get_script_filename(script["name"])
        self.txt_path.set(os.path.join(output_dir, f"{filename}.txt"))
        self.xls_path.set(os.path.join(output_dir, f"{filename}.xls"))

    def _rebuild_params(self):
        """根据分类重建参数配置面板"""
        for w in self.params_frame.winfo_children():
            w.destroy()

        if self.current_category == "查询":
            self._build_query_params()
        elif self.current_category == "下单":
            self._build_order_params()
        elif self.current_category == "组合申报":
            self._build_combo_params()
        elif self.current_category == "交易系统设置":
            self._build_settings_params()

    def _update_params_for_selected_script(self):
        """根据选中的脚本更新参数面板"""
        if self.current_category not in ("下单", "组合申报"):
            return
        script = self._get_selected_script()
        if not script:
            return

        # 先清空参数配置面板
        for w in self.params_frame.winfo_children():
            w.destroy()

        if self.current_category == "下单":
            # 根据脚本决定是否显示数据预览
            if script["name"] in ("5.期权下单_一键导出", "6.全选撤单"):
                # 这两个脚本不需要数据预览
                self.preview_frame.pack_forget()
            else:
                # 其他脚本需要数据预览
                self.preview_frame.pack(fill=tk.X, pady=(0, 10))

            # 根据脚本更新参数配置面板
            if script["name"] == "6.全选撤单":
                # 全选撤单：显示提示信息
                ttk.Label(
                    self.params_frame,
                    text="该功能无需参数配置，点击执行即可。",
                    foreground="gray"
                ).pack(pady=20)
            elif script["name"] == "5.期权下单_一键导出":
                # 期权下单_一键导出：显示参数配置
                self._build_export_params()
            else:
                # 其他脚本：显示订单参数配置
                self._build_order_params()
        elif self.current_category == "组合申报":
            # 全自动脚本与查询脚本使用不同参数
            self._build_combo_params()

    def _build_query_params(self):
        """查询类参数"""
        ttk.Label(self.params_frame, text="导出格式:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Radiobutton(self.params_frame, text="TXT", variable=self.export_format, value="txt").grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(self.params_frame, text="XLS", variable=self.export_format, value="xls").grid(row=0, column=2, sticky=tk.W)

        ttk.Label(self.params_frame, text="自动打开文件:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(self.params_frame, text="导出后自动打开", variable=self.auto_open).grid(row=1, column=1, sticky=tk.W, columnspan=2)

        ttk.Label(self.params_frame, text="TXT输出路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        txt_entry = ttk.Entry(self.params_frame, textvariable=self.txt_path, width=35)
        txt_entry.grid(row=2, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=lambda: self._browse_path(self.txt_path, ".txt")).grid(row=2, column=2, padx=5)

        ttk.Label(self.params_frame, text="XLS输出路径:").grid(row=3, column=0, sticky=tk.W, pady=5)
        xls_entry = ttk.Entry(self.params_frame, textvariable=self.xls_path, width=35)
        xls_entry.grid(row=3, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=lambda: self._browse_path(self.xls_path, ".xls")).grid(row=3, column=2, padx=5)

        self.params_frame.columnconfigure(1, weight=1)

    def _build_order_params(self):
        """下单参数"""
        # 参数行
        param_row = ttk.Frame(self.params_frame)
        param_row.pack(fill=tk.X)

        ttk.Label(param_row, text="Excel文件:").grid(row=0, column=0, sticky=tk.W, pady=5)
        xlsx_entry = ttk.Entry(param_row, textvariable=self.xlsx_file, width=35)
        xlsx_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        ttk.Button(param_row, text="选择文件", command=self._browse_xlsx).grid(row=0, column=2, padx=5)
        param_row.columnconfigure(1, weight=1)

        self.params_frame.columnconfigure(0, weight=1)

    def _build_combo_params(self):
        """组合申报参数：根据选中的脚本显示不同参数

        - 全自动脚本(组合申报_全自动/拆分申报_全自动): 仅需委托数量
        - 查询类脚本(组合策略持仓查询等): 导出格式/路径/自动打开
        """
        script = self._get_selected_script()
        if script and script["name"] in self.COMBO_AUTO_SCRIPTS:
            self._build_combo_auto_params()
        else:
            self._build_combo_query_params()

    def _build_combo_auto_params(self):
        """组合申报/拆分申报 全自动：仅委托数量"""
        ttk.Label(self.params_frame, text="委托数量:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(self.params_frame, from_=1, to=999, textvariable=self.order_qty, width=10).grid(row=0, column=1, sticky=tk.W, pady=5)

        self.params_frame.columnconfigure(1, weight=1)

    def _build_combo_query_params(self):
        """组合策略查询类脚本：导出格式 / 路径 / 自动打开"""
        ttk.Label(self.params_frame, text="导出格式:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Radiobutton(self.params_frame, text="TXT", variable=self.export_format, value="txt").grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(self.params_frame, text="XLS", variable=self.export_format, value="xls").grid(row=0, column=2, sticky=tk.W)

        ttk.Label(self.params_frame, text="TXT输出路径:").grid(row=1, column=0, sticky=tk.W, pady=5)
        txt_entry = ttk.Entry(self.params_frame, textvariable=self.txt_path, width=35)
        txt_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=lambda: self._browse_path(self.txt_path, ".txt")).grid(row=1, column=2, padx=5)

        ttk.Label(self.params_frame, text="XLS输出路径:").grid(row=2, column=0, sticky=tk.W, pady=5)
        xls_entry = ttk.Entry(self.params_frame, textvariable=self.xls_path, width=35)
        xls_entry.grid(row=2, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=lambda: self._browse_path(self.xls_path, ".xls")).grid(row=2, column=2, padx=5)

        ttk.Label(self.params_frame, text="自动打开文件:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(self.params_frame, text="导出后自动打开", variable=self.auto_open).grid(row=3, column=1, sticky=tk.W, columnspan=2)

        self.params_frame.columnconfigure(1, weight=1)

    def _build_settings_params(self):
        """交易系统设置参数 - 输出路径可自定义"""
        ttk.Label(self.params_frame, text="输出路径:").grid(row=0, column=0, sticky=tk.W, pady=5)
        path_entry = ttk.Entry(self.params_frame, textvariable=self.settings_output_dir, width=35)
        path_entry.grid(row=0, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=self._browse_settings_dir).grid(row=0, column=2, padx=5)

        ttk.Label(
            self.params_frame,
            text="测试报告与截图将保存到该目录下（reports / screenshots 子目录）",
            foreground="gray"
        ).grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.params_frame.columnconfigure(1, weight=1)

    def _build_export_params(self):
        """期权下单_一键导出 参数配置"""
        # 清空旧内容
        for widget in self.params_frame.winfo_children():
            widget.destroy()

        # 导出目标选择
        ttk.Label(self.params_frame, text="导出目标:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Checkbutton(self.params_frame, text="持仓", variable=self.export_target_position).grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Checkbutton(self.params_frame, text="委托", variable=self.export_target_order).grid(row=0, column=2, sticky=tk.W, padx=5)

        # 输出目录配置
        ttk.Label(self.params_frame, text="输出目录:").grid(row=1, column=0, sticky=tk.W, pady=5)
        path_entry = ttk.Entry(self.params_frame, textvariable=self.export_output_dir, width=35)
        path_entry.grid(row=1, column=1, sticky=tk.EW, pady=5)
        ttk.Button(self.params_frame, text="浏览", command=self._browse_export_dir).grid(row=1, column=2, padx=5)

        # 显示文件名格式说明
        ttk.Label(self.params_frame, text="文件名格式: 期权下单(新)-持仓-20260629.xls", foreground="gray").grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=5)

        self.params_frame.columnconfigure(1, weight=1)

        # 确保按钮框架可见
        self.btn_frame.lift()

    def _browse_path(self, var, ext):
        """浏览选择文件路径"""
        initial_dir = get_output_dir(self.user_config, self.current_category)
        # 默认文件名：当前输入框已有值，或根据选中脚本生成
        current = var.get()
        if current and os.path.basename(current):
            initial_file = os.path.basename(current)
        else:
            script = self._get_selected_script()
            if script:
                initial_file = get_script_filename(script["name"]) + ext
            else:
                initial_file = "output" + ext

        path = filedialog.asksaveasfilename(
            initialdir=initial_dir,
            initialfile=initial_file,
            defaultextension=ext,
            filetypes=[("All files", "*.*"), (f"{ext.upper()} files", f"*{ext}")]
        )
        if path:
            var.set(path)
            new_dir = os.path.dirname(path)
            if new_dir != get_output_dir(self.user_config, self.current_category):
                set_output_dir(self.user_config, self.current_category, new_dir)
                save_user_config(self.user_config)
                self._log(f"[配置] 已更新{self.current_category}输出目录: {new_dir}")

    def _browse_xlsx(self):
        """选择Excel文件"""
        path = filedialog.askopenfilename(
            title="选择Excel配置文件",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if path:
            self.xlsx_file.set(path)
            self._log(f"已选择Excel文件: {path}")
            self._preview_excel(path)

    def _browse_export_dir(self):
        """选择导出目录"""
        path = filedialog.askdirectory(
            title="选择导出目录",
            initialdir=self.export_output_dir.get()
        )
        if path:
            self.export_output_dir.set(path)
            self._log(f"已设置导出目录: {path}")

    def _browse_settings_dir(self):
        """选择交易系统设置输出目录"""
        path = filedialog.askdirectory(
            title="选择输出目录",
            initialdir=self.settings_output_dir.get()
        )
        if path:
            self.settings_output_dir.set(path)
            set_output_dir(self.user_config, "交易系统设置", path)
            save_user_config(self.user_config)
            self._log(f"[配置] 已更新交易系统设置输出目录: {path}")

    def _preview_excel(self, filepath: str):
        """读取Excel并显示到预览表格"""
        if not HAS_OPENPYXL:
            self.preview_info.config(text="未安装 openpyxl，无法预览。请运行: pip install openpyxl", foreground="red")
            return

        # 清空旧数据
        self.preview_tree.delete(*self.preview_tree.get_children())
        self.preview_tree["columns"] = ()

        try:
            wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
            ws = wb.active

            # 使用 iter_rows 安全读取，避免直接访问 ws[1] 导致索引错误
            rows_iter = ws.iter_rows(values_only=True)

            # 读取第一行作为表头
            try:
                first_row = next(rows_iter)
            except StopIteration:
                self.preview_info.config(text="Excel文件无任何数据", foreground="red")
                wb.close()
                return

            if not first_row or all(v is None for v in first_row):
                self.preview_info.config(text="Excel文件无表头或为空", foreground="red")
                wb.close()
                return

            headers = [str(v) if v is not None else "" for v in first_row]

            if all(h == "" for h in headers):
                self.preview_info.config(text="Excel文件无表头或为空", foreground="red")
                wb.close()
                return

            # 设置列（紧凑宽度，不拉伸）
            self.preview_tree["columns"] = headers
            for h in headers:
                self.preview_tree.heading(h, text=h)
                self.preview_tree.column(h, width=64, minwidth=30, stretch=False)

            # 读取数据行（最多显示50行）
            row_count = 0
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row or all(v is None for v in row):
                    continue
                vals = [str(v) if v is not None else "" for v in row]
                # 补齐列数
                while len(vals) < len(headers):
                    vals.append("")
                self.preview_tree.insert("", tk.END, values=vals[:len(headers)])
                row_count += 1
                if row_count >= 50:
                    break

            total_rows = sum(1 for r in ws.iter_rows(min_row=2, values_only=True) if r and any(v is not None for v in r))
            show_note = f"（共{total_rows}行，仅显示前50行）" if total_rows > 50 else f"（共{total_rows}行）"
            self.preview_info.config(text=f"字段: {len(headers)} 个 | 数据: {show_note}", foreground="green")

            wb.close()

        except Exception as e:
            self.preview_info.config(text=f"读取失败: {e}", foreground="red")
            self._log(f"[预览] Excel读取失败: {e}")

    def _get_selected_script(self):
        """获取选中的脚本"""
        selection = self.script_listbox.curselection()
        if not selection:
            return None
        script_name = self.script_listbox.get(selection[0])
        for scripts in SCRIPTS_CONFIG.values():
            for s in scripts:
                if s["name"] == script_name:
                    return s
        return None

    # ====================== 脚本列表 -> 任务队列 拖拽 ======================
    def _on_list_drag_start(self, event):
        """从脚本列表开始拖拽（记录待拖出的脚本）"""
        if self.task_center is None or self.task_center.is_running:
            self._drag_script = None
            return
        idx = self.script_listbox.nearest(event.y)
        if idx < 0:
            self._drag_script = None
            return
        name = self.script_listbox.get(idx)
        script = None
        category = None
        for cat, scripts in SCRIPTS_CONFIG.items():
            for s in scripts:
                if s["name"] == name:
                    script, category = s, cat
                    break
            if script:
                break
        if script is None:
            self._drag_script = None
            return
        self._drag_script = dict(script)
        self._drag_script["category"] = category
        self._drag_active = False
        self._drag_start_y = event.y

    def _on_list_drag_motion(self, event):
        """拖动过程中：超过阈值视为拖拽，并在悬停于队列时显示落点"""
        if self._drag_script is None:
            return
        if abs(event.y - self._drag_start_y) < 6:
            return
        self._drag_active = True
        tc = self.task_center
        if tc is None:
            return
        tree = tc.tree
        rx, ry = tree.winfo_rootx(), tree.winfo_rooty()
        if rx <= event.x_root <= rx + tree.winfo_width() and ry <= event.y_root <= ry + tree.winfo_height():
            tc._update_drop_indicator(event.y_root - ry)
        else:
            tc._hide_drop_indicator()

    def _on_list_drag_end(self, event):
        """列表框内释放：若不在队列上方则取消拖拽状态"""
        tc = self.task_center
        over_tree = False
        if tc is not None:
            tree = tc.tree
            rx, ry = tree.winfo_rootx(), tree.winfo_rooty()
            over_tree = (rx <= event.x_root <= rx + tree.winfo_width()
                         and ry <= event.y_root <= ry + tree.winfo_height())
        if not over_tree:
            self._drag_script = None
            self._drag_active = False
            if tc is not None:
                tc._hide_drop_indicator()
        # 在队列上方释放时保留状态，交由 _on_global_drop 处理落点与清理

    def _on_global_drop(self, event):
        """全局捕获释放：处理从脚本列表拖入队列的落点"""
        script = self._drag_script
        active = self._drag_active
        self._drag_script = None
        self._drag_active = False
        tc = self.task_center
        if tc is not None:
            tc._hide_drop_indicator()
        if not active or script is None or tc is None:
            return
        if tc.is_running:
            return
        tree = tc.tree
        rx, ry = tree.winfo_rootx(), tree.winfo_rooty()
        if rx <= event.x_root <= rx + tree.winfo_width() and ry <= event.y_root <= ry + tree.winfo_height():
            tc.add_script_from_drop(script, event.y_root - ry)

    def _execute_script(self):
        """执行脚本"""
        if self.is_running:
            messagebox.showwarning("提示", "有脚本正在运行中，请先停止")
            return
        if self._task_mode:
            messagebox.showwarning("提示", "任务中心正在顺序执行中，请先停止")
            return

        script = self._get_selected_script()
        if not script:
            messagebox.showwarning("提示", "请先选择一个脚本")
            return

        if not os.path.exists(script["path"]):
            messagebox.showerror("错误", f"脚本文件不存在:\n{script['path']}")
            return

        # 下单需要检查Excel文件（全选撤单和期权下单_一键导出除外）
        if self.current_category == "下单" and script["name"] not in ("5.期权下单_一键导出", "6.全选撤单") and not self.xlsx_file.get():
            messagebox.showwarning("提示", "请先选择Excel配置文件")
            return

        # 期权下单_一键导出需要检查导出目标
        export_targets = []
        if script["name"] == "5.期权下单_一键导出":
            if self.export_target_position.get():
                export_targets.append("持仓")
            if self.export_target_order.get():
                export_targets.append("委托")
            if not export_targets:
                messagebox.showwarning("提示", "请至少选择一个导出目标（持仓或委托）")
                return

        # 保存当前配置
        self.user_config["export_format"] = self.export_format.get()
        self.user_config["auto_open"] = self.auto_open.get()
        if self.current_category == "交易系统设置":
            set_output_dir(self.user_config, "交易系统设置", self.settings_output_dir.get())
        save_user_config(self.user_config)

        # 确保路径已设置
        if self.current_category == "查询" and not self.txt_path.get():
            self._update_paths_for_selected_script()

        # 收集运行时参数，构造任务
        params = self.collect_params(export_targets)
        task = Task(script, self.current_category, params)

        self.is_running = True
        self.execute_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)

        self._set_status(f"正在运行: {script['name']}", running=True)

        self._log(f"\n{'='*60}")
        self._log(f"开始执行: {script['name']}")
        self._log(f"脚本路径: {script['path']}")

        # 打印参数
        if self.current_category == "查询":
            self._log(f"导出格式: {self.export_format.get().upper()}")
            self._log(f"自动打开: {'是' if self.auto_open.get() else '否'}")
            self._log(f"TXT路径: {self.txt_path.get()}")
            self._log(f"XLS路径: {self.xls_path.get()}")
        elif self.current_category == "下单":
            if script["name"] == "5.期权下单_一键导出":
                self._log(f"导出目标: {', '.join(export_targets)}")
                self._log(f"输出目录: {self.export_output_dir.get()}")
                self._log(f"文件名格式: 期权下单(新)-持仓-20260629.xls")
            else:
                self._log(f"Excel文件: {self.xlsx_file.get()}")
        elif self.current_category == "组合申报":
            if script["name"] in self.COMBO_AUTO_SCRIPTS:
                self._log(f"委托数量: {self.order_qty.get()}")
            else:
                self._log(f"导出格式: {self.export_format.get().upper()}")
                self._log(f"自动打开: {'是' if self.auto_open.get() else '否'}")
                self._log(f"TXT路径: {self.txt_path.get()}")
                self._log(f"XLS路径: {self.xls_path.get()}")
        elif self.current_category == "交易系统设置":
            self._log(f"输出路径: {self.settings_output_dir.get()}")

        self._log(f"{'='*60}")
        self.logger.info(f"开始执行: {script['name']}")

        # 交给执行引擎在后台线程运行
        self._current_record_id = self.history.add_record(script["name"], self.current_category)
        self._refresh_history()
        self.runner.run(task)

    def collect_params(self, export_targets=None):
        """收集当前界面参数，返回 dict（供执行/任务中心使用）"""
        return {
            "export_format": self.export_format.get(),
            "auto_open": self.auto_open.get(),
            "txt_path": self.txt_path.get(),
            "xls_path": self.xls_path.get(),
            "order_qty": self.order_qty.get(),
            "countdown_sec": self.countdown_sec.get(),
            "xlsx_file": self.xlsx_file.get(),
            "export_targets": export_targets or [],
            "export_output_dir": self.export_output_dir.get(),
            "settings_output_dir": self.settings_output_dir.get(),
        }

    # ====================== 任务中心：顺序执行驱动 ======================
    def run_task_center(self, task_center):
        """由任务中心调用：进入顺序执行模式并启动首个任务"""
        self._task_mode = True
        self.task_center = task_center
        # 禁用主窗口执行/停止按钮，避免与任务中心冲突
        self.execute_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.DISABLED)
        task_center.run_next()

    def stop_task_center(self):
        """由任务中心调用：停止当前正在执行的子进程"""
        if self.runner.is_running:
            self._log("\n[任务中心] 正在终止当前子进程...")
            self.logger.info("任务中心 - 用户手动停止")
            self.runner.stop()
        # 子进程结束后会触发回调，task_center 据此进入停止收尾流程

    def execute_task_item(self, item, record_id, next_category=""):
        """由任务中心调用：执行队列中的单个任务项（带参数快照）

        next_category: 下一个任务的分类，传给脚本用于决定交易系统设置窗口是否保留。
        """
        script = {"name": item["script_name"], "path": item["script_path"]}
        task = Task(script, item["category"], item["params"], next_category=next_category)
        self.runner.run(task)

    def _reset_running_state_if_idle(self):
        """任务中心收尾时复位主窗口运行状态（若非普通执行占用）"""
        self._task_mode = False
        self._reset_running_state()

    def _stop_script(self):
        """停止脚本"""
        if self.runner.is_running:
            self._log("\n[停止] 用户手动停止...")
            self.logger.info("用户手动停止")
            self._set_status("已停止（用户手动）")

            if self._current_record_id is not None:
                self.history.update_record(self._current_record_id, STATUS_STOPPED)
                self._current_record_id = None
                self._refresh_history()

            self.runner.stop()

    def _clear_log(self):
        """清空日志"""
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete(1.0, tk.END)
        self.log_text.config(state=tk.DISABLED)

    def _open_log_dir(self):
        """打开日志目录"""
        os.startfile(self.log_dir)

    # ====================== 执行结果回调（运行在 runner 线程，统一切回主线程更新 UI） ======================
    def _on_run_finish(self, return_code, elapsed, task):
        # 任务中心顺序执行模式：回调转交任务中心处理
        if self._task_mode:
            def _tc_finish():
                self.task_center.on_finish(return_code, elapsed, task)
            self.root.after(0, _tc_finish)
            return

        def _apply():
            if return_code == 0:
                status = STATUS_SUCCESS
                detail = ""
                self._log(f"\n[成功] {task.name} 执行完成")
                self.logger.info(f"执行成功: {task.name}")
                self._set_status(f"完成: {task.name} (用时 {elapsed:.1f}s)")
            else:
                status = STATUS_FAILED
                detail = f"退出码: {return_code}"
                self._log(f"\n[错误] {task.name} 执行失败，退出码: {return_code}")
                self.logger.error(f"执行失败: {task.name}, 退出码: {return_code}")
                self._set_status(f"失败: {task.name} (用时 {elapsed:.1f}s)")

            if self._current_record_id is not None:
                self.history.update_record(self._current_record_id, status, elapsed, detail)
                self._current_record_id = None
                self._refresh_history()

            self._reset_running_state()
        self.root.after(0, _apply)

    def _on_run_error(self, exc, task):
        # 任务中心顺序执行模式：回调转交任务中心处理
        if self._task_mode:
            def _tc_error():
                self.task_center.on_error(exc, task)
            self.root.after(0, _tc_error)
            return

        def _apply():
            self._log(f"\n[异常] 执行出错: {exc}")
            self.logger.error(f"执行异常: {exc}")
            self._set_status(f"异常: {task.name}")

            if self._current_record_id is not None:
                self.history.update_record(self._current_record_id, STATUS_ERROR, detail=str(exc))
                self._current_record_id = None
                self._refresh_history()

            self._reset_running_state()
        self.root.after(0, _apply)

    def _reset_running_state(self):
        """执行结束后复位运行状态与按钮"""
        self.is_running = False
        self.execute_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    # ====================== 状态栏 ======================
    def _set_status(self, text, running=False):
        """设置状态栏（线程安全，可在子线程调用）"""
        self.root.after(0, self._apply_status, text, running)

    def _apply_status(self, text, running):
        """在主线程更新状态栏"""
        self.status_label.config(text=text)
        if running:
            self._status_running = True
            self.task_start_time = time.time()
            self._tick_timer()
        else:
            self._status_running = False
            self.time_label.config(text="")

    def _tick_timer(self):
        """定时刷新运行时间"""
        if not self._status_running:
            return
        elapsed = time.time() - self.task_start_time
        self.time_label.config(text=f"运行时间: {elapsed:.1f}s")
        self.root.after(200, self._tick_timer)

    def _on_log_level_change(self):
        """切换日志级别并持久化"""
        self.user_config["log_level"] = self.log_level.get()
        save_user_config(self.user_config)
        self.level_label.config(text=f"日志: {self.log_level.get()}")
        self._log(f"日志级别已切换为: {self.log_level.get()}")
        self.logger.info(f"日志级别切换为: {self.log_level.get()}")

    def _show_log_level_help(self):
        """日志级别说明"""
        messagebox.showinfo(
            "日志级别说明",
            "详细日志：显示窗口/GUI 操作信息以及子脚本的 print 输出。\n\n"
            "简洁日志：仅显示窗口/GUI 的关键信息（开始、完成、错误等），"
            "不显示子脚本内部的 print 输出，界面更清爽，适合普通用户。\n\n"
            "当前级别可通过「设置」菜单随时切换，默认「详细日志」。"
        )

    def _log(self, message):
        """输出日志（带颜色区分，线程安全）"""
        self.root.after(0, self.log_text.append, message)
