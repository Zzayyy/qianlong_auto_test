# -*- coding: utf-8 -*-
"""GUI自动化主窗口：界面构建、参数配置、日志展示与执行编排"""

import os
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


class AutomationGUI:
    """GUI自动化主界面"""

    def __init__(self, root):
        self.root = root
        self.root.title("钱龙期权交易 - GUI自动化工具")
        self.root.geometry("950x750")
        self.root.minsize(850, 650)

        self.is_running = False
        self.current_category = "查询"

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

        # 右侧：日志
        log_frame = ttk.LabelFrame(self.paned, text="运行日志", padding="5")
        self.paned.add(log_frame, weight=1)

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

        # 默认显示下单（PanedWindow 用 weight=1 均分，无需手动设 sashpos）
        self._switch_category("下单")

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
        if self.current_category != "下单":
            return
        script = self._get_selected_script()
        if not script:
            return

        # 先清空参数配置面板
        for w in self.params_frame.winfo_children():
            w.destroy()

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
        """组合申报参数"""
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

        ttk.Separator(self.params_frame, orient=tk.HORIZONTAL).grid(row=4, column=0, columnspan=3, sticky=tk.EW, pady=10)

        ttk.Label(self.params_frame, text="委托数量:").grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(self.params_frame, from_=1, to=999, textvariable=self.order_qty, width=10).grid(row=5, column=1, sticky=tk.W, pady=5)

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

    def _execute_script(self):
        """执行脚本"""
        if self.is_running:
            messagebox.showwarning("提示", "有脚本正在运行中，请先停止")
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
        params = {
            "export_format": self.export_format.get(),
            "auto_open": self.auto_open.get(),
            "txt_path": self.txt_path.get(),
            "xls_path": self.xls_path.get(),
            "order_qty": self.order_qty.get(),
            "countdown_sec": self.countdown_sec.get(),
            "xlsx_file": self.xlsx_file.get(),
            "export_targets": export_targets,
            "export_output_dir": self.export_output_dir.get(),
            "settings_output_dir": self.settings_output_dir.get(),
        }
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
            self._log(f"导出格式: {self.export_format.get().upper()}")
            self._log(f"自动打开: {'是' if self.auto_open.get() else '否'}")
            self._log(f"TXT路径: {self.txt_path.get()}")
            self._log(f"XLS路径: {self.xls_path.get()}")
        elif self.current_category == "交易系统设置":
            self._log(f"输出路径: {self.settings_output_dir.get()}")

        self._log(f"{'='*60}")
        self.logger.info(f"开始执行: {script['name']}")

        # 交给执行引擎在后台线程运行
        self.runner.run(task)

    def _stop_script(self):
        """停止脚本"""
        if self.runner.is_running:
            self._log("\n[停止] 用户手动停止...")
            self.logger.info("用户手动停止")
            self._set_status("已停止（用户手动）")
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
        def _apply():
            if return_code == 0:
                self._log(f"\n[成功] {task.name} 执行完成")
                self.logger.info(f"执行成功: {task.name}")
                self._set_status(f"完成: {task.name} (用时 {elapsed:.1f}s)")
            else:
                self._log(f"\n[错误] {task.name} 执行失败，退出码: {return_code}")
                self.logger.error(f"执行失败: {task.name}, 退出码: {return_code}")
                self._set_status(f"失败: {task.name} (用时 {elapsed:.1f}s)")
            self._reset_running_state()
        self.root.after(0, _apply)

    def _on_run_error(self, exc, task):
        def _apply():
            self._log(f"\n[异常] 执行出错: {exc}")
            self.logger.error(f"执行异常: {exc}")
            self._set_status(f"异常: {task.name}")
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
