# -*- coding: utf-8 -*-
"""
定时任务面板 - 添加/编辑定时任务对话框
"""
import os as _os
import json, tkinter as tk
import datetime
from tkinter import ttk, messagebox, simpledialog
from config import get_scripts_config, CATEGORIES
from engine.scheduler import compute_next_run, format_schedule_desc

WEEKDAY_NAMES = ["周一","周二","周三","周四","周五","周六","周日"]


class AddSchedDialog(simpledialog.Dialog):
    """existing_data 不为 None 时进入编辑模式，预填所有字段"""
    def __init__(self, parent, groups, gui, client_id=None, title=None, existing_data=None):
        self.groups = groups
        self.gui = gui
        self.client_id = client_id
        self.existing_data = existing_data
        self.is_edit = existing_data is not None
        self.result_data = None
        self._wd_vars = {}
        self._card_frames = []
        self._all_scripts = []
        self._selected_idx = None
        super().__init__(parent, title=title or ("编辑定时任务" if self.is_edit else "添加定时任务"))

    def _center_on_parent(self):
        """居中显示"""
        try:
            self.update_idletasks()
            px = self.parent.winfo_rootx()
            py = self.parent.winfo_rooty()
            pw = self.parent.winfo_width()
            ph = self.parent.winfo_height()
            sw = self.winfo_width()
            sh = self.winfo_height()
            self.geometry(f"+{px+(pw-sw)//2}+{py+(ph-sh)//2}")
        except Exception:
            pass

    def body(self, master):
        # 固定弹窗尺寸
        if self.is_edit:
            self.geometry("450x320")
        else:
            self.geometry("580x600")
        self._center_on_parent()
        master.grid_columnconfigure(0, weight=1)

        # 编辑模式：预填值
        ed = self.existing_data or {}
        sched_type = ed.get("schedule_type", "daily")
        time_str = ed.get("time", "09:00")
        try:
            init_hour, init_min = time_str.split(":", 1)
        except ValueError:
            init_hour, init_min = "09", "00"
        init_cat = ed.get("category", "")
        init_group_name = ed.get("group_name", "")
        init_name = ed.get("name", "")
        init_date = ed.get("scheduled_date", datetime.datetime.now().strftime("%Y-%m-%d"))
        init_weekdays = ed.get("weekdays", [0,1,2,3,4])
        init_month_day = ed.get("month_day", 1)

        row = 0
        # === 任务信息 ===
        info_f = ttk.LabelFrame(master, text="任务信息", padding="8")
        info_f.grid(row=row, column=0, sticky=tk.NSEW, padx=12, pady=(8,3))
        info_f.grid_columnconfigure(0, weight=1)
        row += 1
        ttk.Label(info_f, text="任务名称:").pack(anchor=tk.W)
        self.name_var = tk.StringVar(value=init_name)
        self.name_entry = ttk.Entry(info_f, textvariable=self.name_var)
        self.name_entry.pack(fill=tk.X, pady=(2,0))

        # 编辑模式：显示只读的当前目标信息
        if self.is_edit and ed.get("target_type") == "script":
            info_line = ttk.Label(info_f, text=f"目标脚本: {ed.get('script_name','')}", foreground="gray")
            info_line.pack(anchor=tk.W, pady=(4,0))
        elif self.is_edit and ed.get("target_type") == "group":
            info_line = ttk.Label(info_f, text=f"目标编队: {ed.get('group_name','')}", foreground="gray")
            info_line.pack(anchor=tk.W, pady=(4,0))

        # === 目标设置（仅添加模式） ===
        if not self.is_edit:
            tg_f = ttk.LabelFrame(master, text="目标设置", padding="8")
            tg_f.grid(row=row, column=0, sticky=tk.EW, padx=12, pady=3)
            tg_f.grid_columnconfigure(1, weight=1)
            row += 1
            ttk.Label(tg_f, text="目标类型:").grid(row=0, column=0, sticky=tk.W)
            self.target_type = tk.StringVar(value="script")
            bf = ttk.Frame(tg_f)
            bf.grid(row=0, column=1, sticky=tk.W, padx=(8,0))
            ttk.Radiobutton(bf, text="单个脚本", variable=self.target_type, value="script", command=self._on_target_change).pack(side=tk.LEFT, padx=(0,10))
            ttk.Radiobutton(bf, text="任务编队", variable=self.target_type, value="group", command=self._on_target_change).pack(side=tk.LEFT)
            # 分类下拉
            self.cat_f = ttk.Frame(tg_f)
            self.cat_f.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6,0))
            self.cat_f.grid_columnconfigure(1, weight=1)
            ttk.Label(self.cat_f, text="分类:").grid(row=0, column=0, sticky=tk.W)
            self.cat_var = tk.StringVar(value=init_cat or (CATEGORIES[0] if CATEGORIES else ""))
            cats = [c for c in CATEGORIES if get_scripts_config(self.client_id).get(c)]
            self.cat_combo = ttk.Combobox(self.cat_f, textvariable=self.cat_var, values=cats, state="readonly", width=15)
            if cats and not init_cat:
                self.cat_var.set(cats[0])
            self.cat_combo.grid(row=0, column=1, sticky=tk.W, padx=(4,0))
            self.cat_combo.bind("<<ComboboxSelected>>", self._on_cat_change)
            # 编队选择
            self.grp_f = ttk.Frame(tg_f)
            self.grp_f.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6,0))
            self.grp_f.grid_columnconfigure(1, weight=1)
            ttk.Label(self.grp_f, text="编队:").grid(row=0, column=0, sticky=tk.W)
            gnames = [g["name"] for g in self.groups]
            self.group_var = tk.StringVar()
            self.group_combo = ttk.Combobox(self.grp_f, textvariable=self.group_var, values=gnames, state="readonly", width=25)
            self.group_combo.grid(row=0, column=1, sticky=tk.W, padx=(4,0))
            if gnames:
                if init_group_name and init_group_name in gnames:
                    self.group_combo.set(init_group_name)
                else:
                    self.group_combo.set(gnames[0])
            self.group_combo.bind("<<ComboboxSelected>>", self._on_group_change)
            self.grp_f.grid_remove()

        # === 脚本选择（仅添加模式） ===
        if not self.is_edit:
            self.sc_f = sc_f = ttk.LabelFrame(master, text="选择脚本", padding="8")
            sc_f.grid(row=row, column=0, sticky=tk.NSEW, padx=12, pady=3)
            row += 1
            sc_f.grid_columnconfigure(0, weight=1)
            self.sc_f.grid_rowconfigure(1, weight=1)
            cc = ttk.Frame(sc_f)
            cc.grid(row=1, column=0, sticky=tk.NSEW)
            cc.grid_rowconfigure(0, weight=1)
            cc.grid_columnconfigure(0, weight=1)
            self.card_canvas = tk.Canvas(cc, borderwidth=0, highlightthickness=0, height=100)
            cs = ttk.Scrollbar(cc, orient=tk.VERTICAL, command=self.card_canvas.yview)
            self.card_canvas.configure(yscrollcommand=cs.set)
            self.card_canvas.grid(row=0, column=0, sticky=tk.NSEW)
            cs.grid(row=0, column=1, sticky=tk.NS)
            self.card_inner = ttk.Frame(self.card_canvas)
            self.canvas_win = self.card_canvas.create_window((0,0), window=self.card_inner, anchor="nw")
            self.card_inner.bind("<Configure>", lambda e: self.card_canvas.configure(scrollregion=self.card_canvas.bbox("all")))
            self.card_canvas.bind("<Configure>", lambda e: self.card_canvas.itemconfig(self.canvas_win, width=e.width))
            self.card_canvas.bind("<Enter>", lambda e: self.card_canvas.bind_all("<MouseWheel>", self._on_card_scroll))
            self.card_canvas.bind("<Leave>", lambda e: self.card_canvas.unbind_all("<MouseWheel>"))
            self.sched_type = tk.StringVar(value="daily")
            self._fill_scripts()
        else:
            self.sched_type = tk.StringVar(value=sched_type)

        # === 定时配置 ===
        sch_f = ttk.LabelFrame(master, text="定时配置", padding="8")
        sch_f.grid(row=row, column=0, sticky=tk.NSEW, padx=12, pady=3)
        sch_f.grid_columnconfigure(1, weight=1)
        row += 1
        ttk.Label(sch_f, text="方式:").grid(row=0, column=0, sticky=tk.W, pady=(6,0))
        sf = ttk.Frame(sch_f)
        sf.grid(row=0, column=1, sticky=tk.W, padx=(8,0), pady=(6,0))
        for k,v in [("once","一次"),("daily","每天"),("weekly","每周"),("monthly","每月")]:
            rb = ttk.Radiobutton(sf, text=v, variable=self.sched_type, value=k, command=self._on_sched_change, style="Toolbutton")
            rb.pack(side=tk.LEFT, padx=(0,4))
        ttk.Label(sch_f, text="时间:").grid(row=1, column=0, sticky=tk.W, pady=(5,0))
        tf2 = ttk.Frame(sch_f)
        tf2.grid(row=1, column=1, sticky=tk.W, padx=(8,0), pady=(5,0))
        self.hour_var = tk.StringVar(value=init_hour)
        self.min_var = tk.StringVar(value=init_min)
        ttk.Spinbox(tf2, from_=0, to=23, textvariable=self.hour_var, width=4, format="%02.0f").pack(side=tk.LEFT)
        ttk.Label(tf2, text=":").pack(side=tk.LEFT)
        ttk.Spinbox(tf2, from_=0, to=59, textvariable=self.min_var, width=4, format="%02.0f").pack(side=tk.LEFT)
        # 日期选择
        self.df = ttk.Frame(sch_f)
        self.df.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        ttk.Label(self.df, text="执行日期:").pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=init_date)
        ttk.Entry(self.df, textvariable=self.date_var, width=12).pack(side=tk.LEFT, padx=(4,0))
        self.df.grid_remove()
        # 星期多选
        self.wf = ttk.Frame(sch_f)
        self.wf.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        for i,wn in enumerate(["周一","周二","周三","周四","周五","周六","周日"]):
            self._wd_vars[i] = tk.BooleanVar(value=(i in init_weekdays))
            ttk.Checkbutton(self.wf, text=wn, variable=self._wd_vars[i]).pack(side=tk.LEFT, padx=2)
        self.wf.grid_remove()
        # 每月第几天
        self.mf = ttk.Frame(sch_f)
        self.mf.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        ttk.Label(self.mf, text="每月第").pack(side=tk.LEFT)
        self.md_var = tk.IntVar(value=init_month_day)
        ttk.Spinbox(self.mf, from_=1, to=28, textvariable=self.md_var, width=5).pack(side=tk.LEFT, padx=4)
        ttk.Label(self.mf, text="天").pack(side=tk.LEFT)
        self.mf.grid_remove()
        # 编辑模式：初始化定时配置显示状态
        if self.is_edit:
            self.sched_type.set(sched_type)
            self._on_sched_change()

        # 行权重：对应区块拉伸填满弹窗
        if self.is_edit:
            master.grid_rowconfigure(1, weight=1)
        else:
            master.grid_rowconfigure(2, weight=1)

        return self.name_entry

    def buttonbox(self):
        """底部按钮：右对齐，确定在左 取消在右"""
        box = ttk.Frame(self)
        # 弹性空白将按钮推到右侧
        ttk.Label(box, text="").pack(side=tk.LEFT, fill=tk.X, expand=True)
        # 先 pack 取消（最右），再 pack 确定（取消左边）
        # 最终排列：... [确定] [取消]
        ttk.Button(box, text="取消", width=10, command=self.cancel).pack(side=tk.RIGHT, padx=(0, 12))
        ttk.Button(box, text="确定", width=10, command=self.ok).pack(side=tk.RIGHT)
        box.pack(fill=tk.X, padx=0, pady=(8, 12))
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)

    def _on_card_scroll(self, event):
        self.card_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _on_target_change(self):
        if self.target_type.get() == "script":
            self.cat_f.grid()
            self.grp_f.grid_remove()
            self.sc_f.configure(text="选择脚本")
            self._fill_scripts()
        else:
            self.cat_f.grid_remove()
            self.grp_f.grid()
            self.sc_f.configure(text="")
            self._fill_group_scripts()

    def _on_cat_change(self, event=None):
        self._fill_scripts()

    def _on_group_change(self, event=None):
        if self.target_type.get() == "group":
            self._fill_group_scripts()

    def _fill_group_scripts(self):
        """用选中编队的任务填充卡片网格"""
        for frm, _ in self._card_frames:
            try: frm.destroy()
            except: pass
        self._card_frames = []
        self._all_scripts = []
        self._selected_idx = -1
        gn = self.group_var.get()
        if not gn:
            self._filter_scripts()
            return
        g = next((g for g in self.groups if g["name"] == gn), None)
        if not g:
            self._filter_scripts()
            return
        tasks = g.get("tasks", [])
        self._all_scripts = tasks
        for idx, t in enumerate(tasks):
            card = tk.Frame(self.card_inner, relief="solid", borderwidth=1, bg="white")
            r, c = divmod(idx, 2)
            card.grid(row=r, column=c, sticky=tk.EW, padx=3, pady=2)
            self.card_inner.grid_columnconfigure(0, weight=1)
            self.card_inner.grid_columnconfigure(1, weight=1)
            name = t.get("script_name", "") or "未知脚本"
            cat = t.get("category", "") or ""
            ttk.Label(card, text=f"#{idx+1}", foreground="#0078d4", font=("", 8)).pack(anchor=tk.W)
            ttk.Label(card, text=name, font=("", 9)).pack(anchor=tk.W)
            if cat:
                ttk.Label(card, text=cat, foreground="gray", font=("", 7)).pack(anchor=tk.W)
            self._card_frames.append((card, t))
        self._filter_scripts()

    def _filter_scripts(self):
        """显示全部卡片并调整组件高度"""
        visible = len(self._card_frames)
        for frm, _ in self._card_frames:
            try: frm.grid()
            except: pass
        self.card_canvas.configure(height=min(180, max(80, max(2, (visible+1)//2)*50)))

    def _fill_scripts(self):
        for frm, _ in self._card_frames:
            frm.destroy()
        self._card_frames = []
        self._all_scripts = []
        self._selected_idx = None
        cat = self.cat_var.get()
        self._all_scripts = list(get_scripts_config(self.client_id).get(cat, []))
        for idx, s in enumerate(self._all_scripts):
            card = tk.Frame(self.card_inner, relief="solid", borderwidth=1, bg="white")
            r, c = divmod(idx, 2)
            card.grid(row=r, column=c, sticky=tk.EW, padx=3, pady=2)
            self.card_inner.grid_columnconfigure(0, weight=1)
            self.card_inner.grid_columnconfigure(1, weight=1)
            num = s["name"].split(".")[0] if "." in s["name"] else str(idx+1)
            ttk.Label(card, text=f"#{num}", foreground="#1d3cc7", font=("", 8)).pack(anchor=tk.W)
            ttk.Label(card, text=s["name"].split(".",1)[-1].strip() if "." in s["name"] else s["name"], font=("", 9)).pack(anchor=tk.W)
            card.bind("<Button-1>", lambda e, i=idx: self._on_card_click(i))
            for child in card.winfo_children():
                child.bind("<Button-1>", lambda e, i=idx: self._on_card_click(i))
            self._card_frames.append((card, s))
        self._filter_scripts()

    def _on_card_click(self, idx):
        """点击卡片时高亮选中状态（蓝色背景+加粗边框）"""
        for i, (frm, _) in enumerate(self._card_frames):
            if i == idx:
                frm.configure(bg="#dbeafe", borderwidth=2)
                for child in frm.winfo_children():
                    try: child.configure(bg="#dbeafe")
                    except: pass
            else:
                frm.configure(bg="white", borderwidth=1)
                for child in frm.winfo_children():
                    try: child.configure(bg="white")
                    except: pass
        self._selected_idx = idx if idx != self._selected_idx else None

    def _on_sched_change(self):
        st = self.sched_type.get()
        if st == "once": self.df.grid()
        else: self.df.grid_remove()
        if st == "weekly": self.wf.grid()
        else: self.wf.grid_remove()
        if st == "monthly": self.mf.grid()
        else: self.mf.grid_remove()

    def validate(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入任务名称")
            return False
        import datetime as _dt
        st = self.sched_type.get()
        if st == "once":
            try:
                _dt.datetime.strptime(self.date_var.get(), "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("提示", "日期格式错误")
                return False
        if st == "weekly" and not any(v.get() for v in self._wd_vars.values()):
            messagebox.showwarning("提示", "请至少选择一天")
            return False
        self.result_data = {
            "name": name,
            "time": str(max(0, min(23, int(self.hour_var.get() or 0)))).zfill(2) + ":" + str(max(0, min(59, int(self.min_var.get() or 0)))).zfill(2),
            "schedule_type": st,
            "scheduled_date": self.date_var.get() if st == "once" else "",
            "weekdays": sorted([k for k,v in self._wd_vars.items() if v.get()]) if st == "weekly" else [],
            "month_day": max(1, min(28, self.md_var.get())) if st == "monthly" else 1,
        }
        if self.is_edit:
            # 编辑模式：复用原有目标设置
            ed = self.existing_data
            self.result_data["target_type"] = ed.get("target_type", "script")
            if ed.get("target_type") == "script":
                self.result_data["script_name"] = ed.get("script_name", "")
                self.result_data["script_path"] = ed.get("script_path", "")
                self.result_data["query_key"] = ed.get("query_key", "")
                self.result_data["category"] = ed.get("category", "")
                self.result_data["params"] = ed.get("params", {})
            else:
                self.result_data["group_name"] = ed.get("group_name", "")
        elif self.target_type.get() == "script":
            if self._selected_idx is None:
                messagebox.showwarning("提示", "请选择一个脚本")
                return False
            if self._selected_idx >= len(self._all_scripts):
                messagebox.showwarning("提示", "脚本索引无效，请重新选择")
                return False
            s = self._all_scripts[self._selected_idx]
            self.result_data["target_type"] = "script"
            self.result_data["script_name"] = s["name"]
            self.result_data["script_path"] = s["path"]
            self.result_data["query_key"] = s.get("query_key", "")
            self.result_data["category"] = self.cat_var.get()
            self.result_data["params"] = self._snapshot_params()
        else:
            gn = self.group_var.get()
            if not gn:
                messagebox.showwarning("提示", "请选择一个任务编队")
                return False
            self.result_data["target_type"] = "group"
            self.result_data["group_name"] = gn
        return True

    def _snapshot_params(self):
        if hasattr(self, "gui") and hasattr(self.gui, "collect_params"):
            return self.gui.collect_params()
        return {}


class SchedulerPanel:
    def __init__(self, parent, controller):
        self.parent = parent
        self.controller = controller
        self.gui = controller.gui
        self.scheduler = None
        self._build_ui()

    def _build_ui(self):
        tool_frame = ttk.Frame(self.parent)
        tool_frame.pack(side=tk.TOP, fill=tk.X, pady=(0,5))
        self.count_label = ttk.Label(tool_frame, text="共 0 个任务", foreground="gray")
        self.count_label.pack(side=tk.LEFT)
        ttk.Button(tool_frame, text="添加定时任务", command=self._on_add, width=12).pack(side=tk.RIGHT, padx=(2,0))
        ttk.Button(tool_frame, text="刷新", command=self._on_refresh, width=8).pack(side=tk.RIGHT, padx=(2,0))
        # —— 列表区域 ——
        list_frame = ttk.Frame(self.parent)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)
        cols = ("id","name","target","schedule","next_run","last_run","status","enabled")
        self.tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=15, selectmode=tk.BROWSE)
        col_defs = [("id","ID"),("name","任务名称"),("target","目标"),("schedule","定时方式"),
                     ("next_run","下次执行"),("last_run","上次执行"),("status","状态"),("enabled","启用")]
        for c, h in col_defs:
            self.tree.heading(c, text=h)
        self.tree.column("id", width=28, minwidth=26, stretch=False, anchor=tk.CENTER)
        self.tree.column("name", width=60, minwidth=40, stretch=False)
        self.tree.column("target", width=55, minwidth=40, stretch=True)
        self.tree.column("schedule", width=55, minwidth=40, stretch=True)
        self.tree.column("next_run", width=75, minwidth=65, stretch=True)
        self.tree.column("last_run", width=75, minwidth=65, stretch=True)
        self.tree.column("status", width=46, minwidth=38, stretch=True, anchor=tk.CENTER)
        self.tree.column("enabled", width=35, minwidth=30, stretch=True, anchor=tk.CENTER)
        self.tree.tag_configure("running", foreground="#0000FF")
        self.tree.tag_configure("success", foreground="#008000")
        self.tree.tag_configure("disabled", foreground="#808080")
        v_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        # 双击编辑
        self.tree.bind("<Double-1>", self._on_double_click)
        # 右键菜单
        self._context_menu = tk.Menu(self.tree, tearoff=0)
        self._context_menu.add_command(label="编辑", command=self._on_edit)
        self._context_menu.add_separator()
        self._context_menu.add_command(label="删除", command=self._on_delete)
        self.tree.bind("<Button-3>", self._on_context_menu)
        # —— 底部按钮 ——
        btn_frame = ttk.Frame(self.parent)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(5,0))
        ttk.Button(btn_frame, text="启用/禁用", command=self._on_toggle, width=10).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="编辑", command=self._on_edit, width=8).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self._on_delete, width=8).pack(side=tk.LEFT, padx=2)

    def bind_scheduler(self, scheduler):
        self.scheduler = scheduler
        scheduler.bind_view(self)
        self._refresh()

    def _refresh(self):
        if not self.scheduler: return
        self.tree.delete(*self.tree.get_children())
        for idx, t in enumerate(self.scheduler.tasks, 1):
            tid = t.get("id",0); en = t.get("enabled",False); st = t.get("status","")
            tag = "running" if st == "运行中" else ("success" if st in ("已完成","等待") and en else ("disabled" if not en else ""))
            target = t.get("script_name","") or t.get("group_name","")
            if t.get("target_type") == "group": target = "[编队] " + target
            sd = format_schedule_desc(t); nr = t.get("next_run","")
            lr = t.get("last_run","") or "-"; et = "是" if en else "否"
            self.tree.insert("",tk.END,iid=str(tid),values=(idx,t.get("name",""),target,sd,nr,lr,st,et),tags=(tag,) if tag else ())
        self.count_label.config(text=f"共 {len(self.scheduler.tasks)} 个任务")

    def _on_add(self):
        if not self.scheduler: return
        groups = self._load_groups()
        dlg = AddSchedDialog(self.parent, groups, gui=self.gui, client_id=self.gui.client_id, title="添加定时任务")
        if dlg.result_data:
            self.scheduler.add_task(dlg.result_data)
            self._refresh()
            self.gui._log("[定时任务] 已添加: " + dlg.result_data["name"])

    def _load_groups(self):
        """加载任务编队列表"""
        groups = []
        gp = _os.path.normpath(_os.path.join(self.gui.log_dir, "task_groups.json"))
        if _os.path.exists(gp):
            try:
                with open(gp, "r", encoding="utf-8") as f:
                    data = json.load(f)
                groups = data if isinstance(data, list) else data.get("groups", [])
            except:
                groups = []
        return groups

    def _on_double_click(self, event):
        """双击行直接打开编辑对话框"""
        self._on_edit()

    def _on_context_menu(self, event):
        """右键弹出菜单"""
        iid = self.tree.identify_row(event.y)
        if iid:
            self.tree.selection_set(iid)
            self._context_menu.post(event.x_root, event.y_root)

    def _on_edit(self):
        """打开完整编辑对话框，支持修改名称、定时方式、时间、星期/月日等"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "请先选择一个定时任务")
            return
        tid = int(sel[0])
        t = self.scheduler.get_task(tid)
        if not t or not t.get("id"):
            messagebox.showwarning("提示", "未找到该任务数据")
            return
        groups = self._load_groups()
        dlg = AddSchedDialog(self.parent, groups, gui=self.gui,
                             client_id=self.gui.client_id,
                             title="编辑定时任务", existing_data=t)
        if dlg.result_data:
            updates = {
                "name": dlg.result_data["name"],
                "time": dlg.result_data["time"],
                "schedule_type": dlg.result_data["schedule_type"],
                "scheduled_date": dlg.result_data.get("scheduled_date", ""),
                "weekdays": dlg.result_data.get("weekdays", []),
                "month_day": dlg.result_data.get("month_day", 1),
            }
            self.scheduler.update_task(tid, updates)
            self._refresh()
            self.gui._log("[定时任务] 已更新: " + dlg.result_data["name"])

    def _on_delete(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("提示","请先选择一个定时任务"); return
        tid = int(sel[0]); t = self.scheduler.get_task(tid)
        if not t or not messagebox.askyesno("确认","确定删除定时任务「"+t.get("name","")+"」？"): return
        self.scheduler.delete_task(tid); self._refresh()
        self.gui._log("[定时任务] 已删除: "+t.get("name",""))

    def _on_toggle(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("提示","请先选择一个定时任务"); return
        tid = int(sel[0]); en = self.scheduler.toggle_enabled(tid)
        self._refresh(); t = self.scheduler.get_task(tid)
        self.gui._log("[定时任务] "+("启用" if en else "禁用")+": "+t.get("name",""))

    def _on_refresh(self):
        self._refresh(); self.gui._log("[定时任务] 已刷新")