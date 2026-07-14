# -*- coding: utf-8 -*-
"""
定时任务面板
"""

import os as _os
import json, tkinter as tk
import datetime
from tkinter import ttk, messagebox, simpledialog
from config import SCRIPTS_CONFIG, CATEGORIES
from engine.scheduler import compute_next_run, format_schedule_desc


class AddSchedDialog(simpledialog.Dialog):
    def __init__(self, parent, groups, title=None):
        self.groups = groups
        self.result_data = None
        self._wd_vars = {}
        self._card_frames = []
        self._all_scripts = []
        self._selected_idx = None
        self._grp_display = []
        super().__init__(parent, title=title or "添加定时任务")
        self.after(10, self._center_on_parent)

    def _center_on_parent(self):
        try:
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
        self.geometry("600x660")
        master.grid_columnconfigure(0, weight=1)
        master.grid_rowconfigure(3, weight=1)
        row = 0

        # === 任务信息 ===
        info_f = ttk.LabelFrame(master, text="任务信息", padding="8")
        info_f.grid(row=row, column=0, sticky=tk.EW, padx=8, pady=4)
        info_f.grid_columnconfigure(0, weight=1)
        row += 1
        ttk.Label(info_f, text="任务名称:").pack(anchor=tk.W)
        self.name_var = tk.StringVar(value="")
        self.name_entry = ttk.Entry(info_f, textvariable=self.name_var)
        self.name_entry.pack(fill=tk.X, pady=(2,0))

        # === 目标设置 ===
        tg_f = ttk.LabelFrame(master, text="目标设置", padding="8")
        tg_f.grid(row=row, column=0, sticky=tk.EW, padx=8, pady=4)
        tg_f.grid_columnconfigure(1, weight=1)
        row += 1
        ttk.Label(tg_f, text="目标类型:").grid(row=0, column=0, sticky=tk.W)
        self.target_type = tk.StringVar(value="script")
        bf = ttk.Frame(tg_f)
        bf.grid(row=0, column=1, sticky=tk.W, padx=(8,0))
        ttk.Radiobutton(bf, text="单个脚本", variable=self.target_type, value="script", command=self._on_target_change).pack(side=tk.LEFT, padx=(0,10))
        ttk.Radiobutton(bf, text="任务编队", variable=self.target_type, value="group", command=self._on_target_change).pack(side=tk.LEFT)

        # 分类下拉（脚本模式）
        self.cat_f = ttk.Frame(tg_f)
        self.cat_f.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6,0))
        self.cat_f.grid_columnconfigure(1, weight=1)
        ttk.Label(self.cat_f, text="分类:").grid(row=0, column=0, sticky=tk.W)
        self.cat_var = tk.StringVar(value=CATEGORIES[0] if CATEGORIES else "")
        self.cat_combo = ttk.Combobox(self.cat_f, textvariable=self.cat_var, values=CATEGORIES, state="readonly", width=15)
        self.cat_combo.grid(row=0, column=1, sticky=tk.W, padx=(4,0))
        self.cat_combo.bind("<<ComboboxSelected>>", self._on_cat_change)

        # 编队选择（编队模式）
        self.grp_f = ttk.Frame(tg_f)
        self.grp_f.grid(row=1, column=0, columnspan=2, sticky=tk.EW, pady=(6,0))
        self.grp_f.grid_columnconfigure(1, weight=1)
        ttk.Label(self.grp_f, text="编队:").grid(row=0, column=0, sticky=tk.W)
        gnames = [g["name"] for g in self.groups]
        self.group_var = tk.StringVar()
        self.group_combo = ttk.Combobox(self.grp_f, textvariable=self.group_var, values=gnames, state="readonly", width=25)
        self.group_combo.grid(row=0, column=1, sticky=tk.W, padx=(4,0))
        if gnames: self.group_combo.set(gnames[0])
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_change)
        self.grp_f.grid_remove()

        # === 脚本选择（卡片网格） ===
        self.sc_f = sc_f = ttk.LabelFrame(master, text="选择脚本", padding="8")
        sc_f.grid(row=row, column=0, sticky=tk.NSEW, padx=8, pady=4)
        sc_f.grid_columnconfigure(0, weight=1)
        self.sc_f.grid_rowconfigure(1, weight=1)
        row += 1

        # 搜索框
        self.search_var = tk.StringVar(value="")
        sv = self.search_var
        self._search_trace = sv.trace("w", lambda *a: self._filter_scripts())
        self.search_entry = ttk.Entry(sc_f, textvariable=self.search_var)
        self.search_entry.grid(row=0, column=0, sticky=tk.EW, pady=(0,4))

        # 卡片容器 + 滚动
        cc = ttk.Frame(sc_f)
        cc.grid(row=1, column=0, sticky=tk.NSEW)
        cc.grid_rowconfigure(0, weight=1)
        cc.grid_columnconfigure(0, weight=1)
        self.card_canvas = tk.Canvas(cc, borderwidth=0, highlightthickness=0, height=140)
        cs = ttk.Scrollbar(cc, orient=tk.VERTICAL, command=self.card_canvas.yview)
        self.card_canvas.configure(yscrollcommand=cs.set)
        self.card_canvas.grid(row=0, column=0, sticky=tk.NSEW)
        cs.grid(row=0, column=1, sticky=tk.NS)
        self.card_inner = ttk.Frame(self.card_canvas)
        self.canvas_win = self.card_canvas.create_window((0,0), window=self.card_inner, anchor="nw")
        self.card_inner.bind("<Configure>", lambda e: self.card_canvas.configure(scrollregion=self.card_canvas.bbox("all")))
        self.card_canvas.bind("<Configure>", lambda e: self.card_canvas.itemconfig(self.canvas_win, width=e.width))
        # 鼠标滚轮支持
        self.card_canvas.bind("<Enter>", lambda e: self.card_canvas.bind_all("<MouseWheel>", self._on_card_scroll))
        self.card_canvas.bind("<Leave>", lambda e: self.card_canvas.unbind_all("<MouseWheel>"))

        self.sched_type = tk.StringVar(value="daily")
        self._fill_scripts()

        # === 定时配置 ===
        sch_f = ttk.LabelFrame(master, text="定时配置", padding="8")
        sch_f.grid(row=row, column=0, sticky=tk.EW, padx=8, pady=4)
        sch_f.grid_columnconfigure(1, weight=1)
        row += 1
        ttk.Label(sch_f, text="方式:").grid(row=0, column=0, sticky=tk.W)
        sf = ttk.Frame(sch_f)
        sf.grid(row=0, column=1, sticky=tk.W, padx=(8,0))
        for k,v in [("once","一次"),("daily","每天"),("weekly","每周"),("monthly","每月")]:
            rb = ttk.Radiobutton(sf, text=v, variable=self.sched_type, value=k, command=self._on_sched_change, style="Toolbutton")
            rb.pack(side=tk.LEFT, padx=(0,4))

        ttk.Label(sch_f, text="时间:").grid(row=1, column=0, sticky=tk.W, pady=(6,0))
        tf2 = ttk.Frame(sch_f)
        tf2.grid(row=1, column=1, sticky=tk.W, padx=(8,0), pady=(6,0))
        self.hour_var = tk.StringVar(value="09")
        self.min_var = tk.StringVar(value="00")
        ttk.Spinbox(tf2, from_=0, to=23, textvariable=self.hour_var, width=4, format="%02.0f").pack(side=tk.LEFT)
        ttk.Label(tf2, text=":").pack(side=tk.LEFT)
        ttk.Spinbox(tf2, from_=0, to=59, textvariable=self.min_var, width=4, format="%02.0f").pack(side=tk.LEFT)

        # 日期选择
        self.df = ttk.Frame(sch_f)
        self.df.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        ttk.Label(self.df, text="执行日期:").pack(side=tk.LEFT)
        today = datetime.datetime.now().strftime("%Y-%m-%d")
        self.date_var = tk.StringVar(value=today)
        ttk.Entry(self.df, textvariable=self.date_var, width=12).pack(side=tk.LEFT, padx=(4,0))
        self.df.grid_remove()

        # 星期多选
        self.wf = ttk.Frame(sch_f)
        self.wf.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        for i,wn in enumerate(["周一","周二","周三","周四","周五","周六","周日"]):
            self._wd_vars[i] = tk.BooleanVar(value=(i<5))
            ttk.Checkbutton(self.wf, text=wn, variable=self._wd_vars[i]).pack(side=tk.LEFT, padx=2)
        self.wf.grid_remove()

        # 每月第几天
        self.mf = ttk.Frame(sch_f)
        self.mf.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(4,0))
        ttk.Label(self.mf, text="每月第").pack(side=tk.LEFT)
        self.md_var = tk.IntVar(value=1)
        ttk.Spinbox(self.mf, from_=1, to=28, textvariable=self.md_var, width=5).pack(side=tk.LEFT, padx=4)
        ttk.Label(self.mf, text="天").pack(side=tk.LEFT)
        self.mf.grid_remove()

        self.name_var.trace("w", lambda *a: self._update_name_preview())
        self.sched_type.trace("w", lambda *a: self._update_name_preview())
        self._update_name_preview()

        return self.name_entry


    def buttonbox(self):
        """右对齐的确定/取消按钮"""
        box = ttk.Frame(self)
        ok_btn = ttk.Button(box, text="确定", width=10, command=self.ok)
        ok_btn.pack(side=tk.RIGHT, padx=(0, 0))
        ttk.Button(box, text="取消", width=10, command=self.cancel).pack(side=tk.RIGHT, padx=(12, 0))
        box.pack(fill=tk.X, padx=12, pady=(0, 12))
        self.bind("<Return>", self.ok)
        self.bind("<Escape>", self.cancel)
    def _on_card_scroll(self, event):
        self.card_canvas.yview_scroll(int(-1*(event.delta/120)), "units")

    def _update_name_preview(self):
        prefix = {"once":"单次","daily":"每日","weekly":"每周","monthly":"每月"}
        p = prefix.get(self.sched_type.get(), "")
        sel_name = ""
        if self._selected_idx is not None and self._selected_idx < len(self._all_scripts):
            sel_name = self._all_scripts[self._selected_idx].get("name", "") or self._all_scripts[self._selected_idx].get("script_name", "")
        cur = self.name_var.get().strip()
        gen = (p + sel_name) if (not cur) else ""

    def _on_target_change(self):
        if self.target_type.get() == "script":
            self.cat_f.grid()
            self.grp_f.grid_remove()
            self.sc_f.configure(text="选择脚本")
            self.search_entry.grid()
            self._fill_scripts()
        else:
            self.cat_f.grid_remove()
            self.grp_f.grid()
            self.sc_f.configure(text="")
            self.search_entry.grid_remove()
            self.search_var.set("")
            self._fill_group_scripts()
    def _grp_tree_remove(self):
        """清除编队详情显示"""
        for w in getattr(self, "_grp_display", []) or []:
            try: w.destroy()
            except Exception: pass
        self._grp_display = []
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
        # 编队模式不更新名称预览
            return
        g = next((g for g in self.groups if g["name"] == gn), None)
        if not g:
            self._filter_scripts()
        # 编队模式不更新名称预览
            return

        tasks = g.get("tasks", [])
        self._all_scripts = tasks
        for idx, t in enumerate(tasks):
            card = tk.Frame(self.card_inner, relief="solid", borderwidth=1, bg="white")
            r, c = divmod(idx, 2)
            card.grid(row=r, column=c, sticky=tk.EW, padx=3, pady=3)
            self.card_inner.grid_columnconfigure(0, weight=1)
            self.card_inner.grid_columnconfigure(1, weight=1)
            name = t.get("script_name", "") or "未知脚本"
            cat = t.get("category", "") or ""
            ttk.Label(card, text=f"#{idx+1}", foreground="#0078d4", font=("", 8)).pack(anchor=tk.W)
            ttk.Label(card, text=name, font=("", 9)).pack(anchor=tk.W)
            if cat:
                ttk.Label(card, text=cat, foreground="gray", font=("", 7)).pack(anchor=tk.W)
            # 只展示，不可选择
            self._card_frames.append((card, t))
        self._filter_scripts()
        # 编队模式不更新名称预览
    def _filter_scripts(self):
        """根据搜索关键词显隔卡片"""
        kw = self.search_var.get().strip().lower()
        visible = 0
        for i, (frm, script) in enumerate(self._card_frames):
            match = (not kw) or (kw in (script.get("name") or script.get("script_name", "")).lower())
            frm.grid() if match else frm.grid_remove()
            if match: visible += 1
        self.card_canvas.configure(height=min(200, max(80, ((visible+1)//2)*55)))

    def _fill_scripts(self):
        for frm, _ in self._card_frames:
            frm.destroy()
        self._card_frames = []
        self._all_scripts = []
        self._selected_idx = None
        cat = self.cat_var.get()
        self._all_scripts = list(SCRIPTS_CONFIG.get(cat, []))
        for idx, s in enumerate(self._all_scripts):
            card = tk.Frame(self.card_inner, relief="solid", borderwidth=1, bg="white")
            r, c = divmod(idx, 2)
            card.grid(row=r, column=c, sticky=tk.EW, padx=3, pady=3)
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
        self._update_name_preview()
    def _on_card_click(self, idx):
        """点击卡片时高亮选中状态（蓝色背景+加粗边框）"""
        self._selected_idx = idx
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
        s = self._all_scripts[self._selected_idx]
        name = s.get("name") or s.get("script_name", "")
        try:
            self.search_var.trace_vdelete("w", self._search_trace)
        except Exception:
            pass
        self.search_var.set(name)
        self._search_trace = self.search_var.trace("w", lambda *a: self._filter_scripts())
    def _on_sched_change(self):
        st = self.sched_type.get()
        if st == "once": self.df.grid()
        else: self.df.grid_remove()
        if st == "weekly": self.wf.grid()
        else: self.wf.grid_remove()
        if st == "monthly": self.mf.grid()
        else: self.mf.grid_remove()
        self._update_name_preview()

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
            "time": self.hour_var.get().zfill(2) + ":" + self.min_var.get().zfill(2),
            "schedule_type": st,
            "scheduled_date": self.date_var.get() if st == "once" else "",
            "weekdays": sorted([k for k,v in self._wd_vars.items() if v.get()]) if st == "weekly" else [],
            "month_day": self.md_var.get() if st == "monthly" else 1,
        }
        if self.target_type.get() == "script":
            if self._selected_idx is None or self._selected_idx >= len(self._all_scripts):
                messagebox.showwarning("提示", "请选择一个脚本")
                return False
            s = self._all_scripts[self._selected_idx]
            self.result_data["target_type"] = "script"
            self.result_data["script_name"] = s["name"]
            self.result_data["script_path"] = s["path"]
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
        w = self.parent
        while w:
            if hasattr(w, "collect_params"):
                return w.collect_params()
            try: w = w.master
            except AttributeError: break
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
        self.count_label = ttk.Label(tool_frame,text="共 0 个任务",foreground="gray")
        self.count_label.pack(side=tk.LEFT)
        ttk.Button(tool_frame,text="添加定时任务",command=self._on_add,width=12).pack(side=tk.RIGHT,padx=(2,0))
        ttk.Button(tool_frame,text="刷新",command=self._on_refresh,width=8).pack(side=tk.RIGHT,padx=(2,0))

        # —— 列表区域 ——
        list_frame = ttk.Frame(self.parent)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        cols = ("id","name","target","schedule","next_run","last_run","status","enabled")
        self.tree = ttk.Treeview(list_frame,columns=cols,show="headings",height=15,selectmode=tk.BROWSE)
        for c,h in [("id","ID"),("name","任务名称"),("target","目标"),("schedule","定时方式"),("next_run","下次执行"),("last_run","上次执行"),("status","状态"),("enabled","启用")]:
            self.tree.heading(c,text=h)
        self.tree.column("id",width=28,minwidth=26,stretch=False,anchor=tk.CENTER)
        self.tree.column("name",width=80,minwidth=60,stretch=True)
        self.tree.column("target",width=55,minwidth=40,stretch=True)
        self.tree.column("schedule",width=55,minwidth=40,stretch=True)
        self.tree.column("next_run",width=75,minwidth=65,stretch=True)
        self.tree.column("last_run",width=75,minwidth=65,stretch=True)
        self.tree.column("status",width=46,minwidth=38,stretch=True,anchor=tk.CENTER)
        self.tree.column("enabled",width=35,minwidth=30,stretch=True,anchor=tk.CENTER)
        self.tree.tag_configure("running",foreground="#0000FF")
        self.tree.tag_configure("success",foreground="#008000")
        self.tree.tag_configure("disabled",foreground="#808080")

        v_scroll = ttk.Scrollbar(list_frame,orient=tk.VERTICAL,command=self.tree.yview)
        h_scroll = ttk.Scrollbar(list_frame,orient=tk.HORIZONTAL,command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        self.tree.pack(side=tk.LEFT,fill=tk.BOTH,expand=True)
        v_scroll.pack(side=tk.RIGHT,fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM,fill=tk.X)

        # —— 底部按钮 ——
        btn_frame = ttk.Frame(self.parent)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(5,0))
        ttk.Button(btn_frame,text="启用/禁用",command=self._on_toggle,width=10).pack(side=tk.LEFT,padx=2)
        ttk.Button(btn_frame,text="编辑",command=self._on_edit,width=8).pack(side=tk.LEFT,padx=2)
        ttk.Button(btn_frame,text="删除",command=self._on_delete,width=8).pack(side=tk.LEFT,padx=2)
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
        groups = []
        gp = _os.path.normpath(_os.path.join(self.gui.log_dir,"task_groups.json"))
        if _os.path.exists(gp):
            try:
                with open(gp,"r",encoding="utf-8") as f: data = json.load(f)
                groups = data if isinstance(data,list) else data.get("groups",[])
            except: groups = []
        dlg = AddSchedDialog(self.parent,groups,title="添加定时任务")
        if dlg.result_data:
            self.scheduler.add_task(dlg.result_data)
            self._refresh()
            self.gui._log("[定时任务] 已添加: " + dlg.result_data["name"])

    def _on_edit(self):
        sel = self.tree.selection()
        if not sel: messagebox.showwarning("提示","请先选择一个定时任务"); return
        tid = int(sel[0]); t = self.scheduler.get_task(tid)
        if not t: return
        ts = t.get("time","09:00")
        nt = simpledialog.askstring("编辑时间","修改执行时间 (HH:MM):\n当前: "+ts,initialvalue=ts,parent=self.parent)
        if not nt: return
        try:
            h,m = map(int,nt.split(":"))
            if not (0<=h<=23 and 0<=m<=59): raise ValueError
        except (ValueError,AttributeError):
            messagebox.showwarning("提示","时间格式错误，请使用 HH:MM"); return
        self.scheduler.update_task(tid,{"time": str(h).zfill(2)+":"+str(m).zfill(2)})
        self._refresh(); self.gui._log("[定时任务] 已更新: "+t.get("name",""))

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
