# -*- coding: utf-8 -*-
"""
定时任务面板
"""

import os as _os
import json, tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from config import SCRIPTS_CONFIG, CATEGORIES
from engine.scheduler import compute_next_run, format_schedule_desc


class AddSchedDialog(simpledialog.Dialog):
    def __init__(self, parent, groups, title=None):
        self.groups = groups
        self.result_data = None
        super().__init__(parent, title=title or "添加定时任务")

    def body(self, master):
        self.geometry("500x480")
        master.grid_rowconfigure(5, weight=1)
        master.grid_columnconfigure(1, weight=1)
        ttk.Label(master, text="任务名称:").grid(row=0,column=0,sticky=tk.W,padx=8,pady=4)
        self.name_var = tk.StringVar(value="")
        ttk.Entry(master, textvariable=self.name_var, width=30).grid(row=0,column=1,sticky=tk.EW,padx=8,pady=4)
        ttk.Label(master, text="目标类型:").grid(row=1,column=0,sticky=tk.W,padx=8,pady=4)
        self.target_type = tk.StringVar(value="script")
        tf = ttk.Frame(master)
        tf.grid(row=1,column=1,sticky=tk.W,padx=8,pady=4)
        ttk.Radiobutton(tf,text="单个脚本",variable=self.target_type,value="script",command=self._on_target_change).pack(side=tk.LEFT,padx=(0,10))
        ttk.Radiobutton(tf,text="任务编队",variable=self.target_type,value="group",command=self._on_target_change).pack(side=tk.LEFT)
        self.script_frame = ttk.LabelFrame(master,text="选择脚本",padding="5")
        self.script_frame.grid(row=2,column=0,columnspan=2,sticky=tk.EW,padx=8,pady=4)
        self.script_frame.grid_columnconfigure(1,weight=1)
        ttk.Label(self.script_frame,text="分类:").grid(row=0,column=0,sticky=tk.W,padx=4,pady=2)
        self.cat_var = tk.StringVar(value=CATEGORIES[0] if CATEGORIES else "")
        self.cat_combo = ttk.Combobox(self.script_frame,textvariable=self.cat_var,values=CATEGORIES,state="readonly",width=15)
        self.cat_combo.grid(row=0,column=1,sticky=tk.W,padx=4,pady=2)
        self.cat_combo.bind("<<ComboboxSelected>>",self._on_cat_change)
        ttk.Label(self.script_frame,text="脚本:").grid(row=1,column=0,sticky=tk.W,padx=4,pady=2)
        self.script_listbox = tk.Listbox(self.script_frame,height=5,exportselection=False)
        self.script_listbox.grid(row=1,column=1,sticky=tk.EW,padx=4,pady=2)
        self.group_frame = ttk.LabelFrame(master,text="选择编队",padding="5")
        self.group_frame.grid(row=3,column=0,columnspan=2,sticky=tk.EW,padx=8,pady=4)
        self.group_var = tk.StringVar()
        self.group_combo = ttk.Combobox(self.group_frame,textvariable=self.group_var,state="readonly",width=25)
        self.group_combo.pack(fill=tk.X,padx=4,pady=4)
        gnames = [g["name"] for g in self.groups]
        self.group_combo["values"] = gnames
        if gnames: self.group_combo.set(gnames[0])
        self.group_frame.grid_remove()
        sched_frame = ttk.LabelFrame(master,text="定时配置",padding="5")
        sched_frame.grid(row=4,column=0,columnspan=2,sticky=tk.EW,padx=8,pady=4)
        sched_frame.grid_columnconfigure(3,weight=1)
        ttk.Label(sched_frame,text="方式:").grid(row=0,column=0,sticky=tk.W,padx=4,pady=2)
        self.sched_type = tk.StringVar(value="daily")
        stf = ttk.Frame(sched_frame)
        stf.grid(row=0,column=1,columnspan=3,sticky=tk.W,padx=4,pady=2)
        for k,v in [("once","一次"),("daily","每天"),("weekly","每周"),("monthly","每月")]:
            rb = ttk.Radiobutton(stf,text=v,variable=self.sched_type,value=k,command=self._on_sched_change)
            rb.pack(side=tk.LEFT,padx=(0,8))
        ttk.Label(sched_frame,text="时间:").grid(row=1,column=0,sticky=tk.W,padx=4,pady=2)
        self.hour_var = tk.StringVar(value="09")
        self.min_var = tk.StringVar(value="00")
        hf = ttk.Frame(sched_frame)
        hf.grid(row=1,column=1,columnspan=3,sticky=tk.W,padx=4,pady=2)
        ttk.Spinbox(hf,from_=0,to=23,textvariable=self.hour_var,width=4,format="%02.0f").pack(side=tk.LEFT)
        ttk.Label(hf,text=":").pack(side=tk.LEFT)
        ttk.Spinbox(hf,from_=0,to=59,textvariable=self.min_var,width=4,format="%02.0f").pack(side=tk.LEFT)
        self.date_frame = ttk.Frame(sched_frame)
        self.date_frame.grid(row=2,column=0,columnspan=4,sticky=tk.W,padx=4,pady=2)
        ttk.Label(self.date_frame,text="执行日期:").pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value="2026-07-15")
        ttk.Entry(self.date_frame,textvariable=self.date_var,width=12).pack(side=tk.LEFT,padx=4)
        self.date_frame.grid_remove()
        self.wd_frame = ttk.Frame(sched_frame)
        self.wd_frame.grid(row=3,column=0,columnspan=4,sticky=tk.W,padx=4,pady=2)
        wd_names = ["周一","周二","周三","周四","周五","周六","周日"]
        self.wd_vars = {}
        for i,wn in enumerate(wd_names):
            self.wd_vars[i] = tk.BooleanVar(value=(i < 5))
            ttk.Checkbutton(self.wd_frame,text=wn,variable=self.wd_vars[i]).pack(side=tk.LEFT,padx=2)
        self.wd_frame.grid_remove()
        self.md_frame = ttk.Frame(sched_frame)
        self.md_frame.grid(row=4,column=0,columnspan=4,sticky=tk.W,padx=4,pady=2)
        ttk.Label(self.md_frame,text="每月第").pack(side=tk.LEFT)
        self.md_var = tk.IntVar(value=1)
        ttk.Spinbox(self.md_frame,from_=1,to=28,textvariable=self.md_var,width=5).pack(side=tk.LEFT,padx=4)
        ttk.Label(self.md_frame,text="天").pack(side=tk.LEFT)
        self.md_frame.grid_remove()
        self._fill_scripts()
        return master

    def _on_target_change(self):
        if self.target_type.get() == "script":
            self.script_frame.grid()
            self.group_frame.grid_remove()
        else:
            self.script_frame.grid_remove()
            self.group_frame.grid()

    def _on_cat_change(self, event=None):
        self._fill_scripts()

    def _fill_scripts(self):
        self.script_listbox.delete(0, tk.END)
        cat = self.cat_var.get()
        for s in SCRIPTS_CONFIG.get(cat, []):
            self.script_listbox.insert(tk.END, s["name"])
        if self.script_listbox.size() > 0:
            self.script_listbox.selection_set(0)
            n = self.script_listbox.get(0)
            prefix = {"once":"单次","daily":"每日","weekly":"每周","monthly":"每月"}
            self.name_var.set(prefix.get(self.sched_type.get(),"") + n)

    def _on_sched_change(self):
        st = self.sched_type.get()
        if st == "once": self.date_frame.grid()
        else: self.date_frame.grid_remove()
        if st == "weekly": self.wd_frame.grid()
        else: self.wd_frame.grid_remove()
        if st == "monthly": self.md_frame.grid()
        else: self.md_frame.grid_remove()
        sel = self.script_listbox.curselection()
        if sel:
            n = self.script_listbox.get(sel[0])
            prefix = {"once":"单次","daily":"每日","weekly":"每周","monthly":"每月"}
            self.name_var.set(prefix.get(st,"") + n)

    def validate(self):
        name = self.name_var.get().strip()
        if not name:
            messagebox.showwarning("提示", "请输入任务名称")
            return False
        import datetime as _dt
        h = self.hour_var.get().zfill(2)
        m = self.min_var.get().zfill(2)
        time_str = h + ":" + m
        st = self.sched_type.get()
        if st == "once":
            try:
                _dt.datetime.strptime(self.date_var.get(), "%Y-%m-%d")
            except ValueError:
                messagebox.showwarning("提示", "日期格式错误，请使用 YYYY-MM-DD")
                return False
        if st == "weekly" and not any(v.get() for v in self.wd_vars.values()):
            messagebox.showwarning("提示", "请至少选择一天")
            return False
        self.result_data = {
            "name": name, "time": time_str, "schedule_type": st,
            "scheduled_date": self.date_var.get() if st == "once" else "",
            "weekdays": sorted([k for k,v in self.wd_vars.items() if v.get()]) if st == "weekly" else [],
            "month_day": self.md_var.get() if st == "monthly" else 1,
        }
        if self.target_type.get() == "script":
            sel = self.script_listbox.curselection()
            if not sel:
                messagebox.showwarning("提示", "请选择一个脚本")
                return False
            sname = self.script_listbox.get(sel[0])
            cat = self.cat_var.get()
            for s in SCRIPTS_CONFIG.get(cat, []):
                if s["name"] == sname:
                    self.result_data["target_type"] = "script"
                    self.result_data["script_name"] = sname
                    self.result_data["script_path"] = s["path"]
                    self.result_data["category"] = cat
                    self.result_data["params"] = self._snapshot_params()
                    break
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
        self.tree.column("id",width=36,stretch=False,anchor=tk.CENTER)
        self.tree.column("name",width=140,stretch=True)
        self.tree.column("target",width=120,stretch=True)
        self.tree.column("schedule",width=120,stretch=True)
        self.tree.column("next_run",width=130,stretch=False)
        self.tree.column("last_run",width=130,stretch=False)
        self.tree.column("status",width=60,stretch=False,anchor=tk.CENTER)
        self.tree.column("enabled",width=40,stretch=False,anchor=tk.CENTER)
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
        for t in self.scheduler.tasks:
            tid = t.get("id",0); en = t.get("enabled",False); st = t.get("status","")
            tag = "running" if st == "运行中" else ("success" if st in ("已完成","等待") and en else ("disabled" if not en else ""))
            target = t.get("script_name","") or t.get("group_name","")
            if t.get("target_type") == "group": target = "[编队] " + target
            sd = format_schedule_desc(t); nr = t.get("next_run","")
            lr = t.get("last_run","") or "-"; et = "是" if en else "否"
            self.tree.insert("",tk.END,iid=str(tid),values=(tid,t.get("name",""),target,sd,nr,lr,st,et),tags=(tag,) if tag else ())
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
