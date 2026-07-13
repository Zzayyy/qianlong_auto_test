# -*- coding: utf-8 -*-
"""任务中心：可编排一组脚本按固定顺序自动执行

功能：
  - 把「当前分类」中已选中的脚本加入任务队列（带参数快照）
  - 任务队列支持上移 / 下移 / 删除 / 清空
  - 一键「开始顺序执行」，引擎逐个脚本串行执行
  - 执行中可「停止」（停止当前并中止后续任务）
  - 队列持久化到 logs/task_center.json，重启后保留
  - 编队（分组）功能：
      * 把当前队列保存为命名编队（保存为编队）
      * 选中编队即载入其脚本到下方队列，直接编辑（增/删/改/排序）
      * 「更新编队」把对队列的改动存回该编队
      * 「追加编队」把另一编队接在当前队列后（串联多个编队）
      * 「删除编队」删除选中的编队
  - 编队列表持久化到 logs/task_groups.json

说明：
  TaskCenter 通过调用方注入的 controller 引用 AutomationGUI 的能力：
    - controller.collect_params(): 返回当前界面参数 dict
    - controller._get_selected_script(): 返回左侧列表当前选中的脚本
    - controller.gui:               用于调用 _log / _set_status / history
"""

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from gui.history import (
    STATUS_SUCCESS,
    STATUS_FAILED,
    STATUS_ERROR,
    STATUS_STOPPED,
)

# 编队下拉中的占位项：表示当前队列是用户临时编排，未归入任何编队
GROUP_PLACEHOLDER = "（自定义队列）"


class GroupPicker(simpledialog.Dialog):
    """模态选择框：从编队列表中挑选一个（用于「追加编队」）"""

    def __init__(self, parent, groups, title=None):
        self.groups = groups
        self.result = None
        super().__init__(parent, title=title)

    def body(self, master):
        ttk.Label(master, text="选择要追加的编队:").pack(anchor=tk.W, padx=8, pady=(8, 2))
        self.listbox = tk.Listbox(
            master, height=min(10, max(3, len(self.groups))), width=32,
            selectmode=tk.SINGLE, exportselection=False
        )
        for g in self.groups:
            self.listbox.insert(tk.END, g["name"])
        self.listbox.pack(fill=tk.BOTH, expand=True, padx=8, pady=2)
        if self.groups:
            self.listbox.selection_set(0)
        return self.listbox

    def apply(self):
        sel = self.listbox.curselection()
        if sel:
            self.result = self.listbox.get(sel[0])


class TaskCenter:
    """任务中心面板（挂在主窗口右侧 Notebook 的「任务中心」标签页）"""

    # 任务状态
    ST_PENDING = "等待"       # 待执行
    ST_RUNNING = "执行中"      # 正在执行
    ST_SUCCESS = "成功"
    ST_FAILED = "失败"
    ST_ERROR = "异常"
    ST_STOPPED = "已停止"

    def __init__(self, parent, controller):
        """
        parent:   承载本面板的 Frame
        controller: 引用 AutomationGUI 实例（提供 collect_params / gui 等能力）
        """
        self.parent = parent
        self.controller = controller
        self.gui = controller.gui

        # 任务队列：list[dict]，每项 = {"category","script_name","script_path","params","status"}
        self.tasks = []

        # 编队（分组）：list[dict]，每项 = {"name","tasks":[任务快照...]}
        self.groups = []
        self._current_group = None   # 当前队列对应的编队名（None 表示自定义队列）
        self._dirty = False          # 当前队列相对所选编队是否有未保存修改

        # 运行状态
        self.is_running = False
        self._current_index = -1       # 当前执行到的下标（用于状态着色与停止）
        self._record_id = None         # 当前任务的 history 记录 id
        self._stop = False             # 用户请求停止（中止后续任务）

        # 持久化路径
        self.file_path = os.path.join(self.gui.log_dir, "task_center.json")
        self.groups_path = os.path.join(self.gui.log_dir, "task_groups.json")

        self._build_ui()
        self._load_groups()
        self._load()

    # ====================== UI 构建 ======================
    def _build_ui(self):
        parent = self.parent

        # —— 工具条 ——
        tool_frame = ttk.Frame(parent)
        tool_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        self.summary_label = ttk.Label(tool_frame, text="队列: 0 项 | 等待执行", foreground="gray")
        self.summary_label.pack(side=tk.LEFT)

        ttk.Button(tool_frame, text="开始顺序执行", command=self.start, width=13).pack(side=tk.RIGHT, padx=(2, 0))
        ttk.Button(tool_frame, text="停止", command=self._on_stop, width=8).pack(side=tk.RIGHT, padx=(2, 0))

        # —— 编队管理条 ——
        group_frame = ttk.Frame(parent)
        group_frame.pack(side=tk.TOP, fill=tk.X, pady=(0, 5))

        row1 = ttk.Frame(group_frame)
        row1.pack(side=tk.TOP, fill=tk.X)
        ttk.Label(row1, text="编队:").pack(side=tk.LEFT, padx=(0, 4))
        self.group_combo = ttk.Combobox(row1, state="readonly", width=20)
        self.group_combo.pack(side=tk.LEFT, padx=(0, 4))
        self.group_combo.bind("<<ComboboxSelected>>", self._on_group_selected)
        self.group_hint = ttk.Label(row1, text="", foreground="#f44747")
        self.group_hint.pack(side=tk.LEFT)

        row2 = ttk.Frame(group_frame)
        row2.pack(side=tk.TOP, fill=tk.X, pady=(3, 0))
        self.btn_save_group = ttk.Button(
            row2, text="保存为编队", command=self.save_group, width=11
        )
        self.btn_save_group.pack(side=tk.LEFT, padx=2)
        self.btn_update_group = ttk.Button(
            row2, text="更新编队", command=self.update_group, width=11
        )
        self.btn_update_group.pack(side=tk.LEFT, padx=2)
        self.btn_append_group = ttk.Button(
            row2, text="追加编队", command=self.append_group, width=11
        )
        self.btn_append_group.pack(side=tk.LEFT, padx=2)
        self.btn_del_group = ttk.Button(
            row2, text="删除编队", command=self.delete_group, width=11
        )
        self.btn_del_group.pack(side=tk.LEFT, padx=2)

        # —— 队列列表 ——
        list_frame = ttk.Frame(parent)
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        columns = ("no", "task", "category", "status")
        self.tree = ttk.Treeview(
            list_frame, columns=columns, show="headings", height=12, selectmode=tk.BROWSE
        )
        self.tree.heading("no", text="#")
        self.tree.heading("task", text="任务")
        self.tree.heading("category", text="分类")
        self.tree.heading("status", text="状态")
        self.tree.column("no", width=34, stretch=False, anchor=tk.CENTER)
        self.tree.column("task", width=160, stretch=False)
        self.tree.column("category", width=80, stretch=False)
        self.tree.column("status", width=55, stretch=False, anchor=tk.CENTER)

        # 状态配色
        self.tree.tag_configure("running", foreground="#0000FF")
        self.tree.tag_configure("success", foreground="#008000")
        self.tree.tag_configure("failed", foreground="#f44747")
        self.tree.tag_configure("error", foreground="#f44747")
        self.tree.tag_configure("stopped", foreground="#BDB76B")
        self.tree.tag_configure("pending", foreground="#808080")

        v_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=v_scroll.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # 拖拽排序 + 右键菜单
        self.tree.bind("<ButtonPress-1>", self._on_drag_start)
        self.tree.bind("<B1-Motion>", self._on_drag_motion)
        self.tree.bind("<ButtonRelease-1>", self._on_drag_end)
        self.tree.bind("<Button-3>", self._on_right_click)
        self.menu = tk.Menu(self.tree, tearoff=0)
        self.menu.add_command(label="复制", command=self._ctx_copy)
        self.menu.add_command(label="上移", command=self._ctx_move_up)
        self.menu.add_command(label="下移", command=self._ctx_move_down)
        self.menu.add_command(label="删除", command=self._ctx_delete)
        self.menu.add_separator()
        self.menu.add_command(label="清空", command=self._ctx_clear)
        self._drag_index = None
        self._drag_motion = False
        self._drag_start_y = 0

        # 拖拽落点指示线（跟随光标显示插入位置）
        self.drop_indicator = tk.Frame(self.tree.master, height=2, background="#0078d4")
        self.drop_indicator.place_forget()

        # —— 任务操作按钮 ——
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))

        ttk.Button(btn_frame, text="加入当前选中脚本", command=self.add_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="上移", command=self._move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下移", command=self._move_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self._delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空", command=self._clear).pack(side=tk.LEFT, padx=2)

    # ====================== 编队（分组）管理 ======================
    def _load_groups(self):
        """从 task_groups.json 载入编队列表"""
        self.groups = []
        if os.path.exists(self.groups_path):
            try:
                with open(self.groups_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    self.groups = data.get("groups", [])
                elif isinstance(data, list):
                    self.groups = data
            except Exception:
                self.groups = []
        self._refresh_group_combo()

    def _save_groups(self):
        """保存编队列表到 task_groups.json"""
        try:
            with open(self.groups_path, "w", encoding="utf-8") as f:
                json.dump({"groups": self.groups}, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.gui.logger.error(f"编队保存失败: {e}")

    def _refresh_group_combo(self, select=None):
        """刷新编队下拉框，并复位当前编队标记"""
        names = [GROUP_PLACEHOLDER] + [g["name"] for g in self.groups]
        self.group_combo["values"] = names
        if select is None or select not in names:
            self.group_combo.set(GROUP_PLACEHOLDER)
            self._current_group = None
        else:
            self.group_combo.set(select)
            self._current_group = None if select == GROUP_PLACEHOLDER else select
        self._refresh_group_controls()
        self._update_group_hint()

    def _refresh_group_controls(self):
        """根据下拉选择启用/禁用「更新编队 / 删除编队」按钮"""
        name = self.group_combo.get()
        has_group = bool(name) and name != GROUP_PLACEHOLDER and any(g["name"] == name for g in self.groups)
        state = tk.NORMAL if has_group else tk.DISABLED
        self.btn_del_group.config(state=state)
        self.btn_update_group.config(state=state)

    def _update_group_hint(self):
        """显示当前编队是否有未保存修改"""
        if self._current_group and self._dirty:
            self.group_hint.config(text="● 有未保存修改")
        else:
            self.group_hint.config(text="")

    def _on_group_selected(self, event=None):
        """下拉选择编队：直接载入该编队组合好的脚本（覆盖当前队列）"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            self._refresh_group_combo(self._current_group or GROUP_PLACEHOLDER)
            return
        name = self.group_combo.get()
        if not name:
            return
        # 主动选「自定义队列」：脱离编队，但保留当前队列内容
        if name == GROUP_PLACEHOLDER:
            self._current_group = None
            self._dirty = False
            self._refresh_group_controls()
            self._update_group_hint()
            return
        group = next((g for g in self.groups if g["name"] == name), None)
        if not group:
            self._refresh_group_combo(GROUP_PLACEHOLDER)
            return
        self._load_group(group)
        self._current_group = name
        self._dirty = False
        self._refresh_group_controls()
        self._update_group_hint()
        self.gui._log(f"[任务中心] 已载入编队: {name}（{len(self.tasks)} 个脚本）")

    def _load_group(self, group):
        """把编队中的任务快照载入到当前队列"""
        self.tasks = []
        for row in group.get("tasks", []):
            if not os.path.exists(row.get("script_path", "")):
                continue  # 脚本已不存在则跳过
            self.tasks.append({
                "category": row["category"],
                "script_name": row["script_name"],
                "script_path": row["script_path"],
                "params": row.get("params", {}),
                "status": self.ST_PENDING,
            })
        self._save()
        self._refresh()

    def _task_to_dict(self, t):
        """任务项 -> 可序列化 dict"""
        return {
            "category": t["category"],
            "script_name": t["script_name"],
            "script_path": t["script_path"],
            "params": t["params"],
        }

    def save_group(self):
        """把当前队列保存为命名编队（已存在则提示覆盖）"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            return
        if not self.tasks:
            messagebox.showwarning("提示", "当前队列为空，无法保存编队")
            return

        name = simpledialog.askstring("保存为编队", "请输入编队名称:", parent=self.parent)
        if not name:
            return
        name = name.strip()
        if not name:
            return
        if name == GROUP_PLACEHOLDER:
            messagebox.showwarning("提示", f"名称「{GROUP_PLACEHOLDER}」为保留字，请换一个")
            return

        existing = next((g for g in self.groups if g["name"] == name), None)
        if existing:
            if not messagebox.askyesno("确认", f"已存在编队「{name}」，是否覆盖？"):
                return
            existing["tasks"] = [self._task_to_dict(t) for t in self.tasks]
        else:
            self.groups.append({
                "name": name,
                "tasks": [self._task_to_dict(t) for t in self.tasks],
            })

        self._save_groups()
        self._dirty = False
        self._refresh_group_combo(select=name)
        self.gui._log(f"[任务中心] 已保存编队: {name}（{len(self.tasks)} 个脚本）")

    def update_group(self):
        """把当前队列（含增删改后的结果）存回当前选中的编队"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            return
        name = self._current_group
        if not name:
            messagebox.showwarning("提示", "请先在「编队」下拉中选择一个编队，再对其脚本进行编辑并保存")
            return
        group = next((g for g in self.groups if g["name"] == name), None)
        if not group:
            messagebox.showwarning("提示", f"编队「{name}」不存在，请重新选择")
            self._refresh_group_combo(GROUP_PLACEHOLDER)
            return
        if not self.tasks:
            if not messagebox.askyesno("确认", f"当前队列为空，保存后将清空编队「{name}」，是否继续？"):
                return
        group["tasks"] = [self._task_to_dict(t) for t in self.tasks]
        self._save_groups()
        self._dirty = False
        self._refresh_group_combo(select=name)
        self.gui._log(f"[任务中心] 已更新编队: {name}（{len(self.tasks)} 个脚本）")

    def append_group(self):
        """弹窗选择另一编队，将其脚本追加到当前队列末尾（串联多个编队）"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            return
        if not self.groups:
            messagebox.showinfo("提示", "还没有任何编队，请先「保存为编队」")
            return
        # 用独立弹窗选择要追加的编队，避免改动当前下拉选择（否则会被覆盖）
        picker = GroupPicker(self.parent, self.groups, title="追加编队")
        name = picker.result
        if not name:
            return
        group = next((g for g in self.groups if g["name"] == name), None)
        if not group:
            return
        added = 0
        for row in group.get("tasks", []):
            if not os.path.exists(row.get("script_path", "")):
                continue
            self.tasks.append({
                "category": row["category"],
                "script_name": row["script_name"],
                "script_path": row["script_path"],
                "params": row.get("params", {}),
                "status": self.ST_PENDING,
            })
            added += 1
        if added == 0:
            messagebox.showinfo("提示", f"编队「{name}」中的脚本均已不存在，未追加")
            return
        self._save()
        self._refresh()
        # 追加后脱离原编队：下拉复位为「自定义队列」，避免「更新编队」误存回原编队
        self._dirty = False
        self._refresh_group_combo(GROUP_PLACEHOLDER)
        self.gui._log(f"[任务中心] 已追加编队: {name}（新增 {added} 个脚本，共 {len(self.tasks)} 项）")

    def delete_group(self):
        """删除下拉选中的编队"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            return
        name = self.group_combo.get()
        if not name or name == GROUP_PLACEHOLDER:
            return
        group = next((g for g in self.groups if g["name"] == name), None)
        if not group:
            return
        if not messagebox.askyesno("确认", f"确定删除编队「{name}」？（仅删除编队，不影响当前队列）"):
            return
        self.groups.remove(group)
        self._save_groups()
        self._refresh_group_combo(GROUP_PLACEHOLDER)
        self._dirty = False
        self.gui._log(f"[任务中心] 已删除编队: {name}")

    # ====================== 队列操作 ======================
    def add_selected(self):
        """把「当前分类」中选中的脚本加入队列（带参数快照）"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中，请先停止")
            return

        script = self.controller._get_selected_script()
        if not script:
            messagebox.showwarning("提示", "请先在左侧列表中选择一个脚本")
            return

        if not os.path.exists(script["path"]):
            messagebox.showerror("错误", f"脚本文件不存在:\n{script['path']}")
            return

        # 下单需 Excel，但全选撤单/一键导出不需要
        category = self.controller.current_category
        if category == "下单" and script["name"] not in ("5.期权下单_一键导出", "6.全选撤单"):
            if not self.controller.xlsx_file.get():
                messagebox.showwarning("提示", "该下单脚本需要先选择 Excel 配置文件")
                return

        params = self.controller.collect_params()
        self.tasks.append({
            "category": category,
            "script_name": script["name"],
            "script_path": script["path"],
            "params": params,
            "status": self.ST_PENDING,
        })
        self._save()
        self._refresh()
        # 手动修改后队列相对编队发生变化
        self._dirty = True
        self._update_group_hint()
        self.gui._log(f"[任务中心] 已加入队列: {script['name']}（{category}）")

    def _selected_index(self):
        sel = self.tree.selection()
        if not sel:
            return -1
        return int(self.tree.index(sel[0]))

    def _move_up(self):
        if self.is_running:
            return
        i = self._selected_index()
        if i <= 0:
            return
        self.tasks[i - 1], self.tasks[i] = self.tasks[i], self.tasks[i - 1]
        self._save()
        self._refresh()
        self.tree.selection_set(self.tree.get_children()[i - 1])
        self._dirty = True
        self._update_group_hint()

    def _move_down(self):
        if self.is_running:
            return
        i = self._selected_index()
        if i < 0 or i >= len(self.tasks) - 1:
            return
        self.tasks[i + 1], self.tasks[i] = self.tasks[i], self.tasks[i + 1]
        self._save()
        self._refresh()
        self.tree.selection_set(self.tree.get_children()[i + 1])
        self._dirty = True
        self._update_group_hint()

    def _delete(self):
        if self.is_running:
            return
        i = self._selected_index()
        if i < 0:
            return
        removed = self.tasks.pop(i)
        self._save()
        self._refresh()
        self._dirty = True
        self._update_group_hint()
        self.gui._log(f"[任务中心] 已移除: {removed['script_name']}")

    def _clear(self):
        if self.is_running:
            return
        if not self.tasks:
            return
        if messagebox.askyesno("确认", "确定清空任务队列？"):
            self.tasks = []
            self._save()
            self._refresh()
            self._dirty = True
            self._update_group_hint()
            self.gui._log("[任务中心] 已清空任务队列")

    # ====================== 拖拽排序 / 右键菜单 ======================
    def _on_drag_start(self, event):
        """记录拖拽起点（鼠标左键按下）"""
        if self.is_running:
            return
        # 重置来自脚本列表的拖入状态，避免误判
        self.controller._drag_script = None
        self.controller._drag_active = False
        row = self.tree.identify_row(event.y)
        if not row:
            self._drag_index = None
            return
        self._drag_index = int(row)
        self._drag_motion = False
        self._drag_start_y = event.y

    def _on_drag_motion(self, event):
        """移动超过阈值即视为拖拽（避免与单击选择冲突）"""
        ctrl = self.controller
        # 来自脚本列表的拖入：实时显示落点
        if getattr(ctrl, "_drag_active", False) and getattr(ctrl, "_drag_script", None) is not None:
            if not self.is_running:
                self._update_drop_indicator(event.y)
            return
        if self._drag_index is None or self.is_running:
            return
        if abs(event.y - self._drag_start_y) < 6:
            return
        self._drag_motion = True
        self._update_drop_indicator(event.y)

    def _on_drag_end(self, event):
        """松手时根据落点重排队列"""
        self._hide_drop_indicator()
        if self._drag_index is None:
            return
        src = self._drag_index
        self._drag_index = None
        if not self._drag_motion or self.is_running:
            return
        target = self._drop_target(event.y)
        if not target:
            return
        self._reorder(src, target[0], target[1])

    def _drop_target(self, y):
        """返回 (目标下标, 是否落在目标行下半部)"""
        children = self.tree.get_children()
        if not children:
            return None
        row = self.tree.identify_row(y)
        if not row:
            try:
                first_bbox = self.tree.bbox(children[0])
                last_bbox = self.tree.bbox(children[-1])
            except Exception:
                return None
            if y < first_bbox[1]:
                return (0, False)
            return (len(children) - 1, True)
        tgt = int(row)
        try:
            bbox = self.tree.bbox(row)
        except Exception:
            return (tgt, False)
        after = y > bbox[1] + bbox[3] / 2
        return (tgt, after)

    def _update_drop_indicator(self, y):
        """在落点处显示一条插入指示线（y 相对 tree 坐标）"""
        children = self.tree.get_children()
        if not children:
            self._hide_drop_indicator()
            return
        row = self.tree.identify_row(y)
        if not row:
            try:
                first = self.tree.bbox(children[0])
                last = self.tree.bbox(children[-1])
            except Exception:
                self._hide_drop_indicator()
                return
            line_y = first[1] if y < first[1] else (last[1] + last[3])
        else:
            try:
                bbox = self.tree.bbox(row)
            except Exception:
                self._hide_drop_indicator()
                return
            after = y > bbox[1] + bbox[3] / 2
            line_y = (bbox[1] + bbox[3]) if after else bbox[1]
        # tree 左上角即 list_frame(0,0)，指示线相对 master 放置
        self.drop_indicator.place(x=0, y=line_y, relwidth=1.0, height=2)

    def _hide_drop_indicator(self):
        self.drop_indicator.place_forget()

    def _reorder(self, src, tgt, after):
        """将 src 移动到 tgt 位置（after 表示是否落在 tgt 行之后）"""
        if src < 0 or src >= len(self.tasks):
            return
        item = self.tasks[src]
        new_tasks = list(self.tasks)
        new_tasks.pop(src)
        if after:
            insert_at = tgt + 1 if src > tgt else tgt
        else:
            insert_at = tgt if src > tgt else tgt - 1
        insert_at = max(0, min(insert_at, len(new_tasks)))
        new_tasks.insert(insert_at, item)
        # 无实际变化则不处理
        if [t["script_path"] for t in new_tasks] == [t["script_path"] for t in self.tasks]:
            return
        self.tasks = new_tasks
        self._save()
        self._refresh()
        self._dirty = True
        self._update_group_hint()
        # 按对象身份定位（内容相同的重复项用 == 会命中第一个）
        for idx, t in enumerate(self.tasks):
            if t is item:
                self.tree.selection_set(str(idx))
                break

    # ====================== 从脚本列表拖入 ======================
    def add_script_from_drop(self, script, tree_y):
        """从脚本列表拖入：在 tree_y 对应位置插入脚本"""
        self._hide_drop_indicator()
        if self.is_running:
            return
        target = self._drop_target(tree_y)
        self.add_script(script, self._insert_index_from_target(target))

    def add_script(self, script, index=None):
        """把指定脚本加入队列（index 为插入位置，None 则追加）"""
        if self.is_running:
            return
        if not os.path.exists(script["path"]):
            messagebox.showerror("错误", f"脚本文件不存在:\n{script['path']}")
            return
        category = script.get("category") or self.controller.current_category
        if category == "下单" and script["name"] not in ("5.期权下单_一键导出", "6.全选撤单"):
            if not self.controller.xlsx_file.get():
                messagebox.showwarning("提示", "该下单脚本需要先选择 Excel 配置文件")
                return
        params = self.controller.collect_params()
        item = {
            "category": category,
            "script_name": script["name"],
            "script_path": script["path"],
            "params": params,
            "status": self.ST_PENDING,
        }
        if index is None or index >= len(self.tasks):
            self.tasks.append(item)
            new_pos = len(self.tasks) - 1
        else:
            index = max(0, index)
            self.tasks.insert(index, item)
            new_pos = index
        self._save()
        self._refresh()
        self._dirty = True
        self._update_group_hint()
        self.tree.selection_set(str(new_pos))
        self.gui._log(f"[任务中心] 已加入队列: {script['name']}（{category}）")

    def _insert_index_from_target(self, target):
        """(目标行下标, 是否落在下半部) -> 新项插入位置"""
        if not target:
            return len(self.tasks)
        tgt, after = target
        return tgt + 1 if after else tgt

    def _on_right_click(self, event):
        """右键菜单：复制 / 上移 / 下移 / 删除 / 清空"""
        if self.is_running:
            return
        row = self.tree.identify_row(event.y)
        if row:
            self.tree.selection_set(row)
        has_sel = self._selected_index() >= 0
        state = tk.NORMAL if has_sel else tk.DISABLED
        for i in range(4):  # 复制 / 上移 / 下移 / 删除 需要选中项
            self.menu.entryconfig(i, state=state)
        self.menu.post(event.x_root, event.y_root)

    def _ctx_copy(self):
        """复制选中任务为一项新任务（参数一并复制），插入到其下方"""
        i = self._selected_index()
        if i < 0:
            return
        src = self.tasks[i]
        item = {
            "category": src["category"],
            "script_name": src["script_name"],
            "script_path": src["script_path"],
            "params": dict(src["params"]),
            "status": self.ST_PENDING,
        }
        self.tasks.insert(i + 1, item)
        self._save()
        self._refresh()
        self._dirty = True
        self._update_group_hint()
        self.tree.selection_set(str(i + 1))

    def _ctx_move_up(self):
        self._move_up()

    def _ctx_move_down(self):
        self._move_down()

    def _ctx_delete(self):
        self._delete()

    def _ctx_clear(self):
        self._clear()

    # ====================== 执行编排 ======================
    def start(self):
        """开始顺序执行队列"""
        if self.is_running:
            messagebox.showwarning("提示", "任务正在执行中")
            return
        if not self.tasks:
            messagebox.showwarning("提示", "任务队列为空，请先加入脚本")
            return

        self.is_running = True
        self._stop = False
        self._current_index = -1

        # 复位所有任务状态
        for t in self.tasks:
            t["status"] = self.ST_PENDING

        self._set_running_ui(True)
        self._refresh()
        self.gui._log("\n" + "=" * 60)
        self.gui._log(f"[任务中心] 开始顺序执行，共 {len(self.tasks)} 个任务")
        self.gui._set_status(f"任务中心: 0/{len(self.tasks)} 准备就绪", running=True)

        # 通过主窗口的执行引擎驱动（复用其回调机制）
        self.controller.run_task_center(self)

    def _on_stop(self):
        if not self.is_running:
            return
        self._stop = True
        self.gui._log("\n[任务中心] 收到停止指令，将在当前任务结束后中止后续任务...")
        self.gui._set_status("任务中心: 正在停止...")
        self.controller.stop_task_center()

    def run_next(self):
        """由 controler 在主线程回调：执行下一个未执行任务；无则收尾"""
        if self._stop:
            self._finish_all("已停止（用户手动）")
            return

        self._current_index += 1
        if self._current_index >= len(self.tasks):
            self._finish_all("全部执行完成")
            return

        item = self.tasks[self._current_index]
        item["status"] = self.ST_RUNNING
        self._refresh()
        self.gui._set_status(
            f"任务中心: {self._current_index + 1}/{len(self.tasks)} 执行 {item['script_name']}",
            running=True
        )
        self.gui._log(f"\n{'='*60}")
        self.gui._log(f"[任务中心] ({self._current_index + 1}/{len(self.tasks)}) 开始: {item['script_name']}")

        # 写 history 记录
        self._record_id = self.gui.history.add_record(item["script_name"], item["category"])
        self.gui._refresh_history()

        # 交给主窗口引擎执行（脚本路径 + 已快照的参数）
        self.controller.execute_task_item(item, self._record_id)

    def on_finish(self, return_code, elapsed, item):
        """单个任务结束回调：更新状态后执行下一个"""
        if self._current_index < 0 or self._current_index >= len(self.tasks):
            return
        cur = self.tasks[self._current_index]

        if self._stop:
            cur["status"] = self.ST_STOPPED
        elif return_code == 0:
            cur["status"] = self.ST_SUCCESS
        else:
            cur["status"] = self.ST_FAILED

        detail = f"退出码: {return_code}" if return_code != 0 else ""
        if self._record_id is not None:
            status_map = {
                self.ST_SUCCESS: STATUS_SUCCESS,
                self.ST_FAILED: STATUS_FAILED,
                self.ST_STOPPED: STATUS_STOPPED,
            }
            self.gui.history.update_record(
                self._record_id,
                status_map.get(cur["status"], STATUS_FAILED),
                elapsed, detail
            )
            self._record_id = None
            self.gui._refresh_history()

        self._save()
        self._refresh()
        self.run_next()

    def on_error(self, exc, item):
        """单个任务异常回调"""
        if self._current_index < 0 or self._current_index >= len(self.tasks):
            return
        cur = self.tasks[self._current_index]
        cur["status"] = self.ST_ERROR if not self._stop else self.ST_STOPPED

        if self._record_id is not None:
            self.gui.history.update_record(
                self._record_id,
                STATUS_STOPPED if self._stop else STATUS_ERROR,
                detail=str(exc)
            )
            self._record_id = None
            self.gui._refresh_history()

        self._save()
        self._refresh()
        self.run_next()

    def _finish_all(self, msg):
        self.is_running = False
        self._set_running_ui(False)
        self._refresh()
        self.gui._log(f"\n{'='*60}")
        self.gui._log(f"[任务中心] {msg}")
        self.gui._log(f"{'='*60}")
        self.gui._set_status(f"任务中心: {msg}")
        self.gui._reset_running_state_if_idle()

    # ====================== UI 辅助 ======================
    def _set_running_ui(self, running):
        """切换运行/非运行时的按钮可用状态（由 controller 主线程调用）"""
        # 工具条按钮
        for child in self.parent.winfo_children():
            if isinstance(child, ttk.Frame):
                for b in child.winfo_children():
                    if isinstance(b, ttk.Button):
                        if b.cget("text") == "开始顺序执行":
                            b.config(state=tk.DISABLED if running else tk.NORMAL)
                        elif b.cget("text") == "停止":
                            b.config(state=tk.NORMAL if running else tk.DISABLED)

        # 编队控件：运行中禁用，非运行恢复
        self.group_combo.config(state=tk.DISABLED if running else "readonly")
        gstate = tk.DISABLED if running else tk.NORMAL
        self.btn_save_group.config(state=gstate)
        self.btn_append_group.config(state=gstate)
        if running:
            self.btn_update_group.config(state=tk.DISABLED)
            self.btn_del_group.config(state=tk.DISABLED)
        else:
            self._refresh_group_controls()

    def _refresh(self):
        """刷新队列列表（主线程调用）"""
        self.tree.delete(*self.tree.get_children())
        status_tag = {
            self.ST_PENDING: "pending",
            self.ST_RUNNING: "running",
            self.ST_SUCCESS: "success",
            self.ST_FAILED: "failed",
            self.ST_ERROR: "error",
            self.ST_STOPPED: "stopped",
        }
        done = 0
        for idx, t in enumerate(self.tasks):
            tag = status_tag.get(t["status"], "")
            tree_idx = str(idx)
            self.tree.insert(
                "", tk.END, iid=tree_idx,
                values=(idx + 1, t["script_name"], t["category"], t["status"]),
                tags=(tag,) if tag else (),
            )
            if t["status"] in (self.ST_SUCCESS, self.ST_FAILED, self.ST_ERROR, self.ST_STOPPED):
                done += 1

        if self.is_running:
            self.summary_label.config(
                text=f"队列: {len(self.tasks)} 项 | 进度 {self._current_index + 1}/{len(self.tasks)}"
            )
        else:
            self.summary_label.config(
                text=f"队列: {len(self.tasks)} 项 | 已完成 {done} 项"
            )

    # ====================== 持久化 ======================
    def _save(self):
        try:
            data = [{
                "category": t["category"],
                "script_name": t["script_name"],
                "script_path": t["script_path"],
                "params": t["params"],
            } for t in self.tasks]
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.gui.logger.error(f"任务中心保存失败: {e}")

    def _load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for row in data:
                    if not os.path.exists(row.get("script_path", "")):
                        continue  # 脚本已不存在则跳过
                    self.tasks.append({
                        "category": row["category"],
                        "script_name": row["script_name"],
                        "script_path": row["script_path"],
                        "params": row.get("params", {}),
                        "status": self.ST_PENDING,
                    })
            except Exception:
                self.tasks = []
        # 载入的队列视为自定义队列
        self._current_group = None
        self._dirty = False
        self._refresh_group_combo(GROUP_PLACEHOLDER)
        self._refresh()
