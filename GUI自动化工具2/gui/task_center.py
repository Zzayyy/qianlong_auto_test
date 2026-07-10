# -*- coding: utf-8 -*-
"""任务中心：可编排一组脚本按固定顺序自动执行

功能：
  - 把「当前分类」中已选中的脚本加入任务队列（带参数快照）
  - 任务队列支持上移 / 下移 / 删除 / 清空
  - 一键「开始顺序执行」，引擎逐个脚本串行执行
  - 执行中可「停止」（停止当前并中止后续任务）
  - 队列持久化到 logs/task_center.json，重启后保留

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

        # 运行状态
        self.is_running = False
        self._current_index = -1       # 当前执行到的下标（用于状态着色与停止）
        self._record_id = None         # 当前任务的 history 记录 id
        self._stop = False             # 用户请求停止（中止后续任务）

        # 持久化路径
        self.file_path = os.path.join(self.gui.log_dir, "task_center.json")

        self._build_ui()
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

        # —— 任务操作按钮 ——
        btn_frame = ttk.Frame(parent)
        btn_frame.pack(side=tk.TOP, fill=tk.X, pady=(5, 0))

        ttk.Button(btn_frame, text="加入当前选中脚本", command=self.add_selected).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="上移", command=self._move_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="下移", command=self._move_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="删除", command=self._delete).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame, text="清空", command=self._clear).pack(side=tk.LEFT, padx=2)

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

    def _delete(self):
        if self.is_running:
            return
        i = self._selected_index()
        if i < 0:
            return
        removed = self.tasks.pop(i)
        self._save()
        self._refresh()
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
            self.gui._log("[任务中心] 已清空任务队列")

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
        self._refresh()
