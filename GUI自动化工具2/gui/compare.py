# -*- coding: utf-8 -*-
"""结果比对标签页
====================
UI 功能：
  - 上传两个文件夹（基准 / 对照）：通过「浏览」按钮选择，若已安装 tkinterdnd2 也支持直接拖入
  - 对两文件夹内文件名一致的文件做内容比对（先实现 .xls/.xlsx 行级比对，其余类型做哈希比对）
  - 比对进度实时反馈，结果以树形列表呈现（一致 / 不一致 / 仅基准 / 仅对照 / 错误）
  - 全部比对完成后，输出一份总结报告（可导出为 txt 文件，也可直接查看）

设计：
  - 与 TaskCenter / SchedulerPanel 同构：ComparePanel(parent, controller)
  - 上传以「浏览」按钮作为最可靠入口；
    拖入使用 tkinterdnd2（第三方库，非必需），不依赖 ctypes WNDPROC 子类化 ——
    后者在 Tk 子窗口上极不稳定，容易导致界面渲染/点击崩溃。
  - 比对在后台线程执行，避免大文件夹卡死界面。
"""

import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

from core import compare as cmp


# ====================== 可选拖入支持（tkinterdnd2） ======================
# 拖入通过 tkinterdnd2 实现（pip install tkinterdnd2），完全避免 ctypes WNDPROC
# 子类化带来的稳定性问题。若未安装，仅保留「浏览」按钮，功能不受影响。
_HAS_DND = False
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
    _HAS_DND = True
except ImportError:
    pass

_IS_WIN = (os.name == "nt")


class ComparePanel:
    """结果比对待办面板（挂在主窗口右侧 Notebook 的「结果比对」标签页）"""

    # 结果状态 -> (树标签, 配色)
    TAGS = {
        "equal": ("一致", "c_equal"),
        "diff": ("不一致", "c_diff"),
        "only_a": ("仅基准", "c_only"),
        "only_b": ("仅对照", "c_only"),
        "error": ("错误", "c_diff"),
    }

    def __init__(self, parent, controller):
        self.parent = parent
        self.controller = controller
        self.gui = controller.gui

        self.folder_a = tk.StringVar(value="")
        self.folder_b = tk.StringVar(value="")
        self.recursive = tk.BooleanVar(value=False)
        self.ignore_order = tk.BooleanVar(value=True)

        self.data = None          # 最近一次比对结果结构
        self._running = False
        self._dnd_initialized = False  # tkinterdnd2 只需初始化一次

        self._build_ui()

    # ====================== UI 构建 ======================
    def _build_ui(self):
        parent = self.parent

        # —— 顶部：两个文件夹上传区（并排）——
        top = ttk.Frame(parent)
        top.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))
        top.columnconfigure(0, weight=1)
        top.columnconfigure(1, weight=1)

        self._build_drop_zone(top, 0, "基准文件夹 (A)", self.folder_a, "a")
        self._build_drop_zone(top, 1, "对照文件夹 (B)", self.folder_b, "b")

        # —— 选项 + 操作按钮 ——
        ctrl = ttk.Frame(parent)
        ctrl.pack(side=tk.TOP, fill=tk.X, pady=(0, 6))

        ttk.Checkbutton(ctrl, text="递归子目录", variable=self.recursive).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Checkbutton(ctrl, text="忽略行顺序（按内容比对）", variable=self.ignore_order).pack(side=tk.LEFT, padx=(0, 8))

        self.status_label = ttk.Label(ctrl, text="就绪", foreground="gray")
        self.status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.compare_btn = ttk.Button(ctrl, text="开始比对", command=self._on_compare, width=12)
        self.compare_btn.pack(side=tk.RIGHT, padx=(4, 0))
        self.export_btn = ttk.Button(ctrl, text="导出报告", command=self._on_export, width=12, state=tk.DISABLED)
        self.export_btn.pack(side=tk.RIGHT, padx=(4, 0))

        # —— 结果列表 ——
        list_frame = ttk.LabelFrame(parent, text="比对结果", padding="6")
        list_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # 嵌套内框：tree + v_scroll 在内框并排，h_scroll 放在 list_frame 底部横跨整宽
        tree_box = ttk.Frame(list_frame)
        tree_box.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        columns = ("name", "type", "a", "b", "result", "summary")
        self.tree = ttk.Treeview(tree_box, columns=columns, show="headings",
                                 height=12, selectmode=tk.BROWSE)
        self.tree.heading("name", text="文件名")
        self.tree.heading("type", text="类型")
        self.tree.heading("a", text="基准(A)")
        self.tree.heading("b", text="对照(B)")
        self.tree.heading("result", text="比对结果")
        self.tree.heading("summary", text="摘要")
        # 列对齐与固定宽度（窗口过窄时各列保持此宽度，自然产生横向溢出，通过底部滚动条查看）
        self.tree.column("name", stretch=False, anchor=tk.W, width=75)
        self.tree.column("type", stretch=False, anchor=tk.CENTER, width=50)
        self.tree.column("a", stretch=False, anchor=tk.CENTER, width=48)
        self.tree.column("b", stretch=False, anchor=tk.CENTER, width=48)
        self.tree.column("result", stretch=False, anchor=tk.CENTER, width=60)
        self.tree.column("summary", stretch=False, anchor=tk.W, width=200)

        self.tree.tag_configure("c_equal", foreground="#008000")
        self.tree.tag_configure("c_diff", foreground="#f44747")
        self.tree.tag_configure("c_only", foreground="#FF8C00")

        v_scroll = ttk.Scrollbar(tree_box, orient=tk.VERTICAL, command=self.tree.yview)
        h_scroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        h_scroll.pack(side=tk.BOTTOM, fill=tk.X)

        # 鼠标滚轮优化：悬停即可垂直滚动；Shift+滚轮横向滚动（返回 "break" 阻止重复触发）
        def _on_mousewheel(event):
            # 用位掩码判断 Shift（tk.SHIFT 在本环境可能不存在，直接抛 AttributeError）
            if event.state & 0x1:
                self.tree.xview_scroll(-1 * (event.delta // 120), "units")
            else:
                self.tree.yview_scroll(-1 * (event.delta // 120), "units")
            return "break"
        self.tree.bind("<MouseWheel>", _on_mousewheel)
        tree_box.bind("<MouseWheel>", _on_mousewheel)
        list_frame.bind("<MouseWheel>", _on_mousewheel)

        # 双击查看单文件差异详情
        self.tree.bind('<Double-Button-1>', self._on_row_double)

        # —— 底部总结 ——
        sum_frame = ttk.LabelFrame(parent, text="总结", padding="6")
        sum_frame.pack(side=tk.TOP, fill=tk.X, pady=(6, 0))
        self.summary_text = tk.Text(sum_frame, height=5, wrap=tk.WORD, state=tk.DISABLED,
                                   font=("Consolas", 9))
        self.summary_text.pack(fill=tk.X)

    # ====================== 文件夹上传区 ======================
    def _build_drop_zone(self, parent, col, title, var, side):
        """构建一个文件夹上传区：浏览按钮 + 可选拖入区。"""
        outer = ttk.LabelFrame(parent, text=title, padding="6")
        outer.grid(row=0, column=col, sticky="nsew", padx=4)

        # 路径显示 + 浏览按钮（最可靠的上传入口）
        row1 = ttk.Frame(outer)
        row1.pack(side=tk.TOP, fill=tk.X, pady=(0, 4))
        # 先放「浏览」按钮并固定其右侧空间，再让路径框填充剩余宽度，
        # 这样无论路径多长，「浏览」按钮始终可见
        ttk.Button(row1, text="浏览", width=8,
                  command=lambda s=side: self._browse(s)).pack(side=tk.RIGHT, padx=(4, 0))
        entry = ttk.Entry(row1, textvariable=var, state="readonly", width=40)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 拖入提示区（纯展示，不覆盖任何窗口过程）
        hint_text = "\U0001F4C1 把文件夹拖到这里，或点击「浏览」选择"
        if _HAS_DND and not self._dnd_initialized:
            # tkinterdnd2 只需初始化一次，对 root 调用
            try:
                TkinterDnD.require(self.gui.root)
                self._dnd_initialized = True
            except Exception:
                pass

        if _HAS_DND and self._dnd_initialized:
            # 使用 tkinterdnd2 注册拖入
            drop = tk.Label(
                outer, height=3, cursor="hand2",
                text=hint_text,
                justify=tk.CENTER, anchor="center",
                font=("Microsoft YaHei UI", 10)
            )
            drop.pack(side=tk.TOP, fill=tk.X)
            # 注册拖放
            drop.drop_target_register(DND_FILES)
            drop.dnd_bind('<<Drop>>', lambda e, s=side: self._on_tkdnd_drop(s, e))
        else:
            # 无拖入支持：显示静态提示
            extra = ""
            if not _HAS_DND:
                extra = "（安装 tkinterdnd2 后可启用拖入：pip install tkinterdnd2）"
            drop = tk.Label(
                outer, height=2,
                text=hint_text + ("\n" + extra if extra else ""),
                justify=tk.CENTER, anchor="center",
                font=("Microsoft YaHei UI", 9)
            )
            drop.pack(side=tk.TOP, fill=tk.X)

    def _on_tkdnd_drop(self, side, event):
        """tkinterdnd2 拖入回调。event.data 为空格分隔的路径（带花括号转义）。"""
        raw = event.data
        if not raw:
            return
        # 用 tk 内置 splitlist 处理花括号/空格转义
        paths = list(self.gui.root.tk.splitlist(raw))
        if not paths:
            return
        p = paths[0]
        if os.path.isfile(p):
            p = os.path.dirname(p)
        if os.path.isdir(p):
            self._set_folder(side, p)
            self.gui._log(f"[结果比对] 拖入{('基准' if side=='a' else '对照')}文件夹: {p}")

    def _browse(self, side):
        """点击「浏览」按钮，选择文件夹。"""
        initial = (self.folder_a.get() if side == "a" else self.folder_b.get()) or os.path.expanduser("~")
        path = filedialog.askdirectory(title="选择文件夹", initialdir=initial)
        if path:
            self._set_folder(side, path)

    def _set_folder(self, side, path):
        if side == "a":
            self.folder_a.set(path)
        else:
            self.folder_b.set(path)

    # ====================== 开始比对 ======================
    def _on_compare(self):
        if self._running:
            messagebox.showwarning("提示", "正在比对，请稍候")
            return
        fa, fb = self.folder_a.get().strip(), self.folder_b.get().strip()
        if not fa or not os.path.isdir(fa):
            messagebox.showwarning("提示", "请先选择/拖入基准文件夹 (A)")
            return
        if not fb or not os.path.isdir(fb):
            messagebox.showwarning("提示", "请先选择/拖入对照文件夹 (B)")
            return

        self.compare_btn.config(state=tk.DISABLED)
        self.export_btn.config(state=tk.DISABLED)
        self.tree.delete(*self.tree.get_children())
        self._set_summary("比对中...")
        self.gui._log(f"[结果比对] 开始比对\n  基准: {fa}\n  对照: {fb}")

        self._running = True
        args = (fa, fb, self.recursive.get(), self.ignore_order.get())
        threading.Thread(target=self._worker, args=args, daemon=True).start()

    def _worker(self, fa, fb, recursive, ignore_order):
        def progress(done, total, name):
            self.gui.root.after(0, lambda: self.status_label.config(
                text=f"比对中 {done}/{total}：{name}"))
        try:
            data = cmp.compare_folders(
                fa, fb, recursive=recursive, ignore_row_order=ignore_order,
                progress_cb=progress
            )
        except Exception as e:
            self.gui.root.after(0, self._on_error, e)
            return
        self.gui.root.after(0, self._on_done, data)

    def _on_error(self, exc):
        self._running = False
        self.compare_btn.config(state=tk.NORMAL)
        self.status_label.config(text="比对出错")
        messagebox.showerror("比对出错", str(exc))
        self.gui.logger.error(f"[结果比对] 出错: {exc}")

    def _on_done(self, data):
        self.data = data
        self._running = False
        self.compare_btn.config(state=tk.NORMAL)
        self.export_btn.config(state=tk.NORMAL)
        self._populate(data)
        stat = cmp.summarize(data)
        self.gui._log(
            f"[结果比对] 完成：共 {stat['total']} 个文件 | 一致 {stat['equal']} | "
            f"不一致 {stat['different']} | 仅基准 {stat['only_a']} | 仅对照 {stat['only_b']} | 错误 {stat['error']}"
        )

    # ====================== 结果呈现 ======================
    def _populate(self, data):
        self.tree.delete(*self.tree.get_children())
        for f in data["files"]:
            if not f["in_a"]:
                status, tag = "only_b", "c_only"
                result_text = "仅对照"
                summary = "基准文件夹中不存在"
                atxt, btxt = "\u2014", "\u2713"
            elif not f["in_b"]:
                status, tag = "only_a", "c_only"
                result_text = "仅基准"
                summary = "对照文件夹中不存在"
                atxt, btxt = "\u2713", "\u2014"
            else:
                r = f.get("result", {})
                if not r.get("ok", False):
                    status, tag = "error", "c_diff"
                    result_text = "错误"
                    summary = r.get("error", "")
                    atxt, btxt = "\u2713", "\u2713"
                elif r.get("all_equal"):
                    status, tag = "equal", "c_equal"
                    result_text = "一致"
                    summary = "内容完全相同"
                    atxt, btxt = "\u2713", "\u2713"
                else:
                    status, tag = "diff", "c_diff"
                    result_text = "不一致"
                    summary = self._file_summary(f["name"], r, data["ignore_row_order"])
                    atxt, btxt = "\u2713", "\u2713"

            ext = os.path.splitext(f["name"])[1].lower()
            typ = "Excel" if ext in cmp.EXCEL_EXTS else (ext[1:].upper() if ext else "文件")
            self.tree.insert(
                "", tk.END, iid=f["name"],
                values=(f["name"], typ, atxt, btxt, result_text, summary),
                tags=(tag,),
            )

        # 总结
        stat = cmp.summarize(data)
        lines = [
            f"文件总数: {stat['total']}    "
            f"一致: {stat['equal']}    "
            f"不一致: {stat['different']}",
            f"仅基准(A): {stat['only_a']}    "
            f"仅对照(B): {stat['only_b']}    "
            f"错误: {stat['error']}",
        ]
        self._set_summary("\n".join(lines))
        self.status_label.config(text="比对完成")

    def _file_summary(self, name, r, ignore_order):
        if r.get("type") != "excel":
            return f"大小 A={r.get('size_a')} / B={r.get('size_b')}，内容不同"
        if ignore_order:
            only_a = sum(len(s.get("removed", [])) for s in r.get("sheets", []))
            only_b = sum(len(s.get("added", [])) for s in r.get("sheets", []))
            s = f"新增{only_b}行 / 减少{only_a}行"
        else:
            s = f"{sum(len(s.get('cell_diffs', [])) for s in r.get('sheets', []))}处单元格差异"
        extra = []
        if r.get("only_a"):
            extra.append("缺工作表:" + "/".join(map(str, r["only_a"])))
        if r.get("only_b"):
            extra.append("多工作表:" + "/".join(map(str, r["only_b"])))
        if extra:
            s += "\uff1b" + "\uff0c".join(extra)
        return s

    def _set_summary(self, text):
        self.summary_text.config(state=tk.NORMAL)
        self.summary_text.delete(1.0, tk.END)
        self.summary_text.insert(1.0, text)
        self.summary_text.config(state=tk.DISABLED)

    # ====================== 单文件明细 ======================
    def _on_row_double(self, event):
        sel = self.tree.selection()
        if not sel or self.data is None:
            return
        name = sel[0]
        f = next((x for x in self.data["files"] if x["name"] == name), None)
        if not f or not (f["in_a"] and f["in_b"]):
            return
        r = f.get("result", {})
        if not r.get("ok") or r.get("all_equal"):
            return
        self._show_detail(name, r)

    def _show_detail(self, name, r):
        win = tk.Toplevel(self.parent)
        win.title(f"比对明细 - {name}")
        win.geometry("760x520")
        txt = tk.Text(win, wrap=tk.WORD, font=("Consolas", 9))
        txt.pack(fill=tk.BOTH, expand=True)
        txt.insert(1.0, self._detail_text(name, r))
        txt.config(state=tk.DISABLED)
        ttk.Button(win, text="关闭", command=win.destroy).pack(pady=4)

    def _detail_text(self, name, r):
        L = [f"文件: {name}", ""]
        if r.get("type") == "excel":
            if r.get("only_a"):
                L.append(f"仅基准存在的工作表: {', '.join(map(str, r['only_a']))}")
            if r.get("only_b"):
                L.append(f"仅对照存在的工作表: {', '.join(map(str, r['only_b']))}")
            for s in r.get("sheets", []):
                if s.get("equal"):
                    continue
                L.append(f"\u25cf 工作表\u300c{s['name']}\u300d (基准 {s['rows_a']} 行 / 对照 {s['rows_b']} 行)")
                if s.get("mode") == "ignore_order":
                    removed, rtr = cmp._short_rows(s.get("removed", []))
                    added, atr = cmp._short_rows(s.get("added", []))
                    L.append(f"  减少的行 ({len(s.get('removed', []))}):")
                    for row in removed:
                        L.append("    - " + " | ".join(row))
                    if rtr:
                        L.append("    ...(已截断)")
                    L.append(f"  新增的行 ({len(s.get('added', []))}):")
                    for row in added:
                        L.append("    + " + " | ".join(row))
                    if atr:
                        L.append("    ...(已截断)")
                else:
                    diffs = s.get("cell_diffs", [])
                    shown, trunc = cmp._short_rows(diffs)
                    L.append(f"  单元格差异 ({len(diffs)}):")
                    for (rr, cc, va, vb) in shown:
                        if cc is None:
                            L.append(f"    行{rr + 1}: 基准={'无' if va is None else va} / 对照={'无' if vb is None else vb}")
                        else:
                            L.append(f"    行{rr + 1}列{cc + 1}: 基准={va!r} / 对照={vb!r}")
                    if trunc:
                        L.append("    ...(已截断)")
                L.append("")
        else:
            L.append(f"大小: 基准 {r.get('size_a')} / 对照 {r.get('size_b')}\uff1bMD5 不一致")
        return "\n".join(L)

    # ====================== 导出报告 ======================
    def _on_export(self):
        if self.data is None:
            messagebox.showinfo("提示", "请先完成一次比对")
            return
        default = os.path.join(
            os.path.expanduser("~"), "Desktop",
            f"比对报告_{self.data['generated_at'].replace(':', '').replace(' ', '_')}.txt"
        )
        path = filedialog.asksaveasfilename(
            title="导出比对报告",
            initialfile=os.path.basename(default),
            initialdir=os.path.dirname(default),
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt"), ("All files", "*.*")]
        )
        if not path:
            return
        try:
            cmp.export_report(self.data, path)
            self.gui._log(f"[结果比对] 报告已导出: {path}")
            messagebox.showinfo("成功", f"报告已导出:\n{path}")
            try:
                os.startfile(path)
            except Exception:
                pass
        except Exception as e:
            messagebox.showerror("导出失败", str(e))
