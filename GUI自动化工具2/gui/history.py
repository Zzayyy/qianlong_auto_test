# -*- coding: utf-8 -*-
"""任务历史：记录每次脚本执行的 时间/任务/分类/状态/耗时，并持久化到 JSON

记录字段：
  - id        自增主键（用于双击查看详情时定位记录）
  - time      开始执行的时间（YYYY-MM-DD HH:MM:SS）
  - task      脚本名
  - category  功能分类
  - status    状态：运行中 / 成功 / 失败 / 异常 / 已停止
  - elapsed   耗时（秒，浮点）
  - detail    失败/异常时的附加信息
"""

import os
import json
from datetime import datetime

STATUS_RUNNING = "运行中"
STATUS_SUCCESS = "成功"
STATUS_FAILED = "失败"
STATUS_ERROR = "异常"
STATUS_STOPPED = "已停止"

MAX_RECORDS = 200  # 最多保留最近 200 条


class HistoryManager:
    """任务历史管理器：内存存储 + JSON 持久化"""

    def __init__(self, log_dir):
        self.log_dir = log_dir
        self.file_path = os.path.join(log_dir, "task_history.json")
        self.records = []  # 最新记录在前
        self._seq = 0
        self._load()

    # ====================== 持久化 ======================
    def _load(self):
        if os.path.exists(self.file_path):
            try:
                with open(self.file_path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    self.records = data.get("records", [])
                    self._seq = data.get("seq", 0)
            except Exception:
                self.records = []
                self._seq = 0

    def _save(self):
        try:
            with open(self.file_path, "w", encoding="utf-8") as f:
                json.dump(
                    {"seq": self._seq, "records": self.records},
                    f, ensure_ascii=False, indent=2
                )
        except Exception:
            pass

    # ====================== 增删改 ======================
    def add_record(self, task_name, category):
        """新增一条「运行中」记录，返回记录 id"""
        self._seq += 1
        rec = {
            "id": self._seq,
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "task": task_name,
            "category": category,
            "status": STATUS_RUNNING,
            "elapsed": 0.0,
            "detail": "",
        }
        self.records.insert(0, rec)  # 最新在前
        if len(self.records) > MAX_RECORDS:
            self.records = self.records[:MAX_RECORDS]
        self._save()
        return rec["id"]

    def update_record(self, rec_id, status, elapsed=None, detail=""):
        """更新某条记录的状态/耗时/详情"""
        for rec in self.records:
            if rec["id"] == rec_id:
                rec["status"] = status
                if elapsed is not None:
                    rec["elapsed"] = round(float(elapsed), 1)
                if detail:
                    rec["detail"] = detail
                break
        self._save()

    def clear(self):
        """清空所有记录"""
        self.records = []
        self._save()

    # ====================== 工具 ======================
    @staticmethod
    def format_elapsed(secs):
        """把秒数格式化为友好字符串：12.3s / 1分23秒"""
        secs = float(secs or 0)
        if secs < 60:
            return f"{secs:.1f}s"
        minutes = int(secs // 60)
        seconds = int(secs % 60)
        return f"{minutes}分{seconds}秒"
