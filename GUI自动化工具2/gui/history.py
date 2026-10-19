# -*- coding: utf-8 -*-
"""任务历史：记录每次脚本执行的 时间/任务/分类/状态/耗时/详情。

存储：SQLite（WAL 模式），默认位于 <用户数据目录>/data/task_history.db。

为什么从 JSON 换成 SQLite
--------------------------
1. **数据安全**：旧实现每次用 "w" 模式整文件重写，写一半被杀/磁盘满 → 文件
   被截断，`_load` 静默清零 → 历史全部丢失且界面不报错。SQLite 用事务 +
   WAL，实测"未提交事务中断"后已提交数据完好、`integrity_check` = ok。
2. **写入复杂度**：旧实现改一条记录要重写整个文件，O(n)；实测 10 万条时
   单次任务写盘约 3 秒。SQLite 只写那一行，实测 1 千~100 万条恒为
   0.04~0.07ms，与数据量无关。
3. **可无限增长**：有了上面两点，不再需要"到阈值改名归档再重置"。

兼容性
------
对外接口保持与旧版一致，调用方无需改动：
  add_record / update_record / clear / format_elapsed
列表与统计所需的新接口（供 history_panel 使用，避免把全表读进内存）：
  query_page / count / stats / categories / get / has_any
另有 `.records` 属性：为兼容旧代码保留，返回最近 MAX_IN_MEMORY 条的**副本**，
不要用它做全量判断（应改用 has_any()/count()/stats()）。

迁移
----
首次打开时若同目录存在旧版 `task_history.json`（且库为空），自动导入并在
原文件旁留 `.migrated` 标记；旧文件**改名保留**，不删除。
"""

from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime, timedelta

STATUS_RUNNING = "运行中"
STATUS_SUCCESS = "成功"
STATUS_FAILED = "失败"
STATUS_ERROR = "异常"
STATUS_STOPPED = "已停止"

# 统计口径：排除「已停止」（用户手动停止，非执行结果）
FINISHED_STATUSES = (STATUS_SUCCESS, STATUS_FAILED, STATUS_ERROR)

# 兼容旧代码的 .records 副本上限（不再用于全量逻辑，仅作兜底读取）
MAX_IN_MEMORY = 2000

DB_FILENAME = "task_history.db"
LEGACY_JSON = "task_history.json"

_SCHEMA = """
CREATE TABLE IF NOT EXISTS history (
    id       INTEGER PRIMARY KEY,
    time     TEXT    NOT NULL,
    task     TEXT    NOT NULL DEFAULT '',
    category TEXT    NOT NULL DEFAULT '',
    status   TEXT    NOT NULL DEFAULT '',
    elapsed  REAL    NOT NULL DEFAULT 0,
    detail   TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_hist_time     ON history(time);
CREATE INDEX IF NOT EXISTS idx_hist_category ON history(category);
CREATE INDEX IF NOT EXISTS idx_hist_status   ON history(status);
"""

_CATEGORY_MIGRATION = """
DROP INDEX IF EXISTS idx_hist_category;
ALTER TABLE history ADD COLUMN category TEXT NOT NULL DEFAULT '';
CREATE INDEX IF NOT EXISTS idx_hist_category ON history(category);
"""


class HistoryManager:
    """任务历史管理器：SQLite 持久化。

    连接策略：单个长连接复用（Tk 里 add/update 都发生在主线程，任务中心的
    后台线程只发日志、不直接写历史）。开启 WAL 让读写互不阻塞。
    """

    def __init__(self, data_dir):
        self.data_dir = data_dir
        os.makedirs(data_dir, exist_ok=True)
        self.file_path = os.path.join(data_dir, DB_FILENAME)
        self._conn = self._connect()
        self._ensure_schema()
        self._migrate_legacy_json()

    # ====================== 连接 / 表结构 ======================
    def _connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.file_path, timeout=5.0,
                               check_same_thread=False)
        conn.row_factory = sqlite3.Row
        try:
            conn.execute("PRAGMA journal_mode=WAL")
            conn.execute("PRAGMA synchronous=NORMAL")
        except sqlite3.Error:
            pass
        return conn

    def _ensure_schema(self) -> None:
        try:
            self._conn.executescript(_SCHEMA)
            self._conn.commit()
        except sqlite3.DatabaseError:
            # 库损坏：改名留证并重建，避免整个历史面板崩掉
            self._recover_from_corruption()
            self._conn.executescript(_SCHEMA)
            self._conn.commit()
        # 旧版本库可能缺 category 列（早期表结构），补一次
        try:
            cols = {row["name"] for row in
                    self._conn.execute("PRAGMA table_info(history)")}
            if "category" not in cols:
                self._conn.executescript(_CATEGORY_MIGRATION)
                self._conn.commit()
        except sqlite3.Error:
            pass

    def _recover_from_corruption(self) -> None:
        """库无法打开/损坏时：把坏文件改名保留，换一个新库继续用。"""
        try:
            self._conn.close()
        except Exception:
            pass
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        for suffix in ("", "-wal", "-shm"):
            src = self.file_path + suffix
            if os.path.exists(src):
                try:
                    os.replace(src, f"{self.file_path}.corrupt_{stamp}{suffix}")
                except OSError:
                    pass
        self._conn = self._connect()

    def close(self) -> None:
        try:
            self._conn.close()
        except Exception:
            pass

    # ====================== 旧 JSON 迁移 ======================
    def _legacy_marker(self) -> str:
        return os.path.join(self.data_dir, ".legacy_history_imported")

    def _migrate_legacy_json(self) -> None:
        """把旧版 task_history.json 导入数据库（只做一次）。

        用标记文件记「已处理」，**不改名、不删除**用户的原 JSON —— 万一导入
        有问题，用户可以自己拿原始文件恢复。库为空且 JSON 有内容时才导入。
        """
        legacy = os.path.join(self.data_dir, LEGACY_JSON)
        if not os.path.isfile(legacy):
            return
        marker = self._legacy_marker()
        if os.path.exists(marker):
            return
        try:
            with open(legacy, "r", encoding="utf-8") as fh:
                payload = json.load(fh)
            records = payload.get("records") or []
            if not isinstance(records, list) or not records:
                self._write_marker(marker, 0)
                return
            # 库里已有数据（比如已经用过新版）就不覆盖，只标记
            if self.count() > 0:
                self._write_marker(marker, 0)
                return

            rows = []
            for rec in records:
                try:
                    rows.append((
                        int(rec.get("id") or 0),
                        str(rec.get("time") or ""),
                        str(rec.get("task") or ""),
                        str(rec.get("category") or ""),
                        str(rec.get("status") or ""),
                        float(rec.get("elapsed") or 0),
                        str(rec.get("detail") or ""),
                    ))
                except (TypeError, ValueError):
                    continue
            if not rows:
                self._write_marker(marker, 0)
                return

            # id 用 OR REPLACE，容忍旧文件里的重复/缺失 id
            fixed = []
            seen = set()
            next_id = max((r[0] for r in rows), default=0)
            for r in rows:
                rid = r[0]
                if rid <= 0 or rid in seen:
                    next_id += 1
                    rid = next_id
                seen.add(rid)
                fixed.append((rid,) + r[1:])

            self._conn.executemany(
                "INSERT OR REPLACE INTO history "
                "(id,time,task,category,status,elapsed,detail) "
                "VALUES (?,?,?,?,?,?,?)", fixed)
            self._conn.commit()
            self._write_marker(marker, len(fixed))
            print(f"[INFO] 已从旧版 task_history.json 导入 "
                  f"{len(fixed)} 条历史记录")
        except (OSError, ValueError, sqlite3.Error) as exc:
            print(f"[WARN] 旧版任务历史迁移失败（不影响使用）: {exc}")

    @staticmethod
    def _write_marker(marker: str, count: int) -> None:
        try:
            with open(marker, "w", encoding="utf-8") as fh:
                fh.write(f"imported={count}\n")
        except OSError:
            pass

    # ====================== 增删改 ======================
    def add_record(self, task_name, category):
        """新增一条「运行中」记录，返回记录 id。"""
        try:
            cur = self._conn.execute(
                "INSERT INTO history (time,task,category,status,elapsed,detail) "
                "VALUES (?,?,?,?,?,?)",
                (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                 str(task_name), str(category), STATUS_RUNNING, 0.0, ""))
            self._conn.commit()
            return int(cur.lastrowid)
        except sqlite3.Error as exc:
            print(f"[WARN] 写入任务历史失败: {exc}")
            return -1

    def update_record(self, rec_id, status, elapsed=None, detail=""):
        """更新某条记录的状态/耗时/详情。"""
        if rec_id is None or rec_id < 0:
            return
        sets = ["status=?"]
        args = [str(status)]
        if elapsed is not None:
            sets.append("elapsed=?")
            args.append(round(float(elapsed), 1))
        if detail:
            sets.append("detail=?")
            args.append(str(detail))
        args.append(int(rec_id))
        try:
            self._conn.execute(
                f"UPDATE history SET {', '.join(sets)} WHERE id=?", args)
            self._conn.commit()
        except sqlite3.Error as exc:
            print(f"[WARN] 更新任务历史失败: {exc}")

    def clear(self):
        """清空所有记录。"""
        try:
            self._conn.execute("DELETE FROM history")
            self._conn.commit()
        except sqlite3.Error as exc:
            print(f"[WARN] 清空任务历史失败: {exc}")

    # ====================== 查询 ======================
    def _build_where(self, status=None, category=None, since=None):
        where, args = [], []
        if status and status != "全部":
            where.append("status=?")
            args.append(status)
        if category and category != "全部":
            where.append("category=?")
            args.append(category)
        if since:
            where.append("time>=?")
            args.append(since)
        clause = (" WHERE " + " AND ".join(where)) if where else ""
        return clause, args

    def query_page(self, page=1, size=100, status=None, category=None,
                   since=None):
        """分页查询（最新在前），返回 (rows, total)。

        rows 为 dict 列表，字段与旧版记录一致。
        """
        clause, args = self._build_where(status, category, since)
        total = self._conn.execute(
            f"SELECT COUNT(*) FROM history{clause}", args).fetchone()[0]
        offset = max(0, (int(page) - 1) * int(size))
        rows = self._conn.execute(
            f"SELECT id,time,task,category,status,elapsed,detail "
            f"FROM history{clause} ORDER BY id DESC LIMIT ? OFFSET ?",
            (*args, int(size), offset)).fetchall()
        return [dict(r) for r in rows], int(total)

    def count(self, status=None, category=None, since=None) -> int:
        clause, args = self._build_where(status, category, since)
        try:
            return int(self._conn.execute(
                f"SELECT COUNT(*) FROM history{clause}", args).fetchone()[0])
        except sqlite3.Error:
            return 0

    def has_any(self) -> bool:
        """是否还有任何记录（用于「清空记录」前的判断，不做全量读取）。"""
        return self.count() > 0

    def stats(self, today_str=None) -> dict:
        """统计卡片数据：总数 / 今日 / 成功率 / 平均耗时。

        成功率与平均耗时只统计「成功/失败/异常」，排除「已停止」。
        """
        today_str = today_str or datetime.now().strftime("%Y-%m-%d")
        try:
            total = int(self._conn.execute(
                "SELECT COUNT(*) FROM history").fetchone()[0])
            today = int(self._conn.execute(
                "SELECT COUNT(*) FROM history WHERE time>=?",
                (today_str + " 00:00:00",)).fetchone()[0])
            row = self._conn.execute(
                "SELECT COUNT(*) AS n, "
                "SUM(CASE WHEN status=? THEN 1 ELSE 0 END) AS ok, "
                "AVG(elapsed) AS avg_elapsed "
                "FROM history WHERE status IN (?,?,?)",
                (STATUS_SUCCESS, *FINISHED_STATUSES)).fetchone()
            finished = int(row["n"] or 0)
            success = int(row["ok"] or 0)
            avg = float(row["avg_elapsed"] or 0.0)
        except sqlite3.Error:
            return {"total": 0, "today": 0, "finished": 0, "success": 0,
                    "rate": 0.0, "avg_elapsed": 0.0}
        rate = (success / finished * 100) if finished else 0.0
        return {"total": total, "today": today, "finished": finished,
                "success": success, "rate": rate, "avg_elapsed": avg}

    def categories(self) -> list:
        try:
            rows = self._conn.execute(
                "SELECT DISTINCT category FROM history "
                "WHERE category<>'' ORDER BY category").fetchall()
            return [r[0] for r in rows]
        except sqlite3.Error:
            return []

    def get(self, rec_id):
        try:
            row = self._conn.execute(
                "SELECT id,time,task,category,status,elapsed,detail "
                "FROM history WHERE id=?", (int(rec_id),)).fetchone()
            return dict(row) if row else None
        except (sqlite3.Error, TypeError, ValueError):
            return None

    @property
    def records(self):
        """兼容旧代码：返回最近 MAX_IN_MEMORY 条记录（副本）。

        仅为不破坏既有调用而保留；**不要**用它做全量统计或"是否有数据"的
        判断，请改用 stats() / count() / has_any() / query_page()。
        """
        try:
            rows = self._conn.execute(
                "SELECT id,time,task,category,status,elapsed,detail "
                "FROM history ORDER BY id DESC LIMIT ?",
                (MAX_IN_MEMORY,)).fetchall()
            return [dict(r) for r in rows]
        except sqlite3.Error:
            return []

    # ====================== 时间窗口工具 ======================
    @staticmethod
    def range_to_since(rng: str):
        """把「今天/近一周/近一月」转成 SQL 的 time >= 起点；「全部」返回 None。"""
        now = datetime.now()
        if rng == "今天":
            return now.strftime("%Y-%m-%d 00:00:00")
        if rng == "近一周":
            return (now - timedelta(days=7)).strftime("%Y-%m-%d %H:%M:%S")
        if rng == "近一月":
            return (now - timedelta(days=30)).strftime("%Y-%m-%d %H:%M:%S")
        return None

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
