# -*- coding: utf-8 -*-
"""
定时任务调度引擎
"""
import os as _os, json, time, threading, subprocess, sys
from datetime import datetime, timedelta
from .task import Task
from config import IS_FROZEN, PROJECT_ROOT

WEEKDAY_NAMES = ["周一","周二","周三","周四","周五","周六","周日"]

# 过期任务保护阈值：next_run 超过此秒数则跳过本轮执行，等待下一次正常调度窗口
CATCHUP_THRESHOLD = 300  # 5 分钟

def compute_next_run(sched):
    stype = sched["schedule_type"]
    time_str = sched.get("time","09:00")
    try:
        hour, minute = map(int, time_str.split(":"))
    except (ValueError,AttributeError):
        return ""
    now = datetime.now()
    base = now.replace(hour=hour, minute=minute, second=0, microsecond=0)
    if stype == "once":
        ds = sched.get("scheduled_date","")
        if not ds:
            return ""
        try:
            dt = datetime.strptime(f"{ds} {time_str}","%Y-%m-%d %H:%M")
        except ValueError:
            return ""
        return "" if dt <= now else dt.strftime("%Y-%m-%d %H:%M")
    if stype == "daily":
        c = base
        if c <= now:
            c += timedelta(days=1)
        return c.strftime("%Y-%m-%d %H:%M")
    if stype == "weekly":
        wds = sched.get("weekdays",[])
        if not wds:
            return ""
        for offset in range(8):
            c = base + timedelta(days=offset)
            if c <= now:
                continue
            if c.weekday() in wds:
                return c.strftime("%Y-%m-%d %H:%M")
        return ""
    if stype == "monthly":
        day = max(1, min(28, sched.get("month_day",1)))
        c = base.replace(day=day)
        if c <= now:
            m = c.month + 1
            y = c.year
            if m > 12:
                m = 1; y += 1
            c = c.replace(year=y, month=m)
        return c.strftime("%Y-%m-%d %H:%M")
    return ""

def format_schedule_desc(sched):
    st = sched["schedule_type"]
    tm = sched.get("time","09:00")
    if st == "once":
        d = sched.get("scheduled_date","")
        return f"一次 {d} {tm}" if d else f"一次 {tm}"
    if st == "daily":
        return f"每天 {tm}"
    if st == "weekly":
        wds = sched.get("weekdays",[])
        ns = [WEEKDAY_NAMES[w] for w in wds if 0 <= w < 7]
        return f"每周 {','.join(ns)} {tm}" if ns else f"每周 {tm}"
    if st == "monthly":
        return f"每月第{sched.get('month_day',1)}天 {tm}"
    return st

class TaskScheduler:
    def __init__(self, gui):
        self.gui = gui
        self._tasks = []
        self._seq = 0
        self._running = False
        self._executing = False
        self._lock = threading.Lock()
        self.file_path = _os.path.join(gui.log_dir, "scheduled_tasks.json")
        self._view = None
        self._load()
        self._start()
        gui.logger.info(f"定时任务调度器已启动，共 {len(self._tasks)} 个任务")

    @property
    def tasks(self):
        return list(self._tasks)

    def _load(self):
        if _os.path.exists(self.file_path):
            try:
                with open(self.file_path,"r",encoding="utf-8") as f:
                    d = json.load(f)
                self._tasks = d.get("tasks",[])
                self._seq = d.get("seq",0)
            except Exception:
                self._tasks = []; self._seq = 0

    def _save(self):
        # 备份旧文件，防止写入失败导致数据丢失
        if _os.path.exists(self.file_path):
            try:
                import shutil
                shutil.copy2(self.file_path, self.file_path + ".bak")
            except Exception:
                pass
        try:
            with open(self.file_path,"w",encoding="utf-8") as f:
                json.dump({"seq":self._seq,"tasks":self._tasks},f,ensure_ascii=False,indent=2)
        except Exception as e:
            self.gui.logger.error(f"定时任务保存失败: {e}")

    def add_task(self, sched):
        with self._lock:
            self._seq += 1
            sched["id"] = self._seq
            sched["enabled"] = sched.get("enabled", True)
            sched["status"] = "等待"
            sched["last_run"] = ""
            sched["total_runs"] = 0
            sched["completed_runs"] = 0
            sched["next_run"] = compute_next_run(sched)
            self._tasks.append(sched)
            self._save()
        return self._seq

    def update_task(self, task_id, updates):
        with self._lock:
            for t in self._tasks:
                if t["id"] == task_id:
                    t.update(updates)
                    for k in ("schedule_type","time","scheduled_date","weekdays","month_day"):
                        if k in updates:
                            t["next_run"] = compute_next_run(t)
                            break
                    self._save(); return True
        return False

    def delete_task(self, task_id):
        with self._lock:
            for i,t in enumerate(self._tasks):
                if t["id"] == task_id:
                    self._tasks.pop(i); self._save(); return True
        return False

    def toggle_enabled(self, task_id):
        with self._lock:
            for t in self._tasks:
                if t["id"] == task_id:
                    t["enabled"] = not t["enabled"]
                    if t["enabled"]:
                        t["next_run"] = compute_next_run(t)
                    t["status"] = "等待" if t["enabled"] else "已禁用"
                    self._save(); return t["enabled"]
        return False

    def get_task(self, task_id):
        for t in self._tasks:
            if t["id"] == task_id:
                return dict(t)
        return {}

    def _start(self):
        self._running = True
        threading.Thread(target=self._scheduler_loop, daemon=True).start()

    def stop(self):
        self._running = False


    def _scheduler_loop(self):
        while self._running:
            try:
                if not self._executing:
                    now = datetime.now()
                    due = []
                    skipped = []
                    with self._lock:
                        for t in self._tasks:
                            if not t.get("enabled", False): continue
                            nr = t.get("next_run","")
                            if not nr: continue
                            try:
                                dt = datetime.strptime(nr,"%Y-%m-%d %H:%M")
                            except ValueError: continue
                            if now >= dt:
                                gap = (now - dt).total_seconds()
                                if gap > CATCHUP_THRESHOLD:
                                    t["next_run"] = compute_next_run(t)
                                    skipped.append(t.get("name",""))
                                else:
                                    due.append(t)
                    if skipped:
                        self._save()
                if skipped:
                    self._log(f"跳过 {len(skipped)} 个过期任务（距计划时间超过5分钟）: {', '.join(skipped)}")
                for sched in due:
                    if not self._running: break
                    self._execute(sched)
            except Exception:
                pass
            time.sleep(5)

    def _execute(self, sched):
        if self._executing:
            self._log(f"上一个任务仍在执行，跳过: {sched.get('name')}")
            return
        self._executing = True
        self._set_running_flag(sched["id"], True)
        self._log("\n" + "=" * 60)
        self._log(f"开始执行: {sched.get('name')} (ID={sched['id']})")
        self._log(f"目标: {sched.get('script_name') or sched.get('group_name','')}")
        self._log("=" * 60)
        if sched.get("target_type","script") == "group":
            self._execute_group(sched)
        else:
            self._execute_script(sched)

    def _execute_script(self, sched):
        try:
            sp = sched.get("script_path","")
            if not _os.path.exists(sp):
                self._on_error(sched, f"脚本文件不存在: {sp}"); return
            task = Task({"name":sched.get("script_name",""),"path":sp,"query_key":sched.get("query_key","")},sched.get("category",""),sched.get("params",{}))
            env = task.build_env(PROJECT_ROOT)
            env["PYTHONIOENCODING"]="utf-8"; env["PYTHONUTF8"]="1"
            cmd = [sys.executable,"--_run_script",sp] if IS_FROZEN else [sys.executable,"-u",sp]
            self._log(f"命令: {cmd[0]}")
            proc = subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8",errors="replace",creationflags=subprocess.CREATE_NO_WINDOW,env=env)
            start = time.time()
            try:
                stdout_data, _ = proc.communicate(timeout=600)
                if stdout_data:
                    for l in stdout_data.splitlines():
                        l = l.rstrip()
                        if l: self._log(l)
            except subprocess.TimeoutExpired:
                proc.kill()
                proc.wait()
                raise TimeoutError(f"脚本执行超时 (600s)")
            rc = proc.returncode; elapsed = time.time()-start
            self._log(f"{'exec OK' if rc==0 else 'exec FAIL'}, {elapsed:.1f}s")
            self._on_finish(sched, rc, elapsed)
        except Exception as e:
            self._on_error(sched, str(e))
        finally:
            if self._executing:
                self._executing = False
                self._set_running_flag(sched["id"], False)

    def _execute_group(self, sched):
        gn = sched.get("group_name","")
        gp = _os.path.normpath(_os.path.join(self.gui.log_dir,"task_groups.json"))
        if not _os.path.exists(gp):
            self._on_error(sched,f"任务编队文件不存在: {gp}"); return
        try:
            with open(gp,"r",encoding="utf-8") as f: data = json.load(f)
        except Exception as e:
            self._on_error(sched,f"读取任务编队失败: {e}"); return
        groups = data if isinstance(data,list) else data.get("groups",[])
        g = next((x for x in groups if x.get("name")==gn),None)
        if not g: self._on_error(sched,f"未找到任务编队: {gn}"); return
        gts = g.get("tasks",[])
        if not gts: self._on_error(sched,f"编队 '{gn}' 中没有脚本"); return
        self._log(f"编队 '{gn}' 共 {len(gts)} 个脚本，顺序执行")
        ok = 0; tel = 0.0
        for idx, gt in enumerate(gts):
            sp = gt.get("script_path","")
            if not _os.path.exists(sp): self._log(f"  [{idx+1}/{len(gts)}] 脚本不存在: {sp}"); continue
            task = Task({"name":gt.get("script_name",""),"path":sp,"query_key":gt.get("query_key","")},gt.get("category",""),gt.get("params",{}))
            env = task.build_env(PROJECT_ROOT); env["PYTHONIOENCODING"]="utf-8"; env["PYTHONUTF8"]="1"
            cmd = [sys.executable,"--_run_script",sp] if IS_FROZEN else [sys.executable,"-u",sp]
            self._log('[' + str(idx+1) + '/' + str(len(gts)) + '] \u5f00\u59cb: ' + gt.get('script_name',''))
            try:
                proc = subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding="utf-8",errors="replace",creationflags=subprocess.CREATE_NO_WINDOW,env=env)
                start = time.time()
                try:
                    stdout_data2, _ = proc.communicate(timeout=600)
                    if stdout_data2:
                        for l2 in stdout_data2.splitlines():
                            l2 = l2.rstrip()
                            if l2: self._log(l2)
                except subprocess.TimeoutExpired:
                    proc.kill()
                    proc.wait()
                    self._log(f"[{idx+1}/{len(gts)}] 脚本执行超时 (600s)")
                    rc = -1
                else:
                    rc = proc.returncode
                el = time.time()-start; tel += el
                if rc==0: ok+=1; self._log('[' + str(idx+1) + '/' + str(len(gts)) + '] \u6210\u529f (' + str(round(el,1)) + 's)')
                else: self._log('[' + str(idx+1) + '/' + str(len(gts)) + '] \u5931\u8d25\uff0c\u9000\u51fa\u7801: ' + str(rc))
            except Exception as e:
                self._log('[' + str(idx+1) + '/' + str(len(gts)) + '] \u5f02\u5e38: ' + str(e))
        self._log('\n\u7f16\u961f\u6267\u884c\u5b8c\u6bd5: ' + str(ok) + '/' + str(len(gts)) + ' \u6210\u529f\uff0c\u603b\u8017\u65f6 ' + str(round(tel,1)) + 's')
        try:
            self._on_finish(sched, 0 if ok==len(gts) else 1, tel)
        except Exception as e:
            self._on_error(sched, str(e))
        finally:
            if self._executing:
                self._executing = False
                self._set_running_flag(sched["id"], False)

    def _on_finish(self, sched, rc, elapsed):
        ns = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._lock:
            for t in self._tasks:
                if t["id"]==sched["id"]:
                    t["last_run"]=ns; t["total_runs"]=(t.get("total_runs",0)or 0)+1
                    if rc==0: t["completed_runs"]=(t.get("completed_runs",0)or 0)+1
                    if t.get("schedule_type")=="once": t["enabled"]=False; t["status"]="已完成"; t["next_run"]=""
                    else: t["next_run"]=compute_next_run(t); t["status"]="等待"
                    self._save(); break
        self._executing=False; self._set_running_flag(sched["id"],False)
        st = "成功" if rc==0 else "失败"
        self.gui.root.after(0,lambda: self.gui._set_status("定时任务: " + sched.get("name","") + " " + st))
        self.gui.root.after(0,lambda: self.gui._refresh_history())
        if self._view: self.gui.root.after(0,self._view._refresh)
        self._log("完成: " + sched.get("name","") + " " + st + "，下次: " + sched.get("next_run","无"))
    def _on_error(self, sched, err):
        ns = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with self._lock:
            for t in self._tasks:
                if t["id"]==sched["id"]:
                    t["last_run"]=ns; t["total_runs"]=(t.get("total_runs",0)or 0)+1
                    if t.get("schedule_type")=="once": t["enabled"]=False; t["status"]="出错"; t["next_run"]=""
                    else: t["next_run"]=compute_next_run(t); t["status"]="等待"
                    self._save(); break
        self._executing=False; self._set_running_flag(sched["id"],False)
        self._log(f"异常: {sched.get('name')} - {err}")
        if self._view: self.gui.root.after(0,self._view._refresh)

    def _set_running_flag(self, tid, running):
        with self._lock:
            for t in self._tasks:
                if t["id"]==tid: t["status"]="运行中" if running else ("等待" if t.get("enabled") else "已禁用"); self._save(); break
        if self._view: self.gui.root.after(0,self._view._refresh)

    def _log(self, msg):
        self.gui.root.after(0,self.gui._log,f"[调度] {msg}")

    def bind_view(self, view):
        self._view = view
