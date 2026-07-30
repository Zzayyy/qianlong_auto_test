#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
记忆目录双向同步监视器  (.codebuddy/memory  <->  .workbuddy/memory)

功能：
  - 轮询监听两个 memory 目录（默认 1.5 秒一次，零第三方依赖）。
  - 任意一侧文件新增/修改，即字节级复制到另一侧（保留原换行符，不重新编码）。
  - 自带防回环：自己写入触发的变更不会二次同步回去。
  - 两侧同一文件同时改动时，以 mtime 较新者为准，并记录冲突警告。
  - 默认「不传播删除」：某侧删除文件只在日志提示，避免误删对侧内容。
  - 支持 --once 单次同步后退出；--dry-run 只预览不落盘；--interval 调间隔。

用法：
  python sync_memory_watch.py              # 常驻监视
  python sync_memory_watch.py --once       # 跑一次同步就退出
  python sync_memory_watch.py --dry-run    # 预览会做什么
退出：Ctrl+C
"""
import argparse
import os
import sys
import time
import threading

ROOT = os.path.dirname(os.path.abspath(__file__))
DIRS = [
    os.path.join(ROOT, ".codebuddy", "memory"),
    os.path.join(ROOT, ".workbuddy", "memory"),
]

# 自己刚写入的 (目录, 文件名) -> 时间戳，用于防回环
_self_writes = {}
_lock = threading.Lock()


def log(msg):
    ts = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line, flush=True)


def file_sig(path):
    """(size, mtime_ms)，用于快速判断文件是否变动。"""
    if not os.path.isfile(path):
        return None
    st = os.stat(path)
    return (st.st_size, int(st.st_mtime * 1000))


def copy_bytes(src, dst):
    """字节级复制并保留 mtime，避免下次轮询误判为变更。"""
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    with open(src, "rb") as f:
        data = f.read()
    with open(dst, "wb") as f:
        f.write(data)
    st = os.stat(src)
    os.utime(dst, (st.st_atime, st.st_mtime))


def scan(d):
    sigs = {}
    if os.path.isdir(d):
        for name in os.listdir(d):
            p = os.path.join(d, name)
            if os.path.isfile(p):
                sigs[name] = file_sig(p)
    return sigs


def do_sync(dry_run, interval):
    prev = {d: scan(d) for d in DIRS}
    if dry_run:
        log("DRY-RUN 模式：不会真正写入")
    log("开始监视：" + " | ".join(os.path.basename(os.path.dirname(d)) for d in DIRS))
    try:
        while True:
            time.sleep(interval)
            cur = {d: scan(d) for d in DIRS}

            for i, d in enumerate(DIRS):
                other = DIRS[1 - i]
                dname = os.path.basename(os.path.dirname(d))
                oname = os.path.basename(os.path.dirname(other))

                for name, sig in cur[d].items():
                    oprev = prev[d].get(name)

                    # 防回环：本侧文件若是我们刚写入的，跳过
                    with _lock:
                        sw = _self_writes.get((d, name))
                    if sw and (time.time() - sw) < (interval + 2):
                        continue

                    if sig == oprev:
                        continue  # 本侧无变化

                    dst = os.path.join(other, name)
                    # 冲突：两侧同一文件都变了
                    oth_sig = cur[other].get(name)
                    if (
                        oth_sig is not None
                        and name in prev[other]
                        and oth_sig != prev[other].get(name)
                    ):
                        # 取 mtime 较新者覆盖较旧者
                        src_sig = sig
                        win_is_other = (oth_sig[1] > src_sig[1])
                        src_dir, dst_dir = (
                            (other, d) if win_is_other else (d, other)
                        )
                        src_name = name
                        action = "冲突-取较新"
                    else:
                        src_dir, dst_dir = d, other
                        src_name = name
                        action = "同步"

                    if dry_run:
                        log(
                            f"[DRY] {action}: {src_name}  {os.path.basename(os.path.dirname(src_dir))} -> "
                            f"{os.path.basename(os.path.dirname(dst_dir))}"
                        )
                        continue

                    copy_bytes(os.path.join(src_dir, src_name), os.path.join(dst_dir, src_name))
                    with _lock:
                        _self_writes[(dst_dir, src_name)] = time.time()
                    log(
                        f"[{action}] {src_name}: {os.path.basename(os.path.dirname(src_dir))} -> "
                        f"{os.path.basename(os.path.dirname(dst_dir))}"
                    )

                # 删除检测：不传播，只提示
                for name in prev[d]:
                    if name not in cur[d]:
                        log(f"[删除提示@{dname}] {name} 已删除（不传播到 {oname}）")

            prev = cur
    except KeyboardInterrupt:
        log("已停止监视。")


def reconcile_once(dry_run):
    """单次调解：对每个文件名，按磁盘当前状态做 last-writer-wins，
    缺哪侧补哪侧，两侧都改了取 mtime 较新者。重新读取磁盘，不做陈旧快照。"""
    names = set()
    for d in DIRS:
        if os.path.isdir(d):
            names |= set(os.listdir(d))
    any_change = False
    for name in sorted(names):
        sigs, exists = {}, {}
        for d in DIRS:
            p = os.path.join(d, name)
            exists[d] = os.path.isfile(p)
            sigs[d] = file_sig(p) if exists[d] else None
        present = [d for d in DIRS if exists[d]]
        if len(present) == 0:
            continue
        if len(present) == 1:
            src = present[0]
            dst = DIRS[1] if src == DIRS[0] else DIRS[0]
            if dry_run:
                to_log = f"[DRY] 补齐: {name} {os.path.basename(os.path.dirname(src))} -> " \
                         f"{os.path.basename(os.path.dirname(dst))}"
            else:
                copy_bytes(os.path.join(src, name), os.path.join(dst, name))
                to_log = f"[补齐] {name}: {os.path.basename(os.path.dirname(src))} -> " \
                         f"{os.path.basename(os.path.dirname(dst))}"
                any_change = True
            log(to_log)
            continue
        # 两侧都存在
        if sigs[DIRS[0]] == sigs[DIRS[1]]:
            continue  # 完全一致
        winner = DIRS[0] if sigs[DIRS[0]][1] >= sigs[DIRS[1]][1] else DIRS[1]
        loser = DIRS[1] if winner == DIRS[0] else DIRS[0]
        if dry_run:
            to_log = f"[DRY] 冲突-取较新: {name} {os.path.basename(os.path.dirname(winner))} -> " \
                     f"{os.path.basename(os.path.dirname(loser))}"
        else:
            copy_bytes(os.path.join(winner, name), os.path.join(loser, name))
            to_log = f"[冲突-取较新] {name}: {os.path.basename(os.path.dirname(winner))} -> " \
                     f"{os.path.basename(os.path.dirname(loser))}"
            any_change = True
        log(to_log)
    if not any_change and not dry_run:
        log("两侧已一致，无需同步。")


def main():
    ap = argparse.ArgumentParser(description="记忆目录双向同步监视器")
    ap.add_argument("--once", action="store_true", help="只同步一次后退出")
    ap.add_argument("--dry-run", action="store_true", help="只预览，不真正写入")
    ap.add_argument("--interval", type=float, default=1.5, help="轮询间隔秒数（默认 1.5）")
    args = ap.parse_args()

    if args.once:
        reconcile_once(args.dry_run)
        return

    do_sync(args.dry_run, args.interval)


if __name__ == "__main__":
    main()
