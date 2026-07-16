# -*- coding: utf-8 -*-
"""通用查询驱动：所有「查询类」薄壳脚本统一收敛到本文件（查询 / 组合申报查询 / 通知查询 等）。

工作方式：
  - 由环境变量 GUI_QUERY_KEY 指明要执行的具体查询（值为 queries.json 的 key，
    即默认/钱龙菜单路径，如 "\\查询\\资金持仓" 或 "\\组合申报\\组合策略持仓查询"）。
  - 从同目录 queries.json 读取该查询的导出参数（settle_delay / 输出路径 / 是否另存为等）。
  - 客户端差异由 core.runner + clients.json 在运行时按 GUI_CLIENT_ID 自动解析菜单路径，
    本驱动完全不关心客户端。

历史：原本各分类下大量脚本仅是「一份 CONFIG + 一个 run_export_dialog/run_save_as」的薄壳，
现统一为本文件 + queries.json 数据表。
"""

import os
import sys
import json

HERE = os.path.dirname(os.path.abspath(__file__))
# 保证能 import core（项目根即本文件所在目录；与子进程 PYTHONPATH=PROJECT_ROOT 等效的安全网）
if HERE not in sys.path:
    sys.path.insert(0, HERE)

from core.runner import run_export_dialog, run_save_as


def main():
    key = os.environ.get("GUI_QUERY_KEY")
    if not key:
        print("[错误] 未设置 GUI_QUERY_KEY，无法确定要执行的查询")
        sys.exit(1)

    table_path = os.path.join(HERE, "queries.json")
    try:
        with open(table_path, "r", encoding="utf-8") as f:
            table = json.load(f)
    except Exception as e:
        print(f"[错误] 读取 queries.json 失败: {e}")
        sys.exit(1)

    if key not in table:
        print(f"[错误] queries.json 中未找到查询配置: {key}")
        sys.exit(1)

    entry = dict(table[key])
    entry["panel_path"] = key
    entry["name"] = key.split("\\")[-1]

    # 输出路径兜底：GUI 运行时经 GUI_TXT_PATH/GUI_XLS_PATH 覆盖本默认值；
    # 仅在脱离 GUI 单独运行时，按「桌面\<分类>\<菜单名>」自动生成默认路径
    # （桌面路径与具体用户相关，不能写死在 queries.json 里）。
    parts = key.split("\\")
    cat = parts[1] if len(parts) > 1 else ""
    base = os.path.join(os.path.expanduser("~"), "Desktop", cat) if cat else os.path.join(os.path.expanduser("~"), "Desktop")
    entry.setdefault("default_txt", os.path.join(base, f"{entry['name']}.txt"))
    entry.setdefault("default_xls", os.path.join(base, f"{entry['name']}.xls"))

    if entry.get("export_type") == "xls_only":
        run_save_as(entry)
    else:
        run_export_dialog(entry)


if __name__ == "__main__":
    main()
