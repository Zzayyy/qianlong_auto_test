# -*- coding: utf-8 -*-
"""通知查询 - 一键数据导出（复用 core 引擎，与查询类脚本逻辑一致）

说明：
  - 该菜单为部分客户端独有（如国泰海通期权宝），钱龙模拟期权宝无此菜单。
  - 通过 script_id="通知查询" 在 clients.json 的 unsupported 中控制显隐，
    无需为不同客户端复制脚本。
"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "通知查询",
    "script_id": "通知查询",
    "panel_path": r"\通知查询\合约信息变更数量",
    "export_format": "xls",
    "default_txt": r"E:\Code\3\output\通知查询.txt",
    "default_xls": r"E:\Code\3\output\通知查询.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
