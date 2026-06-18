# -*- coding: utf-8 -*-
"""钱龙期权交易 - 历史成交查询 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "历史成交",
    "panel_path": r"\查询\历史成交",
    "default_txt": r"E:\Code\3\output\历史成交.txt",
    "default_xls": r"E:\Code\3\output\历史成交.xls",
    "countdown_sec": 3,
    "settle_delay": 10.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
