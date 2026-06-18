# -*- coding: utf-8 -*-
"""钱龙期权交易 - 当日委托查询 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "当日委托",
    "panel_path": r"\查询\当日委托",
    "default_txt": r"E:\Code\3\output\当日委托.txt",
    "default_xls": r"E:\Code\3\output\当日委托.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
