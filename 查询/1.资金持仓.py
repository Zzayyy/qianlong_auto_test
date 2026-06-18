# -*- coding: utf-8 -*-
"""钱龙期权交易 - 资金持仓查询 - 一键数据导出"""

from core.runner import run_save_as

CONFIG = {
    "name": "资金持仓",
    "panel_path": r"\查询\资金持仓",
    "export_type": "xls_only",
    "default_xls": r"E:\Code\3\output\资金持仓.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_save_as(CONFIG)
