# -*- coding: utf-8 -*-
"""钱龙期权交易 - 组合策略持仓查询 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "组合策略持仓查询",
    "panel_path": r"\组合申报\组合策略持仓查询",
    "default_txt": r"E:\Code\3\output\组合策略持仓.txt",
    "default_xls": r"E:\Code\3\output\组合策略持仓.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
