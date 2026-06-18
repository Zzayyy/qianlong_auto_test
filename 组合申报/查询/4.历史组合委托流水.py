# -*- coding: utf-8 -*-
"""钱龙期权交易 - 历史组合委托流水 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "历史组合委托流水",
    "panel_path": r"\组合申报\历史组合委托流水",
    "default_txt": r"E:\Code\3\output\历史组合委托流水.txt",
    "default_xls": r"E:\Code\3\output\历史组合委托流水.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
