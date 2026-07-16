# -*- coding: utf-8 -*-
"""国泰海通证券期权宝 - 对账单资金流水 - 一键数据导出（国泰海通专属，钱龙不支持）"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "对账单资金流水",
    "script_id": "gt_dzdjls",
    "panel_path": r"\查询\对账单资金流水",
    "default_txt": r"E:\Code\3\output\对账单资金流水.txt",
    "default_xls": r"E:\Code\3\output\对账单资金流水.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
