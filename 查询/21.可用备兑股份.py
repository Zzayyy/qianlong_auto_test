# -*- coding: utf-8 -*-
"""国泰海通证券期权宝 - 可用备兑股份 - 一键数据导出（国泰海通专属，钱龙不支持）"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "可用备兑股份",
    "script_id": "gt_kybdgf",
    "panel_path": r"\查询\可用备兑股份",
    "default_txt": r"E:\Code\3\output\可用备兑股份.txt",
    "default_xls": r"E:\Code\3\output\可用备兑股份.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
