# -*- coding: utf-8 -*-
"""国泰海通证券期权宝 - 账号查询 - 一键数据导出（国泰海通专属，钱龙不支持）"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "账号查询",
    "script_id": "gt_zhcx",
    "panel_path": r"\查询\账号查询",
    "default_txt": r"E:\Code\3\output\账号查询.txt",
    "default_xls": r"E:\Code\3\output\账号查询.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
