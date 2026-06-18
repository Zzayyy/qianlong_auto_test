# -*- coding: utf-8 -*-
"""钱龙期权交易 - 历史行权负债信息 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "历史行权负债信息",
    "panel_path": r"\查询\历史行权负债信息",
    "default_txt": r"E:\Code\3\output\历史行权负债信息.txt",
    "default_xls": r"E:\Code\3\output\历史行权负债信息.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
