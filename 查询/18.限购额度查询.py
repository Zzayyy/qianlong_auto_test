# -*- coding: utf-8 -*-
"""钱龙期权交易 - 限购额度查询 - 一键数据导出"""

from core.runner import run_export_dialog

CONFIG = {
    "name": "限购额度查询",
    "panel_path": r"\查询\限购额度查询",
    "default_txt": r"E:\Code\3\output\限购额度查询.txt",
    "default_xls": r"E:\Code\3\output\限购额度查询.xls",
    "countdown_sec": 3,
    "settle_delay": 1.0,
}

if __name__ == "__main__":
    run_export_dialog(CONFIG)
