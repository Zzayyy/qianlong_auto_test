# -*- coding: utf-8 -*-
"""任务模型：封装一个待执行脚本及其运行参数，并构造子进程环境变量"""

import os


class Task:
    """一次脚本执行任务：脚本信息 + 分类 + 运行时参数"""

    def __init__(self, script, category, params=None):
        self.script = script            # {"name": str, "path": str}
        self.category = category
        self.params = params or {}

    @property
    def name(self):
        return self.script["name"]

    @property
    def path(self):
        return self.script["path"]

    def build_env(self, project_root):
        """根据参数构造子脚本所需的环境变量"""
        env = os.environ.copy()
        env["PYTHONPATH"] = project_root  # 让子脚本能找到 core 模块

        p = self.params
        env["GUI_EXPORT_FORMAT"] = p.get("export_format", "txt")
        env["GUI_AUTO_OPEN"] = str(p.get("auto_open", False))
        env["GUI_TXT_PATH"] = p.get("txt_path", "")
        env["GUI_XLS_PATH"] = p.get("xls_path", "")
        env["GUI_ORDER_QTY"] = str(p.get("order_qty", 1))
        env["GUI_COUNTDOWN"] = str(p.get("countdown_sec", 3))
        env["GUI_XLSX_FILE"] = p.get("xlsx_file", "")
        env["GUI_CATEGORY"] = self.category

        # 期权下单_一键导出 参数
        export_targets = p.get("export_targets", [])
        env["GUI_EXPORT_TARGETS"] = ",".join(export_targets)
        env["GUI_EXPORT_DIR"] = p.get("export_output_dir", "")

        # 交易系统设置 输出路径
        if self.category == "交易系统设置":
            env["GUI_OUTPUT_DIR"] = p.get("settings_output_dir", "")

        return env
