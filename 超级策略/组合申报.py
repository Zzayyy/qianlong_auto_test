# -*- coding: utf-8 -*-
"""超级策略 - 组合申报：OCR 识别并点击组合申报页签。"""
import os
import sys

_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from core.super_strategy import main_combination_declare


if __name__ == "__main__":
    main_combination_declare()
