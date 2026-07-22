# -*- coding: utf-8 -*-
"""兼容旧文件名的组合申报 Win32 入口；正式实现位于 core.combination_order。"""

import os
import sys


_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.combination_order import main  # noqa: E402


if __name__ == "__main__":
    main()
