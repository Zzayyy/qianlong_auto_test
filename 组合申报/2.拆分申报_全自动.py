# -*- coding: utf-8 -*-
"""正式拆分申报入口：固定以拆分模式运行共享 Win32 兼容驱动。"""

import os
import sys


os.environ["GUI_COMBINATION_ACTION"] = "split"
_PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _PROJECT_ROOT not in sys.path:
    sys.path.insert(0, _PROJECT_ROOT)

from core.combination_order import main  # noqa: E402


if __name__ == "__main__":
    main()
