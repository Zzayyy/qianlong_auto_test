# -*- coding: utf-8 -*-
"""交易系统设置 - 自定义标准(期望值)的加载与保存。

标准按 客户端/面板 拆分存放在 交易系统设置/标准/<client_id>/<panel>.json：
    - 比对脚本: load_standard(panel, client_id, default) 取得期望值字典。
        优先 client 专属文件 -> 其次 default 兜底目录 -> 最后内嵌 default。
    - 抓取脚本: save_standard(panel, client_id, standards, ...) 把当前界面值
        写回 JSON（覆盖前先备份旧文件 .json.bak）。

说明：
    - 控件 auto_id 映射(AUTO_ID)本次仍保留在各面板脚本内，不在本模块处理。
    - 每个 JSON 形如：
        {
          "schema_version": 1,
          "client": "qianlong",
          "panel": "委托设置",
          "description": "...",
          "standards": { ... 与脚本 STANDARD_VALUES 同构 ... }
        }
"""

import json
import os
import shutil
from pathlib import Path
from typing import Any, Dict, Optional

# 标准根目录：<项目根>/交易系统设置/标准
STANDARD_ROOT = Path(__file__).resolve().parent.parent / "交易系统设置" / "标准"
SCHEMA_VERSION = 1


def resolve_standard_path(panel: str, client_id: str) -> Path:
    """返回 标准/<client_id>/<panel>.json 的路径。"""
    return STANDARD_ROOT / client_id / f"{panel}.json"


def _load_standards(path: Path) -> Optional[Dict[str, Any]]:
    """读取 JSON 中的 standards 字典；文件缺失/损坏/缺字段返回 None。"""
    if not path.is_file():
        return None
    try:
        with path.open("r", encoding="utf-8-sig") as fh:
            data = json.load(fh)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] 读取标准文件失败 {path}: {exc}")
        return None
    std = data.get("standards")
    if isinstance(std, dict):
        return std
    print(f"[WARN] 标准文件缺少 standards 字段: {path}")
    return None


def load_standard(
    panel: str,
    client_id: str,
    default: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """返回期望标准字典（面板比对用的期望值）。

    解析顺序：
        1. 标准/<client_id>/<panel>.json（当前客户端自定义标准，优先）
        2. 标准/default/<panel>.json（通用兜底）
        3. 调用方内嵌 default（脚本内的硬编码兜底）
    """
    if client_id:
        std = _load_standards(resolve_standard_path(panel, client_id))
        if std is not None:
            return std
    std = _load_standards(STANDARD_ROOT / "default" / f"{panel}.json")
    if std is not None:
        return std
    return dict(default) if default else {}


def save_standard(
    panel: str,
    client_id: str,
    standards: Dict[str, Any],
    *,
    description: str = "",
    backup: bool = True,
) -> Path:
    """把 standards 写回 标准/<client_id>/<panel>.json。

    覆盖前若已存在旧文件，先复制为同目录 .json.bak 备份。
    返回最终写出的 JSON 路径。
    """
    path = resolve_standard_path(panel, client_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    if backup and path.is_file():
        shutil.copy2(path, path.with_suffix(".json.bak"))
    payload = {
        "schema_version": SCHEMA_VERSION,
        "client": client_id,
        "panel": panel,
        "description": description or "自定义标准；可直接编辑，或用抓取脚本覆盖。",
        "standards": standards,
    }
    tmp = path.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(tmp, path)
    return path
