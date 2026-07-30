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


# ---- 超价参数模板（行表，与面板比对标准分开存储） ----
# 超价参数弹窗的“品种/买超价步长/卖超价步长”是一张行表，不是扁平字典，
# 因此不走 load_standard 的 standards 字段，单独按客户端解析：
#   标准/<client_id>/超价设置.json -> 标准/default/超价设置.json -> 内置兜底。
SUPER_PRICE_TEMPLATE_FILE = "超价设置.json"

_BUILTIN_SUPER_PRICE_ROWS = [
    {"品种": "510050期权", "买": 1, "卖": -1},
    {"品种": "沪510300期权", "买": 1, "卖": -1},
    {"品种": "510500期权", "买": 1, "卖": -1},
    {"品种": "588000期权", "买": 1, "卖": -1},
    {"品种": "588080期权", "买": 1, "卖": -1},
    {"品种": "159901期权", "买": 1, "卖": -1},
    {"品种": "159915期权", "买": 1, "卖": -1},
    {"品种": "深159919期权", "买": 1, "卖": -1},
    {"品种": "159922期权", "买": 1, "卖": -1},
]


def _load_full_json(path: Path) -> Optional[Dict[str, Any]]:
    """读取整个 JSON 文件（不仅限于 standards 字段）。缺失/损坏返回 None。"""
    if not path.is_file():
        return None
    try:
        with path.open("r", encoding="utf-8-sig") as fh:
            return json.load(fh)
    except Exception as exc:  # noqa: BLE001
        print(f"[WARN] 读取 JSON 失败 {path}: {exc}")
        return None


def load_super_price_template(
    client_id: str,
    default: Optional[list] = None,
) -> list:
    """返回超价参数模板行列表 [{"品种":.., "买":.., "卖":..}, ...]。

    解析顺序：
        1. 标准/<client_id>/超价设置.json 的 rows
        2. 标准/default/超价设置.json 的 rows
        3. 调用方传入的 default（脚本内兜底）
        4. 模块内置 _BUILTIN_SUPER_PRICE_ROWS（离线不崩）
    """
    if client_id:
        data = _load_full_json(STANDARD_ROOT / client_id / SUPER_PRICE_TEMPLATE_FILE)
        rows = data.get("rows") if isinstance(data, dict) else None
        if rows:
            return rows
    data = _load_full_json(STANDARD_ROOT / "default" / SUPER_PRICE_TEMPLATE_FILE)
    rows = data.get("rows") if isinstance(data, dict) else None
    if rows:
        return rows
    if default:
        return list(default)
    return list(_BUILTIN_SUPER_PRICE_ROWS)


def save_super_price_template(
    client_id: str,
    rows: list,
    *,
    description: str = "",
    backup: bool = True,
) -> Path:
    """把超价参数行表写回 标准/<client_id>/超价设置.json。

    与 load_super_price_template 对称：覆盖前若已存在旧文件，先复制为
    同目录 .json.bak 备份。返回最终写出的 JSON 路径。
    """
    path = resolve_standard_path("超价设置", client_id)
    path.parent.mkdir(parents=True, exist_ok=True)
    if backup and path.is_file():
        shutil.copy2(path, path.with_suffix(".json.bak"))
    payload = {
        "schema_version": SCHEMA_VERSION,
        "client": client_id,
        "panel": "超价设置",
        "description": description or "自定义超价参数比对模板；可直接编辑，或用抓取脚本覆盖。",
        "rows": rows,
    }
    tmp = path.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
        fh.write("\n")
    os.replace(tmp, path)
    return path
