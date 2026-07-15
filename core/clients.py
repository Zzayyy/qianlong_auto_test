# -*- coding: utf-8 -*-
"""客户端档案：多客户端（不同菜单/窗口标题）的集中配置与解析。

clients.json 结构示例：
{
  "default_client": "qianlong",
  "clients": [
    {
      "id": "qianlong",
      "name": "钱龙模拟期权宝",
      "window_key": "钱龙模拟期权宝",
      "menu_map": {},            // 逻辑键 -> 该客户端真实菜单路径
      "unsupported": []          // 该客户端不支持的 script_id 列表
    },
    {
      "id": "guotai_haitong",
      "name": "国泰海通期权宝",
      "window_key": "国泰海通期权宝",
      "menu_map": {
        "\\查询\\期权合约": "期权\\期权合约查询"
      },
      "unsupported": ["通知查询"]
    }
  ]
}

解析规则（resolve_panel_path）：
  1. script_id 命中 unsupported  -> 返回 None（当前客户端不支持）
  2. script_id 命中 menu_map     -> 返回映射的真实路径（覆盖）
  3. panel_path 命中 menu_map     -> 返回映射的真实路径（兼容旧脚本）
  4. 否则返回脚本内嵌的 panel_path（标准/默认路径，天然共享）

新增客户端/新菜单时，仅需编辑 clients.json，无需改动脚本与 core。
"""

import os
import json

_HERE = os.path.dirname(os.path.abspath(__file__))
_PROJECT_ROOT = os.path.dirname(_HERE)


def _clients_file():
    """返回 clients.json 路径（自动探测若干可能位置）"""
    candidates = [
        os.environ.get("GUI_CLIENTS_FILE"),                       # 允许通过环境变量覆盖
        os.path.join(_PROJECT_ROOT, "clients.json"),               # 开发环境：项目根
        os.path.join(_HERE, "clients.json"),                       # core/ 内备用
        os.path.join(_PROJECT_ROOT, "GUI自动化工具2", "clients.json"),
    ]
    for c in candidates:
        if c and os.path.exists(c):
            return c
    return candidates[1]  # 默认返回项目根位置（即便暂不存在，便于报错提示）


def load_clients():
    """加载 clients.json，返回 {"default_client":..., "clients":[...]}"""
    path = _clients_file()
    if not os.path.exists(path):
        return {"default_client": None, "clients": []}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, list):
            return {"default_client": None, "clients": data}
        return data
    except Exception:
        return {"default_client": None, "clients": []}


def get_clients():
    return load_clients().get("clients", [])


def get_client_ids():
    return [c.get("id") for c in get_clients() if c.get("id")]


def get_client(id_):
    """按 id 取客户端档案；id 为空/未找到返回 None"""
    if not id_:
        return None
    for c in get_clients():
        if c.get("id") == id_:
            return c
    return None


def get_client_name(id_):
    c = get_client(id_)
    return c.get("name") if c else (id_ or "")


def get_default_client_id():
    data = load_clients()
    default = data.get("default_client")
    if default and get_client(default):
        return default
    ids = get_client_ids()
    return ids[0] if ids else None


def resolve_panel_path(config, client_id=None):
    """根据脚本配置与客户端，返回实际菜单路径；不支持时返回 None。

    config: 脚本的 CONFIG 字典，需含 panel_path，可选含 script_id
    client_id: 当前客户端 id（来自 GUI_CLIENT_ID 环境变量）
    """
    client = get_client(client_id) if client_id else None
    std_path = config.get("panel_path", "")
    sid = config.get("script_id")

    unsupported = (client or {}).get("unsupported", [])
    if sid and sid in unsupported:
        return None
    if std_path and std_path in unsupported:
        return None

    if client:
        menu_map = client.get("menu_map", {})
        if sid and sid in menu_map:
            return menu_map[sid]
        if std_path and std_path in menu_map:
            return menu_map[std_path]

    return std_path
