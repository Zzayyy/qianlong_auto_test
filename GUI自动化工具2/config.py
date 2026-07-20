# -*- coding: utf-8 -*-
"""
应用环境与用户配置 / 脚本清单
-------------------------------
集中管理：
  - PyInstaller 打包环境检测 (IS_FROZEN / BASE_DIR / PROJECT_ROOT)
  - 配置文件路径 (CONFIG_FILE)
  - 用户配置读写 (load/save_user_config)
  - 输出目录工具 (get/set_output_dir)
  - 脚本清单 (SCRIPTS_CONFIG / CATEGORIES)
"""

import os
import sys
import json
import re

# ====================== PyInstaller 打包环境 ======================
# 判断是否从 PyInstaller 打包的 exe 运行
IS_FROZEN = getattr(sys, 'frozen', False)

if IS_FROZEN:
    # 打包后：脚本在 _internal 目录下
    BASE_DIR = os.path.join(os.path.dirname(sys.executable), '_internal')
else:
    # 开发环境：项目根目录（GUI自动化工具2 的上一级）
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 项目根目录（打包后指向 _internal，开发环境指向项目根目录）
PROJECT_ROOT = BASE_DIR

# ====================== 客户端档案（多客户端支持） ======================
# 让本模块能引用 core.clients（项目根目录的 core 包），集中维护多客户端逻辑
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)

from core.clients import (  # noqa: E402
    load_clients,
    get_clients,
    get_client,
    get_client_ids,
    get_client_name,
    get_default_client_id,
)

# ====================== 用户配置 ======================
# 配置文件路径：打包后放在exe同级目录，开发环境放在脚本同级目录
if IS_FROZEN:
    CONFIG_DIR = os.path.dirname(sys.executable)
else:
    CONFIG_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(CONFIG_DIR, "config.json")
# 默认输出目录：普通用户机器上放到桌面（开发者机器路径 E:\Code\3\output 不适用）
DEFAULT_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop")
CATEGORIES = ["查询", "通知查询", "结算单", "下单", "组合申报", "交易系统设置"]


def load_user_config() -> dict:
    """加载用户配置"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                # 统一路径斜杠为 Windows 反斜杠
                if "output_dirs" in cfg:
                    for cat in cfg["output_dirs"]:
                        cfg["output_dirs"][cat] = cfg["output_dirs"][cat].replace("/", "\\")
                return cfg
        except Exception:
            pass
    return {
        "output_dirs": {cat: DEFAULT_OUTPUT_DIR for cat in CATEGORIES},
        "auto_open": False,
        "export_format": "txt",
        "log_level": "详细",
        "client": get_default_client_id(),
    }


def save_user_config(config: dict):
    """保存用户配置"""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"保存配置失败: {e}")


def get_output_dir(config: dict, category: str) -> str:
    """获取指定界面的输出目录。

    若配置为空字符串/缺失，则回退到 DEFAULT_OUTPUT_DIR（桌面），
    避免 os.path.join("", "x.txt") 得到裸文件名、导致文件落进程序安装目录。
    """
    dirs = config.get("output_dirs", {})
    d = dirs.get(category, DEFAULT_OUTPUT_DIR)
    if not d or not str(d).strip():
        d = DEFAULT_OUTPUT_DIR
    return d


def set_output_dir(config: dict, category: str, path: str):
    """设置指定界面的输出目录"""
    if "output_dirs" not in config:
        config["output_dirs"] = {}
    config["output_dirs"][category] = path.replace("/", "\\")


def get_script_filename(script_name: str) -> str:
    """从脚本名提取纯文件名（去掉编号和空格）"""
    # "3. 策略委托" -> "策略委托"
    name = re.sub(r"^\d+\.\s*", "", script_name).strip()
    return name


# ====================== 脚本清单 ======================
SCRIPTS_CONFIG = {
    # 查询类脚本已统一为单个驱动 查询/run_query.py + 查询/queries.json 数据表。
    # 每个条目通过 query_key（=默认菜单路径）告诉驱动具体执行哪一项；
    # 客户端差异由 clients.json 的 menu_map/unsupported + core.runner 在运行时解析。
    "查询": [
        {"name": "1. 资金持仓", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\资金持仓"},
        {"name": "2. 策略持仓", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\策略持仓"},
        {"name": "3. 策略委托", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\策略委托"},
        {"name": "4. 历史策略持仓", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史策略持仓"},
        {"name": "5. 当日委托", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\当日委托"},
        {"name": "6. 当日成交", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\当日成交"},
        {"name": "7. 历史委托", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史委托"},
        {"name": "8. 历史成交", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史成交"},
        {"name": "9. 期权合约", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\期权合约"},
        {"name": "10. 资金查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\资金查询"},
        {"name": "11. 当日行权被指派查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\当日行权被指派查询"},
        {"name": "12. 行权结果查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\行权结果查询"},
        {"name": "13. 历史行权指派查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史行权指派查询"},
        {"name": "14. 历史对账单查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史对账单查询"},
        {"name": "15. 历史交割单查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史交割单查询"},
        {"name": "16. 历史行权交割查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史行权交割查询"},
        {"name": "17. 客户限仓信息查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\客户限仓信息查询"},
        {"name": "18. 限购额度查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\限购额度查询"},
        {"name": "19. 行权负债信息查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\行权负债信息查询"},
        {"name": "20. 历史行权负债信息", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\历史行权负债信息"},
        # 以下为国泰海通专属菜单（钱龙无此菜单，由 clients.json 的 unsupported 过滤）
        {"name": "21. 可用备兑股份", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\可用备兑股份"},
        {"name": "22. 备兑股份不足", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\备兑股份不足"},
        {"name": "23. 当日行权成功明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\当日行权成功明细"},
        {"name": "24. 对账单资金流水", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\对账单资金流水"},
        {"name": "25. 账号查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\账号查询"},
        {"name": "26. 对账单资金资产", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\查询\对账单资金资产"},
    ],
    "下单": [
        {"name": "1.期权下单_自动化下单", "path": rf"{PROJECT_ROOT}\下单\自动化下单\4.期权下单(新)_自动化下单_Excel驱动版.py"},
        #{"name": "2.期权持仓_平仓/反手自动化", "path": rf"{PROJECT_ROOT}\下单\表格\9.Excel驱动_OCR定位_平仓操作.py"},
        {"name": "2.三键下单_自动化下单", "path": rf"{PROJECT_ROOT}\下单\自动化下单\4.三键下单_自动化下单_Excel驱动版.py"},
        {"name": "3.四键下单_自动化下单", "path": rf"{PROJECT_ROOT}\下单\自动化下单\4.四键下单_自动化下单_Excel驱动版.py"},
        {"name": "4.期权持仓_平仓/反手自动化_RapidOCR", "path": rf"{PROJECT_ROOT}\下单\表格\10.Excel驱动_OCR_RapidOCR.py"},
        {"name": "5.期权下单_一键导出", "path": rf"{PROJECT_ROOT}\下单\自动化导出\期权下单(新)_自动导出.py"},
        {"name": "6.全选撤单", "path": rf"{PROJECT_ROOT}\撤单\撤单_全选撤单_自动化.py"},
    ],
    # 以下为国泰海通专属菜单
    "通知查询": [
        {"name": "1.合约信息变更数量", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\通知查询\合约信息变更数量"},
        {"name": "2.合约信息变更明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\通知查询\合约信息变更明细"},
        {"name": "3.当日风险通知", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\通知查询\当日风险通知"},
        {"name": "4.历史风险通知", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\通知查询\历史风险通知"},
    ],
    # 以下为国泰海通专属菜单
    "结算单": [
        {"name": "1.结算单", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\结算单"},
        {"name": "2.历史出入金流水明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史出入金流水明细"},
        {"name": "3.历史出入金流水汇总", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史出入金流水汇总"},
        {"name": "4.历史合约持仓明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史合约持仓明细"},
        {"name": "5.历史交收明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史交收明细"},
        {"name": "6.历史交收金额汇总", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史交收金额汇总"},
        {"name": "7.历史行权指派明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史行权指派明细"},
        {"name": "8.历史资金资产", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史资金资产"},
        {"name": "9.历史行权交收流水明细", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史行权交收流水明细"},
        {"name": "10.历史风险信息", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\结算单\历史风险信息"},
    ],
    "组合申报": [
        {"name": "1.组合申报_全自动", "path": rf"{PROJECT_ROOT}\组合申报\2.组合申报_全自动.py"},
        {"name": "2.拆分申报_全自动", "path": rf"{PROJECT_ROOT}\组合申报\2.拆分申报_全自动.py"},
        {"name": "3.组合策略持仓查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\组合申报\组合策略持仓查询"},
        {"name": "4.组合策略信息查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\组合申报\组合策略信息查询"},
        {"name": "5.组合委托流水查询", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\组合申报\组合委托流水查询"},
        {"name": "6.历史组合委托流水", "path": rf"{PROJECT_ROOT}\run_query.py", "query_key": r"\组合申报\历史组合委托流水"},
    ],
    "交易系统设置": [
        {"name": "1.委托设置", "path": rf"{PROJECT_ROOT}\交易系统设置\1_委托设置.py"},
        {"name": "2.期权设置", "path": rf"{PROJECT_ROOT}\交易系统设置\2_期权设置.py"},
        {"name": "3.自动拆单设置", "path": rf"{PROJECT_ROOT}\交易系统设置\3_自动拆单设置.py"},
        {"name": "4.自动追单设置", "path": rf"{PROJECT_ROOT}\交易系统设置\4_自动追单设置.py"},
        {"name": "5.快捷设置", "path": rf"{PROJECT_ROOT}\交易系统设置\5_快捷设置.py"},
        {"name": "6.一键炒单设置", "path": rf"{PROJECT_ROOT}\交易系统设置\6_一键炒单设置.py"},
        {"name": "7.价格提醒设置", "path": rf"{PROJECT_ROOT}\交易系统设置\6_价格提醒设置.py"},
    ],
}


# ====================== 按客户端解析脚本清单 ======================
def _script_menu_key(category: str, menu_name: str) -> str:
    """构造与 clients.json 中 menu_map / unsupported 对应的菜单路径键。

    例如 ("查询", "资金查询") -> "\\查询\\资金查询"
    """
    return f"\\{category}\\{menu_name}"


def get_scripts_config(client_id=None) -> dict:
    """根据客户端返回解析后的脚本清单（对应各自软件的菜单）。

    解析规则（与 core.clients.resolve_panel_path 对齐）：
      1. 该客户端 unsupported 中出现的脚本（去掉编号后的菜单名 / 菜单路径键）
         将被过滤掉，不在界面显示；
      2. 命中 menu_map 的脚本，其显示名替换为该客户端软件中的真实菜单名
         （保留原编号前缀，仅替换菜单文本）。
    默认（client_id 为空 / 找不到对应客户端）返回基准（钱龙）配置。
    """
    client = get_client(client_id) if client_id else None
    menu_map = (client or {}).get("menu_map", {}) or {}
    unsupported = set((client or {}).get("unsupported", []) or [])

    resolved = {}
    for category, scripts in SCRIPTS_CONFIG.items():
        out = []
        for s in scripts:
            menu_name = get_script_filename(s["name"])  # 去掉编号前缀后的菜单名
            # 1) 该客户端不支持 -> 过滤
            if menu_name in unsupported:
                continue
            key = _script_menu_key(category, menu_name)
            if key in unsupported:
                continue
            # 2) 命中菜单映射 -> 改名（保留编号前缀）
            name = s["name"]
            if key in menu_map:
                mapped = menu_map[key].split("\\")[-1]
                m = re.match(r"^(\s*\d+\.\s*)", name)
                prefix = m.group(1) if m else ""
                name = prefix + mapped
            out.append({**s, "name": name})
        resolved[category] = out
    return resolved
