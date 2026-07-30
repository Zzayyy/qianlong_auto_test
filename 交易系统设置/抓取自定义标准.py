# -*- coding: utf-8 -*-
"""交易系统设置 - 抓取当前客户端自定义标准。

把“当前客户端版本”的委托设置等界面值读取出来，写回
交易系统设置/标准/<client_id>/<panel>.json，作为新的自定义比对标准。
覆盖前自动备份原 JSON 为 .json.bak。

特点：只读采集，不点击“应用/恢复默认”，不改任何设置。

用法：
    python 抓取自定义标准.py                 # 用 GUI_CLIENT_ID / 默认客户端
    python 抓取自定义标准.py --client guotai_haitong
    python 抓取自定义标准.py --panel 委托设置 期权设置
    GUI 参数配置勾选后通过环境变量 GUI_CAPTURE_PANELS 传入（逗号分隔）。
"""

import argparse
import importlib.util
import os
import sys
import time
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import open_settings_dialog, switch_settings_panel
from core.settings_standard import save_standard, save_super_price_template

try:
    from core.clients import get_client, get_default_client_id
except Exception:  # noqa: BLE001
    get_client = None
    get_default_client_id = None

# 面板 -> 脚本相对路径。仅登记已实现 collect_current_settings 的面板；
# 其余面板待各自接入后可在此追加。
PANEL_MODULES = {
    "委托设置": r"1_委托设置.py",
    "期权设置": r"2_期权设置.py",
    "自动拆单设置": r"3_自动拆单设置.py",
    "自动追单设置": r"4_自动追单设置.py",
    "快捷设置": r"5_快捷设置.py",
    "价格提醒设置": r"6_价格提醒设置.py",
    "一键炒单设置": r"6_一键炒单设置.py",
}

SETTINGS_BUTTON_AUTO_ID = "1008"
SETTINGS_MENU_ITEM_AUTO_ID = "20025"
SETTINGS_DIALOG_TITLE = "交易系统设置"
COUNTDOWN_SEC = 3


def _import_panel(module_rel_path: str):
    path = Path(__file__).resolve().parent / module_rel_path
    spec = importlib.util.spec_from_file_location("capture_panel_mod", str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def _resolve_client_id(arg_client: str) -> str:
    env = os.environ.get("GUI_CLIENT_ID", "") or ""
    if arg_client:
        return arg_client
    if env:
        return env
    if get_default_client_id is not None:
        try:
            return get_default_client_id()
        except Exception:  # noqa: BLE001
            return ""
    return ""


def main():
    parser = argparse.ArgumentParser(description="抓取当前客户端自定义标准")
    parser.add_argument("--client", default="", help="客户端 id，如 guotai_haitong")
    parser.add_argument("--panel", nargs="*", default=None, help="只抓取指定面板")
    parser.add_argument("--super-price", action="store_true",
                        help="额外采集超价参数弹窗全部行，写 标准/<client_id>/超价设置.json")
    args = parser.parse_args()

    client_id = _resolve_client_id(args.client)
    panels = args.panel
    if not panels:
        # GUI 参数配置勾选后通过环境变量传入（逗号分隔的面板名）
        env_panels = os.environ.get("GUI_CAPTURE_PANELS", "").strip()
        if env_panels:
            panels = [p.strip() for p in env_panels.split(",") if p.strip()]
    if not panels:
        panels = list(PANEL_MODULES.keys())

    print("=" * 60)
    print("交易系统设置 - 抓取自定义标准")
    print("=" * 60)
    print(f"客户端: {client_id or '(默认兜底)'}")

    # 解析主窗口标题关键字（按客户端取 window_key）
    window_key = None
    if client_id and get_client is not None:
        try:
            client = get_client(client_id)
            window_key = (client or {}).get("window_key")
        except Exception:  # noqa: BLE001
            window_key = None
    if not window_key:
        window_key = "钱龙模拟"

    dlg = None
    hwnd = None
    try:
        countdown(COUNTDOWN_SEC)
        hwnd = find_window(window_key)
        print(f"[OK] 已找到主窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)

        dlg = open_settings_dialog(
            win, SETTINGS_BUTTON_AUTO_ID, SETTINGS_MENU_ITEM_AUTO_ID, SETTINGS_DIALOG_TITLE
        )
        dlg.wait("ready", timeout=10)

        for panel in panels:
            if panel not in PANEL_MODULES:
                print(f"[跳过] 未登记采集逻辑的面板: {panel}")
                continue
            mod = _import_panel(PANEL_MODULES[panel])
            collector = getattr(mod, "collect_current_settings", None)
            if collector is None:
                print(f"[跳过] {panel} 尚未实现 collect_current_settings")
                continue
            if not switch_settings_panel(dlg, panel):
                print(f"[错误] 无法切换到 {panel} 面板")
                continue
            time.sleep(0.5)
            data = collector(dlg)
            path = save_standard(panel, client_id, data)
            print(f"[OK] 已保存 {panel} 标准 -> {path}（共 {len(data)} 项）")

        # 超价参数弹窗（多页滚动采集，按客户端生成 超价设置.json）
        if args.super_price:
            if not client_id:
                print("[跳过] 未指定 --client，超价参数按客户端保存需要 client_id")
            else:
                try:
                    mod = _import_panel(PANEL_MODULES["期权设置"])
                    collect = getattr(mod, "collect_super_price_rows", None)
                    if collect is None:
                        print("[跳过] 2_期权设置.py 未实现 collect_super_price_rows")
                    else:
                        sp_rows = collect(dlg)
                        if sp_rows:
                            sp_path = save_super_price_template(client_id, sp_rows)
                            print(f"[OK] 已保存 超价设置 标准 -> {sp_path}（共 {len(sp_rows)} 行）")
                        else:
                            print("[WARN] 超价参数采集为空，未写入文件")
                except Exception as e:
                    print(f"[错误] 超价参数采集失败: {e}")

        print("\n=== 抓取完成 ===")
    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
    finally:
        if dlg is not None:
            close_settings_dialog(dlg, keep_open=False, main_hwnd=hwnd)


if __name__ == "__main__":
    main()
