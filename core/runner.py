# -*- coding: utf-8 -*-
"""统一执行入口：接收配置→解析参数→执行自动化流程"""

import os
import sys
import time

from core.window import find_window, activate_window, switch_panel, click_output_button, countdown
from core.export_dialog import handle_export_dialog
from core.save_as_dialog import handle_save_as_dialog


def parse_env_config(config: dict) -> dict:
    """从环境变量+默认配置中解析最终参数。

    Args:
        config: 脚本定义的默认配置字典
    Returns:
        合并后的参数字典
    """
    result = config.copy()

    # GUI传入的环境变量
    result["window_key"] = os.environ.get("GUI_WINDOW_KEY") or config.get("window_key", "钱龙模拟期权宝")
    result["export_format"] = os.environ.get("GUI_EXPORT_FORMAT") or config.get("export_format", "txt")
    result["countdown_sec"] = int(os.environ.get("GUI_COUNTDOWN") or config.get("countdown_sec", 3))
    result["settle_delay"] = float(os.environ.get("GUI_DELAY") or config.get("settle_delay", 1.0))

    # auto_open 处理
    env_auto = os.environ.get("GUI_AUTO_OPEN")
    if env_auto is not None:
        result["auto_open"] = env_auto.lower() == "true"
    else:
        result["auto_open"] = config.get("auto_open", True)

    # 输出路径
    if result.get("export_type") == "xls_only":
        # 仅xls模式(资金持仓)
        result["output_path"] = os.environ.get("GUI_XLS_PATH") or config.get("default_xls", "")
    else:
        # txt/xls双模式
        fmt = result["export_format"]
        if fmt == "txt":
            result["output_path"] = os.environ.get("GUI_TXT_PATH") or config.get("default_txt", "")
        else:
            result["output_path"] = os.environ.get("GUI_XLS_PATH") or config.get("default_xls", "")

    # 清理路径中的不可见字符，并统一斜杠为 Windows 反斜杠
    result["output_path"] = ''.join(c for c in result["output_path"] if c.isprintable()).strip()
    result["output_path"] = result["output_path"].replace("/", "\\")

    return result


def run_export_dialog(config: dict):
    """执行标准数据输出弹窗流程。

    Args:
        config: 配置字典，需包含:
            - panel_path: 树形面板路径,如 r"\查询\策略委托"
            - name: 显示名称,如 "策略委托"
            - default_txt: txt默认路径
            - default_xls: xls默认路径
            - export_format: "txt" 或 "xls" (可选,默认txt)
            - auto_open: 是否自动打开 (可选,默认True)
            - countdown_sec: 倒计时秒数 (可选,默认3)
            - settle_delay: 切换面板后等待时间 (可选,默认1.0)
            - window_key: 窗口关键字 (可选,默认"钱龙模拟期权宝")
            - use_title: 是否用title定位TreeItem (可选,默认False)
    """
    cfg = parse_env_config(config)

    print(f"目标: {cfg['name']} → 点击输出 → 格式={cfg['export_format'].upper()},自动打开={cfg['auto_open']}")
    print(f"窗口关键字: {cfg['window_key']}")
    print(f"输出路径: {cfg['output_path']}")
    countdown(cfg["countdown_sec"])

    hwnd = find_window(cfg["window_key"])
    print(f"[OK] 已找到窗口,句柄 = {hwnd}")

    win = activate_window(hwnd)

    # 切换面板
    switch_panel(win, cfg["panel_path"], use_title=cfg.get("use_title", False))
    time.sleep(cfg["settle_delay"])

    # 点击输出按钮
    if not click_output_button(win, button_auto_id="1159"):
        print("[错误] 无法点击'输出'按钮")
        sys.exit(2)

    # 确保输出目录存在
    output_dir = os.path.dirname(cfg["output_path"])
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
        print(f"[OK] 已创建输出目录: {output_dir}")

    # 处理数据输出弹窗
    ok = handle_export_dialog(
        timeout=8,
        export_type=cfg["export_format"],
        auto_open=cfg["auto_open"],
        output_path=cfg["output_path"]
    )
    if not ok:
        print("[WARN] '数据输出'弹窗处理异常,请手动确认")

    print(f"\n=== {cfg['name']}数据导出完成 ===")


def run_save_as(config: dict):
    """执行另存为对话框流程（资金持仓专用）。

    Args:
        config: 配置字典，需包含:
            - panel_path: 树形面板路径
            - name: 显示名称
            - default_xls: xls默认路径
            - countdown_sec: 倒计时秒数 (可选)
            - settle_delay: 切换面板后等待时间 (可选)
            - window_key: 窗口关键字 (可选)
            - use_title: 是否用title定位TreeItem (可选)
    """
    cfg = parse_env_config(config)

    print(f"目标: {cfg['name']} → 点击输出 → 保存为 xls")
    print(f"窗口关键字: {cfg['window_key']}")
    print(f"输出路径: {cfg['output_path']}")
    countdown(cfg["countdown_sec"])

    hwnd = find_window(cfg["window_key"])
    print(f"[OK] 已找到窗口,句柄 = {hwnd}")

    win = activate_window(hwnd)

    # 切换面板
    switch_panel(win, cfg["panel_path"], use_title=cfg.get("use_title", False))
    time.sleep(cfg["settle_delay"])

    # 点击输出按钮
    if not click_output_button(win, button_auto_id="1159"):
        print("[错误] 无法点击'输出'按钮")
        sys.exit(2)

    # 处理另存为对话框
    save_dir = os.path.dirname(cfg["output_path"])
    filename = os.path.basename(cfg["output_path"])
    ok = handle_save_as_dialog(
        save_dir=save_dir,
        filename=filename,
        timeout=8,
    )
    if not ok:
        print("[WARN] '另存为'对话框处理异常,请手动确认")
    else:
        print(f"\n=== {cfg['name']}数据导出完成 ===")
