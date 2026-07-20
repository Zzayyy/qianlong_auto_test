# -*- coding: utf-8 -*-
"""交易系统设置 - 一键炒单设置自动化检查。

打开“交易系统设置”，进入“一键炒单设置”，读取快捷键方案、沪深期权下单
价格类型和默认期权合约，与国泰海通界面截图确认的标准值比对并保存报告/截图。

快捷键表格（auto_id=2216）是自绘 List，UI Automation 当前只暴露表头、不暴露
数据行，因此本脚本采集其可访问文本并在完整截图中留档，不点击表格或“恢复默认”。
"""

import os
import sys
import time
from datetime import datetime
from typing import Any, Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pywinauto import Application, findwindows

from core.window import find_window, activate_window, countdown, close_settings_dialog


# GUI 启动时 core.window 会按 GUI_CLIENT_ID 覆盖此值；直接运行本脚本时，
# 使用本次控件采集来源（国泰海通）作为默认窗口关键字。
WINDOW_KEYWORD = "国泰海通证券期权宝"
SETTINGS_BUTTON_AUTO_ID = "1008"
SETTINGS_MENU_ITEM_AUTO_ID = "20025"
SETTINGS_DIALOG_TITLE = "交易系统设置"
PANEL_NAME = "一键炒单设置"

# 首版标准值来自用户提供的国泰海通完整截图。
STANDARD_VALUES = {
    "快捷键方案": "钱龙推荐快捷键方案",
    "上海期权_买入开仓": "对手价",
    "上海期权_卖出开仓": "对手价",
    "上海期权_平仓": "对手价",
    "深圳期权_买入开仓": "对手价",
    "深圳期权_卖出开仓": "对手价",
    "深圳期权_平仓": "对手价",
    "默认期权合约1": "默认当前标的的当月认购平值期权",
    "默认期权合约2": "默认当前标的的当月认沽平值期权",
}

AUTO_ID = {
    "快捷键方案": "2100",
    "快捷键表格": "2216",
    "上海期权_买入开仓": "2078",
    "上海期权_卖出开仓": "2109",
    "上海期权_平仓": "2095",
    "深圳期权_买入开仓": "2079",
    "深圳期权_卖出开仓": "2110",
    "深圳期权_平仓": "2096",
    "默认期权合约1": "2154",
    "默认期权合约2": "2174",
}

_OUTPUT_DIR_DEFAULT = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "交易系统设置_测试结果",
)
OUTPUT_DIR = os.environ.get("GUI_OUTPUT_DIR", "") or _OUTPUT_DIR_DEFAULT
RESULT_SUBDIR = PANEL_NAME
COUNTDOWN_SEC = 3


class SettingsTestResult:
    def __init__(self):
        self.results: List[Dict[str, Any]] = []
        self.differences: List[Dict[str, Any]] = []
        self.observations: List[Dict[str, Any]] = []

    def add_result(self, name: str, actual_value: Any, expected_value: Any):
        matched = actual_value == expected_value
        row = {
            "名称": name,
            "期望值": expected_value,
            "实际值": actual_value,
            "是否一致": "✓" if matched else "✗ 差异",
        }
        self.results.append(row)
        if not matched:
            self.differences.append(row)

    def add_observation(self, name: str, value: Any, detail: str):
        self.observations.append({"名称": name, "采集值": value, "说明": detail})

    def print_summary(self):
        print(f"\n{'=' * 60}")
        print("测试结果汇总")
        print(f"{'=' * 60}")
        print(f"总比对项: {len(self.results)}")
        print(f"通过: {len(self.results) - len(self.differences)}")
        print(f"差异: {len(self.differences)}")
        print(f"采集项: {len(self.observations)}")
        for row in self.results:
            print(
                f"  {row['是否一致']} {row['名称']}: "
                f"期望={row['期望值']!r}, 实际={row['实际值']!r}"
            )
        for row in self.observations:
            print(f"  ○ {row['名称']}: {row['采集值']}（{row['说明']}）")

    def to_file(self, filepath: str):
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"{PANEL_NAME}测试报告\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'=' * 60}\n\n")
            f.write(f"总比对项: {len(self.results)}\n")
            f.write(f"通过: {len(self.results) - len(self.differences)}\n")
            f.write(f"差异: {len(self.differences)}\n")
            f.write(f"采集项: {len(self.observations)}\n\n")
            for row in self.results:
                f.write(
                    f"[{row['是否一致']}] {row['名称']}\n"
                    f"  期望值: {row['期望值']}\n"
                    f"  实际值: {row['实际值']}\n"
                )
            if self.observations:
                f.write("\n采集项（不计入差异）:\n")
                for row in self.observations:
                    f.write(
                        f"  - {row['名称']}: {row['采集值']}\n"
                        f"    说明: {row['说明']}\n"
                    )
        print(f"[OK] 测试报告已保存: {filepath}")


def _find_existing_settings_dlg(win=None) -> Optional[Any]:
    """兼容顶级窗口、#32770 对话框和主窗口子窗口。"""
    if win is not None:
        try:
            spec = win.child_window(title=SETTINGS_DIALOG_TITLE, control_type="Window")
            if spec.exists(timeout=0.5):
                dlg = spec.wrapper_object()
                dlg.wait("ready", timeout=3)
                print(f"[OK] 已找到设置对话框(子窗口): {SETTINGS_DIALOG_TITLE}")
                return dlg
        except Exception:
            pass

    try:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                title = elem.window_text() or ""
                if SETTINGS_DIALOG_TITLE in title:
                    dlg = Application(backend="uia").connect(handle=elem.handle).window(
                        handle=elem.handle
                    )
                    dlg.wait("ready", timeout=3)
                    print(f"[OK] 已找到设置对话框(顶级): {title}")
                    return dlg
            except Exception:
                continue
    except Exception:
        pass

    try:
        for elem in findwindows.find_elements(top_level_only=True, class_name="#32770"):
            try:
                dlg = Application(backend="uia").connect(
                    handle=elem.handle, timeout=1
                ).window(handle=elem.handle)
                title_bar = dlg.child_window(control_type="TitleBar")
                value = ""
                if title_bar.exists(timeout=0.3):
                    try:
                        value = title_bar.legacy_properties().get("Value", "") or ""
                    except Exception:
                        value = title_bar.element_info.name or ""
                if SETTINGS_DIALOG_TITLE in value:
                    dlg.wait("ready", timeout=3)
                    print(f"[OK] 已找到设置对话框(#32770): {value}")
                    return dlg
            except Exception:
                continue
    except Exception:
        pass
    return None


def open_settings_dialog(win):
    dlg = _find_existing_settings_dlg(win)
    if dlg is not None:
        return dlg

    print("\n正在通过工具栏按钮打开'交易系统设置'...")
    button = win.child_window(
        auto_id=SETTINGS_BUTTON_AUTO_ID, control_type="Button"
    )
    button.wait("enabled", timeout=5)
    button.click_input()
    time.sleep(0.5)

    end = time.time() + 4
    clicked = False
    while time.time() < end and not clicked:
        for elem in findwindows.find_elements(top_level_only=True):
            try:
                menu = Application(backend="uia").connect(
                    handle=elem.handle, timeout=0.5
                ).window(handle=elem.handle)
                item = menu.child_window(
                    auto_id=SETTINGS_MENU_ITEM_AUTO_ID, control_type="MenuItem"
                )
                if item.exists(timeout=0.2):
                    item.click_input()
                    clicked = True
                    break
            except Exception:
                continue
        if not clicked:
            time.sleep(0.15)

    end = time.time() + 10
    while time.time() < end:
        dlg = _find_existing_settings_dlg(win)
        if dlg is not None:
            return dlg
        time.sleep(0.3)
    raise RuntimeError("无法打开'交易系统设置'对话框")


def switch_to_settings_panel(dlg) -> bool:
    try:
        nav = dlg.child_window(auto_id="2210", control_type="List")
        nav.wait("ready", timeout=5)
        item = nav.child_window(title=PANEL_NAME, control_type="ListItem")
        item.wait("visible", timeout=3)
        item.click_input()
        time.sleep(0.6)
        print(f"[OK] 已切换到'{PANEL_NAME}'面板")
        return True
    except Exception as e:
        print(f"[WARN] 切换到'{PANEL_NAME}'面板失败: {e}")
        return False


def get_combobox_value(dlg, auto_id: str) -> Optional[str]:
    try:
        combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox")
        combo.wait("ready", timeout=3)
        try:
            value = combo.selected_text()
            if value and value.strip():
                return value.strip()
        except Exception:
            pass
        try:
            value = combo.legacy_properties().get("Value", "") or ""
            if value.strip():
                return value.strip()
        except Exception:
            pass
        return None
    except Exception as e:
        print(f"  [WARN] 获取下拉框(auto_id={auto_id})失败: {e}")
        return None


def get_edit_value(dlg, auto_id: str) -> Optional[str]:
    try:
        edit = dlg.child_window(auto_id=auto_id, control_type="Edit")
        edit.wait("ready", timeout=3)
        try:
            return (edit.get_value() or "").strip()
        except Exception:
            return (edit.window_text() or "").strip()
    except Exception as e:
        print(f"  [WARN] 获取文本框(auto_id={auto_id})失败: {e}")
        return None


def collect_shortcut_table_text(dlg) -> List[str]:
    """采集自绘快捷键表格能通过 UIA 暴露的文本，不点击、不滚动、不修改。"""
    try:
        table = dlg.child_window(auto_id=AUTO_ID["快捷键表格"], control_type="List")
        table.wait("exists", timeout=3)
        values: List[str] = []
        seen = set()
        for ctrl in table.descendants():
            try:
                text = (ctrl.window_text() or "").strip()
            except Exception:
                text = ""
            if text and text not in {"序号", "选项名称", "当前快捷键（点击设置）"}:
                if text not in seen:
                    seen.add(text)
                    values.append(text)
        return values
    except Exception as e:
        print(f"  [WARN] 采集快捷键表格失败: {e}")
        return []


def take_screenshot(dlg, save_path: str):
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        rect = dlg.rectangle()
        from mss import MSS
        from mss.tools import to_png

        monitor = {
            "top": int(rect.top),
            "left": int(rect.left),
            "width": int(rect.right - rect.left),
            "height": int(rect.bottom - rect.top),
        }
        with MSS() as sct:
            image = sct.grab(monitor)
            to_png(image.rgb, image.size, output=save_path)
        print(f"[OK] 截图已保存: {save_path}")
    except Exception as e:
        print(f"[WARN] 截图失败: {e}")


def test_one_click_trading(dlg, result: SettingsTestResult):
    print("\n--- 一键炒单设置检查 ---")
    for key in (
        "快捷键方案",
        "上海期权_买入开仓",
        "上海期权_卖出开仓",
        "上海期权_平仓",
        "深圳期权_买入开仓",
        "深圳期权_卖出开仓",
        "深圳期权_平仓",
    ):
        actual = get_combobox_value(dlg, AUTO_ID[key])
        result.add_result(key, actual if actual is not None else "(无法读取)", STANDARD_VALUES[key])

    for key in ("默认期权合约1", "默认期权合约2"):
        actual = get_edit_value(dlg, AUTO_ID[key])
        result.add_result(key, actual if actual is not None else "(无法读取)", STANDARD_VALUES[key])

    shortcut_text = collect_shortcut_table_text(dlg)
    result.add_observation(
        "快捷键表格",
        "、".join(shortcut_text) if shortcut_text else "UIA未暴露数据行，详见截图",
        "自绘List仅采集可访问文本，不点击、不修改快捷键",
    )


def main():
    print("=" * 60)
    print("交易系统设置 - 一键炒单设置自动化测试")
    print("=" * 60)

    result = SettingsTestResult()
    hwnd = None
    dlg = None
    try:
        countdown(COUNTDOWN_SEC)
        hwnd = find_window(WINDOW_KEYWORD)
        print(f"[OK] 已找到主窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)

        dlg = open_settings_dialog(win)
        dlg.wait("ready", timeout=10)
        if not switch_to_settings_panel(dlg):
            raise RuntimeError(f"无法切换到'{PANEL_NAME}'面板")

        test_one_click_trading(dlg, result)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        screenshot_path = os.path.join(
            OUTPUT_DIR, RESULT_SUBDIR, f"{PANEL_NAME}_{timestamp}.png"
        )
        take_screenshot(dlg, screenshot_path)
        result.print_summary()
        report_path = os.path.join(
            OUTPUT_DIR, RESULT_SUBDIR, f"{PANEL_NAME}测试报告_{timestamp}.txt"
        )
        result.to_file(report_path)
        print("\n=== 测试完成 ===")
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
            keep_open = os.environ.get("GUI_NEXT_CATEGORY", "") == "交易系统设置"
            close_ok = close_settings_dialog(
                dlg, keep_open=keep_open, main_hwnd=hwnd
            )
            if not close_ok:
                print("[WARN] 交易系统设置窗口未正常关闭，请确认后再执行后续非设置类任务")


if __name__ == "__main__":
    main()
