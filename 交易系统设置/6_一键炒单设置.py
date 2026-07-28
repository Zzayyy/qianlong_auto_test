# -*- coding: utf-8 -*-
"""交易系统设置 - 一键炒单设置自动化检查。

打开“交易系统设置”，进入“一键炒单设置”，读取快捷键方案、沪深期权下单
价格类型、默认期权合约和完整快捷键表格，与独立标准配置比对并保存报告/截图。

快捷键表格（auto_id=2216）是自绘 ListView，通过分页截图并结构化 OCR 读取；
全程不发送快捷键、不点击“应用”或“恢复默认”。
"""

import os
import sys
import time
import ctypes
import json

# 国泰海通在高分屏上混用了逻辑坐标和物理坐标。必须在导入 pywinauto
# 之前声明 DPI 感知，否则菜单、导航项和截图都会发生坐标偏移。
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

import win32gui
import win32con
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.one_click_settings import (
    evaluate_shortcuts,
    merge_shortcut_pages,
    normalize_text,
    parse_shortcut_ocr_tokens,
)
from core.settings import SettingsTestResult
from core.settings_standard import load_standard


# GUI 启动时 core.window 会按 GUI_CLIENT_ID 覆盖此值；直接运行本脚本时，
# 使用本次控件采集来源（国泰海通）作为默认窗口关键字。
WINDOW_KEYWORD = "国泰海通证券期权宝"
SETTINGS_BUTTON_AUTO_ID = "1008"
SETTINGS_MENU_ITEM_AUTO_ID = "20025"
SETTINGS_DIALOG_TITLE = "交易系统设置"
PANEL_NAME = "一键炒单设置"

PROFILE_PATH = Path(__file__).with_name("一键炒单设置标准.json")
with PROFILE_PATH.open("r", encoding="utf-8-sig") as profile_file:
    VALIDATION_PROFILE = json.load(profile_file)

# 标准值（可比字段）：下拉框当前选择 + 默认合约输入值。
# 优先从 交易系统设置/标准/<客户端>/一键炒单设置.json 读取（抓取脚本可覆盖）；
# 找不到时用下方由 VALIDATION_PROFILE 推导的兜底，保证离线不崩。
# 注意（本面板特殊）：完整校验标准仍来自独立 JSON（一键炒单设置标准.json）的
# VALIDATION_PROFILE，含下拉候选项、默认合约、快捷键表格指纹与 OCR 配置；
# 此处 STANDARD_VALUES 仅承载会与界面“当前值”比对的扁平字段。
DEFAULT_STANDARD_VALUES = {
    key: value["selected"]
    for key, value in VALIDATION_PROFILE["dropdowns"].items()
}
DEFAULT_STANDARD_VALUES.update(VALIDATION_PROFILE["default_contracts"])

# 当前客户端（GUI 启动时由 GUI_CLIENT_ID 环境变量注入；空则用内嵌兜底）
CLIENT_ID = os.environ.get("GUI_CLIENT_ID", "") or ""
STANDARD_VALUES = load_standard(PANEL_NAME, CLIENT_ID, DEFAULT_STANDARD_VALUES)

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


def _find_visible_native_child(dlg, auto_id: str, class_name: str) -> int:
    matches: List[int] = []

    def _enum(hwnd, _):
        try:
            if (
                win32gui.GetDlgCtrlID(hwnd) == int(auto_id)
                and win32gui.GetClassName(hwnd) == class_name
                and win32gui.IsWindowVisible(hwnd)
            ):
                matches.append(hwnd)
        except Exception:
            pass

    win32gui.EnumChildWindows(int(dlg.handle), _enum, None)
    if len(matches) != 1:
        raise RuntimeError(
            f"控件(auto_id={auto_id}, class={class_name})匹配数量为 {len(matches)}"
        )
    return matches[0]


def _native_combobox_snapshot(dlg, auto_id: str) -> Dict[str, Any]:
    """Read selection and every candidate without opening or changing the combo."""
    hwnd = _find_visible_native_child(dlg, auto_id, "ComboBox")
    count = win32gui.SendMessage(hwnd, win32con.CB_GETCOUNT, 0, 0)
    selected_index = win32gui.SendMessage(hwnd, win32con.CB_GETCURSEL, 0, 0)
    if count < 0:
        raise RuntimeError(f"下拉框(auto_id={auto_id})返回无效项目数 {count}")

    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.SendMessageW.argtypes = [
        wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t
    ]
    user32.SendMessageW.restype = ctypes.c_ssize_t
    items: List[str] = []
    for index in range(count):
        length = user32.SendMessageW(hwnd, win32con.CB_GETLBTEXTLEN, index, 0)
        if length < 0:
            raise RuntimeError(f"无法读取下拉框(auto_id={auto_id})第 {index} 项长度")
        buffer = ctypes.create_unicode_buffer(max(int(length) + 1, 2))
        copied = user32.SendMessageW(
            hwnd, win32con.CB_GETLBTEXT, index, ctypes.addressof(buffer)
        )
        if copied < 0:
            raise RuntimeError(f"无法读取下拉框(auto_id={auto_id})第 {index} 项")
        items.append(normalize_text(buffer.value))

    current = items[selected_index] if 0 <= selected_index < len(items) else ""
    selected_after = win32gui.SendMessage(hwnd, win32con.CB_GETCURSEL, 0, 0)
    if selected_after != selected_index:
        raise RuntimeError("只读采集前后下拉框选择发生变化，已拒绝继续")
    return {
        "current": current,
        "items": items,
        "selected_index": selected_index,
        "source": "原生ComboBox消息（未展开、未修改）",
    }


def get_combobox_snapshot(dlg, auto_id: str) -> Dict[str, Any]:
    try:
        return _native_combobox_snapshot(dlg, auto_id)
    except Exception as native_error:
        # 旧版非标准 ComboBox 兜底；仅读取，不调用 select()。
        try:
            combo = dlg.child_window(auto_id=auto_id, control_type="ComboBox")
            combo.wait("ready", timeout=3)
            current = get_combobox_value(dlg, auto_id) or ""
            items = [normalize_text(value) for value in combo.item_texts() if value]
            return {
                "current": current,
                "items": items,
                "selected_index": None,
                "source": f"UIA只读兜底；原生失败: {native_error}",
            }
        except Exception as uia_error:
            raise RuntimeError(
                f"原生读取失败: {native_error}；UIA读取失败: {uia_error}"
            ) from uia_error


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


def _capture_hwnd_image(hwnd: int):
    """Capture an HWND with PrintWindow, independent of foreground/DPI coords."""
    import win32ui
    from PIL import Image

    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width, height = right - left, bottom - top
    if width <= 0 or height <= 0:
        raise RuntimeError(f"控件矩形无效: {(left, top, right, bottom)}")
    window_dc = win32gui.GetWindowDC(hwnd)
    source_dc = win32ui.CreateDCFromHandle(window_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    try:
        bitmap.CreateCompatibleBitmap(source_dc, width, height)
        memory_dc.SelectObject(bitmap)
        if not ctypes.windll.user32.PrintWindow(hwnd, memory_dc.GetSafeHdc(), 2):
            raise RuntimeError("PrintWindow 返回失败")
        info = bitmap.GetInfo()
        bits = bitmap.GetBitmapBits(True)
        return Image.frombuffer(
            "RGB", (info["bmWidth"], info["bmHeight"]),
            bits, "raw", "BGRX", 0, 1
        ).copy()
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, window_dc)


def _ocr_shortcut_pages(table_hwnd: int, row_count: int,
                        artifact_dir: str, timestamp: str) -> List[Dict[str, Any]]:
    import numpy as np
    from rapidocr import RapidOCR

    os.makedirs(artifact_dir, exist_ok=True)
    profile = VALIDATION_PROFILE["ocr"]
    original_top = win32gui.SendMessage(table_hwnd, 0x1027, 0, 0)  # LVM_GETTOPINDEX
    targets = [0]
    if row_count > 0:
        targets.append(row_count - 1)
    ocr = RapidOCR()
    pages: List[List[Dict[str, Any]]] = []
    captured_tops = set()
    try:
        for page_number, target in enumerate(targets, start=1):
            win32gui.SendMessage(table_hwnd, 0x1013, target, 0)  # LVM_ENSUREVISIBLE
            time.sleep(0.5)
            top_index = win32gui.SendMessage(table_hwnd, 0x1027, 0, 0)
            if top_index in captured_tops:
                continue
            captured_tops.add(top_index)
            image = _capture_hwnd_image(table_hwnd)
            image_path = os.path.join(
                artifact_dir, f"快捷键表格_{timestamp}_页{page_number}.png"
            )
            image.save(image_path)
            output = ocr(np.array(image))
            tokens: List[Dict[str, Any]] = []
            if output:
                for box, text, score in zip(output.boxes, output.txts, output.scores):
                    tokens.append(
                        {"box": box.tolist(), "text": text, "score": float(score)}
                    )
            pages.append(
                parse_shortcut_ocr_tokens(
                    tokens,
                    column_boundaries=tuple(profile["column_boundaries"]),
                    y_tolerance=float(profile["row_y_tolerance"]),
                    name_aliases=profile.get("name_aliases"),
                )
            )
            print(
                f"[OK] 快捷键表格第{page_number}页OCR完成: "
                f"top={top_index}, 截图={image_path}"
            )
    finally:
        win32gui.SendMessage(table_hwnd, 0x1013, max(int(original_top), 0), 0)
        time.sleep(0.2)
    return merge_shortcut_pages(pages)


def collect_shortcut_table(dlg, artifact_dir: str,
                           timestamp: str) -> Dict[str, Any]:
    """通过分页 OCR 读取快捷键表格全部行。"""
    try:
        table = dlg.child_window(auto_id=AUTO_ID["快捷键表格"], control_type="List")
        table.wait("exists", timeout=3)
        table_hwnd = int(table.wrapper_object().handle)
        row_count = win32gui.SendMessage(table_hwnd, 0x1004, 0, 0)  # LVM_GETITEMCOUNT
        header_hwnd = win32gui.SendMessage(table_hwnd, 0x101F, 0, 0)  # LVM_GETHEADER
        column_count = (
            win32gui.SendMessage(header_hwnd, 0x1200, 0, 0) if header_hwnd else 0
        )  # HDM_GETITEMCOUNT
        try:
            rows = _ocr_shortcut_pages(
                table_hwnd, row_count, artifact_dir, timestamp
            )
            return {
                "rows": rows,
                "source": "OCR分页",
                "row_count": row_count,
                "column_count": column_count,
                "error": "",
            }
        except Exception as ocr_error:
            return {
                "rows": [],
                "source": "采集失败",
                "row_count": row_count,
                "column_count": column_count,
                "error": f"OCR失败: {ocr_error}",
            }
    except Exception as e:
        print(f"  [WARN] 采集快捷键表格失败: {e}")
        return {
            "rows": [], "source": "采集失败", "row_count": -1,
            "column_count": -1, "error": str(e)
        }


def take_screenshot(dlg, save_path: str):
    try:
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        # pywinauto 直接按窗口句柄截图，避免高 DPI 下 UIA 逻辑坐标与
        # mss 物理坐标不一致而截到其他窗口区域。
        dlg.capture_as_image().save(save_path)
        print(f"[OK] 截图已保存: {save_path}")
    except Exception as e:
        print(f"[WARN] 截图失败: {e}")


def test_one_click_trading(dlg, result: SettingsTestResult,
                           artifact_dir: str, timestamp: str):
    print("\n--- 一键炒单设置检查 ---")
    for key, expected in VALIDATION_PROFILE["dropdowns"].items():
        # 下拉“当前值”的期望值优先用 STANDARD_VALUES（可被抓取脚本覆盖），
        # 候选项列表仍来自 VALIDATION_PROFILE（特殊，不随比对标准迁移）。
        expected_selected = STANDARD_VALUES.get(key, expected["selected"])
        try:
            snapshot = get_combobox_snapshot(dlg, AUTO_ID[key])
            result.add_result(
                f"{key}_当前值", snapshot["current"], expected_selected,
                snapshot["source"]
            )
            result.add_result(
                f"{key}_候选项列表",
                "、".join(snapshot["items"]),
                "、".join(expected["items"]),
                snapshot["source"]
            )
        except Exception as error:
            result.add_unverified(
                f"{key}_当前值", expected_selected, str(error)
            )
            result.add_unverified(
                f"{key}_候选项列表", "、".join(expected["items"]), str(error)
            )

    for key, expected in VALIDATION_PROFILE["default_contracts"].items():
        # 默认合约期望值优先用 STANDARD_VALUES（可被抓取脚本覆盖）。
        expected_contract = STANDARD_VALUES.get(key, expected)
        actual = get_edit_value(dlg, AUTO_ID[key])
        if actual is None:
            result.add_unverified(key, expected_contract, "Edit控件无法读取")
        else:
            result.add_result(key, actual, expected_contract, "UIA只读")

    table = collect_shortcut_table(dlg, artifact_dir, timestamp)
    fingerprint = VALIDATION_PROFILE["fingerprint"]
    result.add_result(
        "快捷键表格_行数", table["row_count"],
        fingerprint["shortcut_row_count"], table["source"]
    )
    result.add_result(
        "快捷键表格_列数", table["column_count"],
        fingerprint["shortcut_column_count"], table["source"]
    )
    if table["rows"]:
        checks = evaluate_shortcuts(
            VALIDATION_PROFILE["shortcuts"], table["rows"],
            source=table["source"],
            min_ocr_confidence=float(
                VALIDATION_PROFILE["ocr"]["minimum_confidence"]
            ),
        )
        for check in checks:
            result.add_status(
                check["name"], check["actual"], check["expected"],
                check["status"], check["detail"]
            )
    else:
        for expected in VALIDATION_PROFILE["shortcuts"]:
            result.add_unverified(
                f"快捷键[{expected['sequence']}]_{expected['name']}",
                expected["shortcut"], table["error"] or "表格未返回数据"
            )

    rows_text = "；".join(
        f"{row['sequence']}.{row['name']}={row['shortcut']}"
        for row in table["rows"]
    )
    result.add_observation(
        "快捷键表格",
        rows_text or "未采集到结构化行",
        f"来源：{table['source']}；只读采集，不发送快捷键、不修改设置",
    )
    result.add_observation("校验标准", str(PROFILE_PATH), "独立JSON配置")


def collect_current_settings(dlg) -> dict:
    """读取当前面板的可比设置值，返回与 STANDARD_VALUES 同构的扁平字典。

    供“抓取自定义标准”脚本把当前客户端界面值采集为新的比对标准。
    说明（本面板特殊）：
      - 完整校验标准仍来自独立 JSON（一键炒单设置标准.json）的 VALIDATION_PROFILE，
        含下拉候选项、默认合约、快捷键表格指纹与 OCR 配置；
        此处只采集会与 STANDARD_VALUES 比对的“下拉当前选择”与“默认合约输入值”，
        不采集 OCR 快捷键表格（采集重、且依赖专属 profile，不随比对标准迁移）。
      - 全程只读：不发送快捷键、不点击“应用”/“恢复默认”，不改变任何设置。
      - 钱龙客户端无此面板，调用方（抓取脚本）会先因切换面板失败而跳过本面板。
    """
    data: dict = {}

    # 下拉框当前选择（selected）
    for key in VALIDATION_PROFILE["dropdowns"]:
        try:
            snapshot = get_combobox_snapshot(dlg, AUTO_ID[key])
            data[key] = snapshot["current"]
        except Exception as e:
            print(f"  [WARN] 采集下拉框({key})失败: {e}")
            data[key] = ""

    # 默认合约输入框
    for key in VALIDATION_PROFILE["default_contracts"]:
        try:
            val = get_edit_value(dlg, AUTO_ID[key])
            data[key] = val if val is not None else ""
        except Exception as e:
            print(f"  [WARN] 采集默认合约({key})失败: {e}")
            data[key] = ""

    return data


def main():
    print("=" * 60)
    print("交易系统设置 - 一键炒单设置自动化测试")
    print("=" * 60)

    client_id = os.environ.get("GUI_CLIENT_ID") or ""
    if client_id == "qianlong":
        print("[不支持] 钱龙客户端没有'一键炒单设置'页面，已安全跳过。")
        return

    result = SettingsTestResult(PANEL_NAME, normalizer=normalize_text)
    hwnd = None
    dlg = None
    try:
        countdown(COUNTDOWN_SEC)
        hwnd = find_window(WINDOW_KEYWORD)
        print(f"[OK] 已找到主窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)

        dlg = open_settings_dialog_compat(
            win, SETTINGS_BUTTON_AUTO_ID, SETTINGS_MENU_ITEM_AUTO_ID, SETTINGS_DIALOG_TITLE
        )
        dlg.wait("ready", timeout=10)
        if not switch_settings_panel_compat(dlg, PANEL_NAME):
            raise RuntimeError(f"无法切换到'{PANEL_NAME}'面板")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        artifact_dir = os.path.join(OUTPUT_DIR, RESULT_SUBDIR)
        test_one_click_trading(dlg, result, artifact_dir, timestamp)

        screenshot_path = os.path.join(
            artifact_dir, f"{PANEL_NAME}_{timestamp}.png"
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
