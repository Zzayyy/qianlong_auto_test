# -*- coding: utf-8 -*-
"""交易系统设置 - 一键炒单设置自动化检查。

打开“交易系统设置”，进入“一键炒单设置”，读取快捷键方案、沪深期权下单
价格类型、默认期权合约和完整快捷键表格，与独立标准配置比对并保存报告/截图。

快捷键表格（auto_id=2216）是自绘 ListView。脚本优先用原生控件读取，失败时
分页截图并结构化 OCR；全程不发送快捷键、不点击“应用”或“恢复默认”。
"""

import os
import sys
import time
import ctypes
import json
import struct

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
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from pywinauto import Application, findwindows

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.one_click_settings import (
    canonical_hotkey,
    evaluate_shortcuts,
    merge_shortcut_pages,
    normalize_text,
    parse_shortcut_ocr_tokens,
)
from core.native_tree import RemoteProcessMemory


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

STANDARD_VALUES = {
    key: value["selected"]
    for key, value in VALIDATION_PROFILE["dropdowns"].items()
}
STANDARD_VALUES.update(VALIDATION_PROFILE["default_contracts"])

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
        self.unverified: List[Dict[str, Any]] = []
        self.observations: List[Dict[str, Any]] = []

    def add_result(self, name: str, actual_value: Any, expected_value: Any,
                   detail: str = ""):
        matched = normalize_text(actual_value) == normalize_text(expected_value)
        self.add_status(
            name, actual_value, expected_value,
            "通过" if matched else "差异", detail
        )

    def add_status(self, name: str, actual_value: Any, expected_value: Any,
                   status: str, detail: str = ""):
        row = {
            "名称": name,
            "期望值": expected_value,
            "实际值": actual_value,
            "状态": status,
            "说明": detail,
        }
        self.results.append(row)
        if status in {"差异", "新增", "冲突"}:
            self.differences.append(row)
        elif status == "未验证":
            self.unverified.append(row)

    def add_unverified(self, name: str, expected_value: Any, detail: str,
                       actual_value: Any = "(无法确认)"):
        self.add_status(name, actual_value, expected_value, "未验证", detail)

    def add_observation(self, name: str, value: Any, detail: str):
        self.observations.append({"名称": name, "采集值": value, "说明": detail})

    def print_summary(self):
        counts = Counter(row["状态"] for row in self.results)
        print(f"\n{'=' * 60}")
        print("测试结果汇总")
        print(f"{'=' * 60}")
        print(f"总比对项: {len(self.results)}")
        print(f"通过: {counts['通过']}")
        print(f"差异/新增/冲突: {len(self.differences)}")
        print(f"未验证: {counts['未验证']}")
        print(f"采集项: {len(self.observations)}")

        # 与委托设置、期权设置等模块保持一致：汇总后逐项展示全部
        # 比对结果。候选项列表在控制台中截断显示，完整内容仍写入报告。
        if self.results:
            print(f"\n{'名称':<35} {'期望值':<25} {'实际值':<25} {'状态'}")
            print(f"{'-' * 100}")
            for row in self.results:
                name = str(row["名称"])[:33]
                expected = str(row["期望值"])[:23]
                actual = str(row["实际值"])[:23]
                print(
                    f"{name:<35} {expected:<25} {actual:<25} "
                    f"{row['状态']}"
                )

        for row in self.observations:
            print(f"  采集项 {row['名称']}: {row['说明']}")

    def to_file(self, filepath: str):
        counts = Counter(row["状态"] for row in self.results)
        os.makedirs(os.path.dirname(filepath), exist_ok=True)
        with open(filepath, "w", encoding="utf-8") as f:
            f.write(f"{PANEL_NAME}测试报告\n")
            f.write(f"生成时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"{'=' * 60}\n\n")
            f.write(f"总比对项: {len(self.results)}\n")
            f.write(f"通过: {counts['通过']}\n")
            f.write(f"差异/新增/冲突: {len(self.differences)}\n")
            f.write(f"未验证: {counts['未验证']}\n")
            f.write(f"采集项: {len(self.observations)}\n\n")
            for row in self.results:
                f.write(
                    f"[{row['状态']}] {row['名称']}\n"
                    f"  期望值: {row['期望值']}\n"
                    f"  实际值: {row['实际值']}\n"
                )
                if row["说明"]:
                    f.write(f"  说明: {row['说明']}\n")
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
        # 国泰海通的设置菜单是原生 #32768 弹出窗口。若先枚举并连接其他
        # 顶级窗口，菜单会因失去焦点而关闭，因此先用 Win32 直接锁定它。
        menu_handles = []

        def _collect_popup_menu(hwnd, _):
            try:
                if (
                    win32gui.IsWindowVisible(hwnd)
                    and win32gui.GetClassName(hwnd) == "#32768"
                ):
                    menu_handles.append(hwnd)
            except Exception:
                pass

        win32gui.EnumWindows(_collect_popup_menu, None)
        for handle in menu_handles:
            try:
                menu = Application(backend="uia").connect(
                    handle=handle, timeout=0.5
                ).window(handle=handle)
                item = menu.child_window(
                    auto_id=SETTINGS_MENU_ITEM_AUTO_ID, control_type="MenuItem"
                )
                if item.exists(timeout=0.2):
                    # 高 DPI 下该原生菜单的 UIA 坐标可能按逻辑像素返回，
                    # click_input 会点击到错误位置；Invoke 不依赖屏幕坐标。
                    item.invoke()
                    clicked = True
                    break
            except Exception:
                continue

        # 兼容少数将菜单实现为普通顶级窗口的旧版本。
        if not clicked:
            for elem in findwindows.find_elements(top_level_only=True):
                try:
                    if elem.class_name == "#32768":
                        continue
                    menu = Application(backend="uia").connect(
                        handle=elem.handle, timeout=0.5
                    ).window(handle=elem.handle)
                    item = menu.child_window(
                        auto_id=SETTINGS_MENU_ITEM_AUTO_ID,
                        control_type="MenuItem",
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
        # 直接向原生 ListBox 设置选择并发送 LBN_SELCHANGE。该方式与用户
        # 选择列表项触发相同通知，同时绕开 QLOption 的 DPI 坐标虚拟化。
        items = nav.descendants(control_type="ListItem")
        target_index = next(
            i for i, ctrl in enumerate(items)
            if (ctrl.window_text() or "").strip() == PANEL_NAME
        )
        nav_wrapper = nav.wrapper_object()
        list_hwnd = int(nav_wrapper.handle)
        parent_hwnd = win32gui.GetParent(list_hwnd)
        control_id = win32gui.GetDlgCtrlID(list_hwnd)
        win32gui.SendMessage(list_hwnd, 0x0186, target_index, 0)  # LB_SETCURSEL
        win32gui.SendMessage(
            parent_hwnd,
            win32con.WM_COMMAND,
            control_id | (1 << 16),  # LBN_SELCHANGE == 1
            list_hwnd,
        )
        time.sleep(0.8)
        # 用本面板的唯一控件确认页面确实完成切换，避免只记录“点击成功”。
        dlg.child_window(
            auto_id=AUTO_ID["快捷键方案"], control_type="ComboBox"
        ).wait("exists", timeout=5)
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


def _read_listview_native(table_hwnd: int, row_count: int) -> List[Dict[str, Any]]:
    """Read ListView text with bounded messages when privileges allow it."""
    lvif_text = 0x0001
    lvm_getitemw = 0x104B

    def _pack_item(row_index: int, column: int, text_address: int,
                   target_bits: int) -> bytes:
        size = 88 if target_bits == 64 else 60
        package = bytearray(size)
        struct.pack_into("<IiiII", package, 0, lvif_text, row_index, column, 0, 0)
        if target_bits == 64:
            struct.pack_into("<Q", package, 24, text_address)
            struct.pack_into("<i", package, 32, 1024)
        else:
            struct.pack_into("<I", package, 20, text_address & 0xFFFFFFFF)
            struct.pack_into("<i", package, 24, 1024)
        return bytes(package)

    rows: List[Dict[str, Any]] = []
    with RemoteProcessMemory(table_hwnd) as memory:
        struct_size = 88 if memory.target_bits == 64 else 60
        text_size = 2048
        remote_address = memory.allocate(struct_size + text_size)
        text_address = remote_address + struct_size
        try:
            for row_index in range(row_count):
                values: List[str] = []
                for column in range(3):
                    package = bytearray(struct_size + text_size)
                    item = _pack_item(
                        row_index, column, text_address, memory.target_bits
                    )
                    package[:len(item)] = item
                    memory.write(remote_address, bytes(package))
                    response = win32gui.SendMessageTimeout(
                        table_hwnd, lvm_getitemw, 0, remote_address,
                        win32con.SMTO_ABORTIFHUNG, 3000
                    )
                    result = response[1] if isinstance(response, tuple) else response
                    if not result:
                        raise RuntimeError(
                            f"ListView第{row_index + 1}行第{column + 1}列读取失败"
                        )
                    raw = memory.read(text_address, text_size)
                    values.append(
                        normalize_text(
                            raw.decode("utf-16-le", errors="replace").split("\0", 1)[0]
                        )
                    )
                if not values[0].isdigit():
                    raise RuntimeError(
                        f"第 {row_index + 1} 行原生文本无效: {values!r}"
                    )
                rows.append(
                    {
                        "sequence": int(values[0]),
                        "name": values[1],
                        "shortcut": canonical_hotkey(values[2]),
                        "confidence": 1.0,
                        "source": "原生ListView",
                    }
                )
        finally:
            memory.free(remote_address)
    return rows


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
    """Read all rows, with native text first and paged OCR as fallback."""
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
            rows = _read_listview_native(table_hwnd, row_count)
            return {
                "rows": rows,
                "source": "原生ListView",
                "row_count": row_count,
                "column_count": column_count,
                "error": "",
            }
        except Exception as native_error:
            print(f"[INFO] 原生ListView文字读取不可用，改用分页OCR: {native_error}")
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
                    "error": f"原生失败: {native_error}；OCR失败: {ocr_error}",
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
        try:
            snapshot = get_combobox_snapshot(dlg, AUTO_ID[key])
            result.add_result(
                f"{key}_当前值", snapshot["current"], expected["selected"],
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
                f"{key}_当前值", expected["selected"], str(error)
            )
            result.add_unverified(
                f"{key}_候选项列表", "、".join(expected["items"]), str(error)
            )

    for key, expected in VALIDATION_PROFILE["default_contracts"].items():
        actual = get_edit_value(dlg, AUTO_ID[key])
        if actual is None:
            result.add_unverified(key, expected, "Edit控件无法读取")
        else:
            result.add_result(key, actual, expected, "UIA只读")

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


def main():
    print("=" * 60)
    print("交易系统设置 - 一键炒单设置自动化测试")
    print("=" * 60)

    client_id = os.environ.get("GUI_CLIENT_ID") or ""
    if client_id == "qianlong":
        print("[不支持] 钱龙客户端没有'一键炒单设置'页面，已安全跳过。")
        return

    result = SettingsTestResult()
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
