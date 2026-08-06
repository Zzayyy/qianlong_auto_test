# -*- coding: utf-8 -*-
"""交易系统设置 - 一键炒单设置自动化检查。

打开“交易系统设置”，进入“一键炒单设置”，读取快捷键方案、沪深期权下单
价格类型、默认期权合约和完整快捷键表格，与自定义/独立标准配置比对并保存
报告/截图。

快捷键表格（auto_id=2216）是自绘 ListView，通过分页截图并结构化 OCR 读取；
该表格本身是可自定义的比对基线：抓取脚本会把当前表格存进
标准/<client>/一键炒单设置.json 的 “快捷键表格” 字段，用户可手改，
测试时按此自定义表格与当前客户端逐行比对（希望完全一致）。
全程不发送快捷键、不点击“应用”或“恢复默认”。
"""

import os
import sys
import time
import ctypes

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
from typing import Any, Dict, List, Optional

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.window import find_window, activate_window, countdown, close_settings_dialog
from core.settings_window import (
    open_settings_dialog as open_settings_dialog_compat,
    switch_settings_panel as switch_settings_panel_compat,
)
from core.one_click_settings import (
    canonical_hotkey,
    evaluate_shortcuts,
    filter_phantom_rows,
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

# ───────────────────────────────────────────────────────────────────────────
# 校验 / OCR 配置（原独立 JSON“一键炒单设置标准.json”已并入此处，与其他面板一致：
#  - 结构指纹、OCR 参数等硬编码为本文件常量；
#  - 可比期望值走 标准/<client>/一键炒单设置.json → 标准/default/一键炒单设置.json
#    → 下方内嵌兜底 三级优先级，由 core.settings_standard.load_standard 统一解析。）
# ───────────────────────────────────────────────────────────────────────────

# 下拉“当前值”的键（与候选项列表同源）。
DROPDOWN_KEYS = [
    "快捷键方案",
    "上海期权_买入开仓",
    "上海期权_卖出开仓",
    "上海期权_平仓",
    "深圳期权_买入开仓",
    "深圳期权_卖出开仓",
    "深圳期权_平仓",
]

# 快捷键表格结构指纹（原 fingerprint 中仍在使用的字段；导航相关字段已废弃）。
DEFAULT_SHORTCUT_ROW_COUNT = 12
DEFAULT_SHORTCUT_COL_COUNT = 3

# OCR 引擎参数（原 VALIDATION_PROFILE["ocr"]）；仅供快捷键表格分页识别使用，
# 与可比期望值无关，硬编码为常量。
OCR_MIN_CONFIDENCE = 0.75
OCR_COLUMN_BOUNDARIES = (50.0, 165.0)
OCR_ROW_Y_TOLERANCE = 6.0
OCR_NAME_ALIASES = {"全撒": "全撤"}

# 快捷键表格的离线兜底基线（原 VALIDATION_PROFILE["shortcuts"]）。
# 客户端标准（标准/<client>/一键炒单设置.json 的“快捷键表格”）优先；此处为最后兜底。
DEFAULT_SHORTCUT_ROWS = [
    {"sequence": 1, "name": "启用一键炒单", "shortcut": "CTRL+S"},
    {"sequence": 2, "name": "买合约1", "shortcut": "小键盘1"},
    {"sequence": 3, "name": "卖合约1", "shortcut": "小键盘2"},
    {"sequence": 4, "name": "平合约1权利仓", "shortcut": "小键盘3"},
    {"sequence": 5, "name": "平合约1义务仓", "shortcut": "小键盘4"},
    {"sequence": 6, "name": "买合约2", "shortcut": "小键盘5"},
    {"sequence": 7, "name": "卖合约2", "shortcut": "小键盘6"},
    {"sequence": 8, "name": "平合约2权利仓", "shortcut": "小键盘7"},
    {"sequence": 9, "name": "平合约2义务仓", "shortcut": "小键盘8"},
    {"sequence": 10, "name": "全撤", "shortcut": "小键盘."},
    {"sequence": 11, "name": "增加张数", "shortcut": "小键盘+"},
    {"sequence": 12, "name": "减少张数", "shortcut": "小键盘-"},
]

# 标准值（可比字段）：下拉框当前选择 + 各下拉候选项 + 默认合约输入值。
# 优先从 标准/<客户端>/一键炒单设置.json 读取（抓取脚本可覆盖）；
# 其次 标准/default/一键炒单设置.json；最后用下方内嵌兜底，保证离线不崩。
# 注：快捷键表格基线不放入 DEFAULT_STANDARD_VALUES（否则会被 CONTRACT_KEYS 推导
# 误当作合约键），单独用 DEFAULT_SHORTCUT_ROWS 常量承载。
DEFAULT_STANDARD_VALUES = {
    # 下拉当前选择
    "快捷键方案": "钱龙推荐快捷键方案",
    "上海期权_买入开仓": "对手价",
    "上海期权_卖出开仓": "对手价",
    "上海期权_平仓": "对手价",
    "深圳期权_买入开仓": "对手价",
    "深圳期权_卖出开仓": "对手价",
    "深圳期权_平仓": "对手价",
    # 各下拉候选项（与客户端一致；可被抓取脚本覆盖，与其他面板一致）
    "快捷键方案_候选项": ["钱龙推荐快捷键方案"],
    "上海期权_买入开仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "市价转限", "市价FAK", "市价F0K"],
    "上海期权_卖出开仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "市价转限", "市价FAK", "市价F0K"],
    "上海期权_平仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "市价转限", "市价FAK", "市价F0K"],
    "深圳期权_买入开仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "对方最优价格", "本方最优价格", "即时成交剩余撤销", "五档即成剩撤", "全额成交或撤销"],
    "深圳期权_卖出开仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "对方最优价格", "本方最优价格", "即时成交剩余撤销", "五档即成剩撤", "全额成交或撤销"],
    "深圳期权_平仓_候选项": ["对手价", "挂盘价", "涨停价", "跌停价", "限价", "超价", "对方最优价格", "本方最优价格", "即时成交剩余撤销", "五档即成剩撤", "全额成交或撤销"],
    # 默认合约输入值
    "默认期权合约1": "默认当前标的的当月认购平值期权",
    "默认期权合约2": "默认当前标的的当月认沽平值期权",
}
# 默认合约键（不含下拉键与候选项键），供采集/比对遍历
CONTRACT_KEYS = [
    k for k in DEFAULT_STANDARD_VALUES
    if k not in DROPDOWN_KEYS and not k.endswith("_候选项")
]

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
COUNTDOWN_SEC = int(os.environ.get("GUI_COUNTDOWN", "3"))








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


def _recover_key_via_crop(orig_image, anchor, ocr_scale: int, ocr) -> str:
    """当整表 OCR 把“小键盘”后的细按键(尤其减号-)漏检、只剩“小键盘”时，
    按 anchor(已上采样坐标系)定位到“小键盘”右侧的按键区，紧裁出小图、再高倍
    放大后用 use_det=False 强制识别器读取该单元格(参考超级策略读单元格的方法)，
    从结果尾部提取单个按键字符。返回键位字符(0-9/+-.)或空串。"""
    if anchor is None:
        return ""
    try:
        from PIL import Image as _PILImage
        import numpy as np
        import re as _re

        cx_raw = anchor[0] / ocr_scale
        cy_raw = anchor[1] / ocr_scale
        W, H = orig_image.size
        # 裁剪窗：紧贴“小键盘”右侧(从中心+6起，约覆盖末字“盘”到按键区)，
        # 纵向仅取单行高度。窗过宽会把注意力拉回“小键盘”本身而再次漏掉细按键。
        left = max(0, int(cx_raw + 6))
        right = min(W, int(cx_raw + 44))
        top = max(0, int(cy_raw - 12))
        bottom = min(H, int(cy_raw + 12))
        if right <= left or bottom <= top:
            return ""
        cell = orig_image.crop((left, top, right, bottom))
        zoom = 8
        big = cell.resize((cell.width * zoom, cell.height * zoom), _PILImage.LANCZOS)
        out = ocr(np.asarray(big), use_det=False, use_cls=False, use_rec=True)
        txts = list(getattr(out, "txts", ()) or ())
        scores = list(getattr(out, "scores", ()) or ())
        for text, score in zip(txts, scores):
            text = str(text)
            if float(score) < 0.5:
                continue
            # 键位在文本尾部：小键盘X / 盘X / 单独 X
            m = _re.search(r"[0-9+\-.]\s*$", text)
            if m:
                return m.group(0).strip()
    except Exception as exc:  # 复核失败不应阻断主流程
        print(f"  [WARN] 按键区复核失败: {exc}")
    return ""


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
    from PIL import Image as _PILImage
    from rapidocr import RapidOCR

    os.makedirs(artifact_dir, exist_ok=True)
    original_top = win32gui.SendMessage(table_hwnd, 0x1027, 0, 0)  # LVM_GETTOPINDEX
    # 估算可视行数，生成带重叠的多页采集目标，避免某行始终落在滚动边界而漏读。
    targets = [0]
    if row_count > 1:
        win32gui.SendMessage(table_hwnd, 0x1013, row_count - 1, 0)  # LVM_ENSUREVISIBLE
        time.sleep(0.3)
        top_last = win32gui.SendMessage(table_hwnd, 0x1027, 0, 0)
        visible = (row_count - 1) - top_last
        if visible < 1:
            visible = 1
        stride = max(1, visible // 2)
        for t in range(0, row_count, stride):
            if t not in targets:
                targets.append(t)
        if (row_count - 1) not in targets:
            targets.append(row_count - 1)
    # 上采样倍数：细按键(-/+/.)在原始分辨率下易被漏检，2x 显著提升召回。
    # 注意：图像放大后 OCR 坐标同步放大，列边界与行容差也按同倍率放大以保持几何一致。
    ocr_scale = 2
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
            orig_image = _capture_hwnd_image(table_hwnd)
            image = orig_image
            if ocr_scale != 1:
                image = image.resize(
                    (image.width * ocr_scale, image.height * ocr_scale),
                    _PILImage.LANCZOS,
                )
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
            page_rows = parse_shortcut_ocr_tokens(
                tokens,
                column_boundaries=tuple(
                    b * ocr_scale for b in OCR_COLUMN_BOUNDARIES),
                y_tolerance=float(OCR_ROW_Y_TOLERANCE) * ocr_scale,
                name_aliases=OCR_NAME_ALIASES,
            )
            # 键位复核：整表 OCR 漏掉细按键(尤其是减号-)时，对“只剩小键盘”的行
            # 紧裁其按键区做高倍强制识别，补回缺失的键位。
            for row in page_rows:
                anchor = row.pop("_anchor", None)
                if anchor is not None and row.get("shortcut") == "小键盘":
                    key = _recover_key_via_crop(orig_image, anchor, ocr_scale, ocr)
                    if key:
                        row["shortcut"] = canonical_hotkey("小键盘" + key)
                        row["source"] = "OCR+键位复核"
                        print(
                            f"  [复核] 行{row['sequence']} {row['name']} "
                            f"补回键位 -> {row['shortcut']}"
                        )
            pages.append(page_rows)
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
            # 过滤超出实际表格范围(1..row_count)的 OCR 幻影行(如误识的序号 40)，
            # 避免其以「标准配置中不存在」的形式污染比对结果。
            before = len(rows)
            rows = filter_phantom_rows(rows, row_count)
            dropped = before - len(rows)
            if dropped:
                print(
                    f"  [WARN] 丢弃 {dropped} 行 OCR 幻影行"
                    f"(序号超出 1..{row_count})"
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
    for key in DROPDOWN_KEYS:
        # 下拉“当前值”与“候选项列表”的期望值均来自 STANDARD_VALUES
        #（可被抓取脚本覆盖，也可手改，与其他面板一致）。
        expected_selected = STANDARD_VALUES.get(
            key, DEFAULT_STANDARD_VALUES.get(key, ""))
        expected_items = STANDARD_VALUES.get(
            f"{key}_候选项", DEFAULT_STANDARD_VALUES.get(f"{key}_候选项", []))
        try:
            snapshot = get_combobox_snapshot(dlg, AUTO_ID[key])
            result.add_result(
                f"{key}_当前值", snapshot["current"], expected_selected,
                snapshot["source"]
            )
            result.add_result(
                f"{key}_候选项列表",
                "、".join(snapshot["items"]),
                "、".join(expected_items),
                snapshot["source"]
            )
        except Exception as error:
            result.add_unverified(
                f"{key}_当前值", expected_selected, str(error)
            )
            result.add_unverified(
                f"{key}_候选项列表", "、".join(expected_items), str(error)
            )

    for key in CONTRACT_KEYS:
        # 默认合约期望值来自 STANDARD_VALUES（可被抓取脚本覆盖）。
        expected_contract = STANDARD_VALUES.get(
            key, DEFAULT_STANDARD_VALUES.get(key, ""))
        actual = get_edit_value(dlg, AUTO_ID[key])
        if actual is None:
            result.add_unverified(key, expected_contract, "Edit控件无法读取")
        else:
            result.add_result(key, actual, expected_contract, "UIA只读")

    table = collect_shortcut_table(dlg, artifact_dir, timestamp)
    # 比对基线优先用自定义标准里的“快捷键表格”；未自定义时回退内联兜底。
    custom_table = STANDARD_VALUES.get("快捷键表格")
    expected_shortcuts = (
        (custom_table or {}).get("rows") or DEFAULT_SHORTCUT_ROWS
    )
    _crow = (custom_table or {}).get("row_count")
    _ccol = (custom_table or {}).get("column_count")
    expected_row_count = (
        _crow if isinstance(_crow, int) and _crow >= 0
        else DEFAULT_SHORTCUT_ROW_COUNT
    )
    expected_col_count = (
        _ccol if isinstance(_ccol, int) and _ccol >= 0
        else DEFAULT_SHORTCUT_COL_COUNT
    )
    baseline_note = (
        "自定义标准(快捷键表格)" if custom_table else "固定校验配置(快捷键)"
    )
    result.add_result(
        "快捷键表格_行数", table["row_count"],
        expected_row_count, f"{table['source']};基线={baseline_note}"
    )
    result.add_result(
        "快捷键表格_列数", table["column_count"],
        expected_col_count, f"{table['source']};基线={baseline_note}"
    )
    if table["rows"]:
        checks = evaluate_shortcuts(
            expected_shortcuts, table["rows"],
            source=table["source"],
            min_ocr_confidence=float(OCR_MIN_CONFIDENCE),
            expected_row_count=expected_row_count,
        )
        for check in checks:
            result.add_status(
                check["name"], check["actual"], check["expected"],
                check["status"], check["detail"]
            )
    else:
        for expected in expected_shortcuts:
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
    result.add_observation(
        "校验标准", "标准目录/内置兜底",
        "比对外置到 标准/<client>/一键炒单设置.json；结构指纹与 OCR 参数内置于脚本",
    )


def collect_current_settings(dlg, artifact_dir: str = "", timestamp: str = "") -> dict:
    """读取当前面板的可比设置值，返回与 STANDARD_VALUES 同构的字典。

    供“抓取自定义标准”脚本把当前客户端界面值采集为新的比对标准。
    返回结构（与 STANDARD_VALUES 同构）：
      - 下拉当前选择（DROPDOWN_KEYS 各键 -> 字符串）
      - 默认合约输入值（CONTRACT_KEYS 各键 -> 字符串）
      - 快捷键表格（"快捷键表格" -> {"rows", "row_count",
        "column_count", "source", "error"}）：OCR 分页读取的全部行，
        作为可自定义的比对基线，用户可手改后让测试按此比对。
    说明：
      - 下拉候选项、默认合约、快捷键表格均为可自定义比对基线（存于
        标准/<client>/一键炒单设置.json，可被抓取脚本覆盖、也可手改）；
        结构指纹与 OCR 参数已内置于本脚本常量，不随比对标准迁移。
      - 全程只读：不发送快捷键、不点击“应用”/“恢复默认”，不改变任何设置。
      - 钱龙客户端无此面板，调用方（抓取脚本）会先因切换面板失败而跳过本面板。
    """
    data: dict = {}

    # 下拉框当前选择（selected）+ 候选项列表
    for key in DROPDOWN_KEYS:
        try:
            snapshot = get_combobox_snapshot(dlg, AUTO_ID[key])
            data[key] = snapshot["current"]
            data[f"{key}_候选项"] = snapshot["items"]
        except Exception as e:
            print(f"  [WARN] 采集下拉框({key})失败: {e}")
            data[key] = ""
            data[f"{key}_候选项"] = []

    # 默认合约输入框
    for key in CONTRACT_KEYS:
        try:
            val = get_edit_value(dlg, AUTO_ID[key])
            data[key] = val if val is not None else ""
        except Exception as e:
            print(f"  [WARN] 采集默认合约({key})失败: {e}")
            data[key] = ""

    # 快捷键表格（OCR 分页读取，作为可自定义比对基线）
    try:
        if not timestamp:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        if not artifact_dir:
            artifact_dir = os.path.join(OUTPUT_DIR, RESULT_SUBDIR)
        table = collect_shortcut_table(dlg, artifact_dir, timestamp)
        data["快捷键表格"] = {
            "rows": table["rows"],
            "row_count": table["row_count"],
            "column_count": table["column_count"],
            "source": table["source"],
            "error": table["error"],
        }
    except Exception as e:
        print(f"  [WARN] 采集快捷键表格失败: {e}")
        data["快捷键表格"] = {
            "rows": [], "row_count": -1, "column_count": -1,
            "source": "采集失败", "error": str(e),
        }

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