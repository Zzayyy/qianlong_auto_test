# -*- coding: utf-8 -*-
import sys
import os
_here = os.path.dirname(os.path.abspath(__file__))
for _d in (_here, os.path.dirname(_here), os.path.dirname(os.path.dirname(_here))):
    if os.path.isdir(os.path.join(_d, "core")) and _d not in sys.path:
        sys.path.insert(0, _d)
        break
from core.window import find_window, activate_window

"""
Excel驱动 + RapidOCR定位 + 表格平仓操作
============================================================
功能:
    1. 读取 Excel 表格(字段: 合约代码, 持仓类别, 平33%, 平50%, 平100%, 反手)
    2. 遍历每一行,根据持仓类别定位合约并点击
    3. 执行对应的平仓操作(平33%/平50%/平100%/反手 四选一)

依赖安装:
    pip install openpyxl pandas pywinauto opencv-python numpy mss pillow rapidocr

使用方法:
    1. 打开钱龙旗舰版,登录交易账号,切到"期权下单(新)"
    2. 准备好 Excel 文件
    3. 修改下方 EXCEL_PATH 或通过环境变量 GUI_XLSX_FILE 传入
    4. 运行本脚本
"""

import os
import re
import sys
import time
import ctypes

import cv2
import numpy as np
import mss
from PIL import Image
import pandas as pd
from pywinauto import Application, findwindows
from core.window import switch_panel


# ====================== 可配置参数 ======================
WINDOW_KEY = "钱龙模拟期权宝"

TREE_ITEM = "期权下单(新)"
TABLE_AUTO_ID = "3000"      # 定位表格的 auto_id
TABLE_OP_AUTO_ID = "1170"    # 操作表格的 auto_id(持仓表格)
CLICK_COORDS = (0, 100)      # 定位表格内部相对点击坐标
COUNTDOWN = 3

# ---- Excel 路径 ----
# 优先使用环境变量(GUI 调用传入);未设置时用默认路径(脚本同目录),方便直接运行
DEFAULT_EXCEL_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), r"C:\Users\Administrator\Desktop\平仓-反手模板.xlsx"
)
_gui_xlsx = os.environ.get("GUI_XLSX_FILE", "").strip()
if not _gui_xlsx:
    print(f"[INFO] 未设置 GUI_XLSX_FILE 环境变量,使用默认路径: {DEFAULT_EXCEL_PATH}")
    _gui_xlsx = DEFAULT_EXCEL_PATH
if not os.path.exists(_gui_xlsx):
    print(f"[错误] Excel 文件不存在: {_gui_xlsx}")
    print("请修改 DEFAULT_EXCEL_PATH,或在 GUI 中选择 Excel 配置文件")
    sys.exit(1)
EXCEL_PATH = _gui_xlsx

SCROLL_UP_TIMES = 3          # 初始:上滚回顶部
SCROLL_DOWN_TIMES = 1        # 每轮:未命中时下滚多少
STOP_KEYWORD = "统计"        # 表格底部标志
MAX_ROUNDS = 4               # 翻找最大轮次

# ---- 图像预处理参数 ----
H_LINE_FRACTION = 8
V_LINE_FRACTION = 8
LINE_DILATE_KERNEL = (2, 2)
ENLARGE_FACTOR = 2
BIN_METHOD = "otsu"
ADAPTIVE_BLOCK = 31
ADAPTIVE_C = 10

# ---- RapidOCR 参数 ----
# RapidOCR 默认使用 ONNXRuntime CPU + PP-OCRv4 中文模型
# 如需自定义配置，可生成 default_rapidocr.yaml 并通过 config_path 传入
OCR_MIN_CONF = 0.30
OCR_FUZZY_DIGITS = 1

# ---- 输出路径 ----
DEBUG_IMAGE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "table_screenshot_rapid.png"
)
ORIGINAL_IMAGE_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "table_screenshot_original.png"
)

# ---- 平仓按钮配置 ----
POSITION_BUTTONS = {
    "平33%":  "9039",
    "平50%":  "9040",
    "平100%": "9041",
    "反手":   "9042",
}
# ========================================================





def get_table(win, auto_id: str):
    table = win.child_window(auto_id=auto_id, control_type="Button")
    table.wait("ready", timeout=10)
    rect = table.rectangle()
    print(f"[INFO] 表格: auto_id={auto_id} "
          f"区域=({rect.left},{rect.top})-({rect.right},{rect.bottom}) "
          f"尺寸={rect.width()}x{rect.height()}")
    return rect, table


# ---------- 鼠标 ----------
def move_mouse(x: int, y: int):
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.05)


def click_screen(x: int, y: int):
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.05)
    ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)
    ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)
    print(f"[OK] 已在屏幕坐标 ({x}, {y}) 单击")


def scroll_wheel(delta: int = 120, times: int = 1):
    for i in range(times):
        ctypes.windll.user32.mouse_event(0x0800, 0, 0, delta, 0)
        direction = "上" if delta > 0 else "下"
        print(f"[..] 滚轮{direction} 第 {i + 1}/{times} 次")
        time.sleep(0.3)


def grab(rect, padding: int = 2) -> np.ndarray:
    bbox = {
        "left": rect.left + padding,
        "top": rect.top + padding,
        "width": max(1, rect.width() - 2 * padding),
        "height": max(1, rect.height() - 2 * padding),
    }
    with mss.mss() as sct:
        sct_img = sct.grab(bbox)
        rgb = np.array(sct_img)[:, :, :3][:, :, ::-1].copy()
    return rgb


# ---------- 图像预处理 ----------
def remove_table_lines(gray: np.ndarray) -> np.ndarray:
    h, w = gray.shape
    binary = cv2.adaptiveThreshold(
        gray, 255,
        cv2.ADAPTIVE_THRESH_MEAN_C,
        cv2.THRESH_BINARY_INV,
        31, 10,
    )
    h_w = max(20, w // H_LINE_FRACTION)
    h_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (h_w, 1))
    h_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, h_kernel)
    v_h = max(20, h // V_LINE_FRACTION)
    v_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, v_h))
    v_lines = cv2.morphologyEx(binary, cv2.MORPH_OPEN, v_kernel)
    line_mask = cv2.add(h_lines, v_lines)
    dk = cv2.getStructuringElement(cv2.MORPH_RECT, LINE_DILATE_KERNEL)
    line_mask = cv2.dilate(line_mask, dk, iterations=1)
    cleaned = gray.copy()
    cleaned[line_mask > 0] = 255
    return cleaned


def enlarge(img: np.ndarray, factor: float = 2.0) -> np.ndarray:
    h, w = img.shape[:2]
    new_size = (int(w * factor), int(h * factor))
    return cv2.resize(img, new_size, interpolation=cv2.INTER_CUBIC)


def binarize(gray: np.ndarray) -> np.ndarray:
    if BIN_METHOD == "adaptive":
        return cv2.adaptiveThreshold(
            gray, 255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            ADAPTIVE_BLOCK, ADAPTIVE_C,
        )
    _, b = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
    return b


def preprocess(img_rgb: np.ndarray) -> np.ndarray:
    gray = cv2.cvtColor(img_rgb, cv2.COLOR_RGB2GRAY)
    print(f"[..]   1) 灰度化 -> {gray.shape}")
    cleaned = remove_table_lines(gray)
    print("[..]   2) 表格横竖线已抹除")
    enlarged = enlarge(cleaned, ENLARGE_FACTOR)
    print(f"[..]   3) 放大 {ENLARGE_FACTOR}x -> {enlarged.shape}")
    binary = binarize(enlarged)
    print(f"[..]   4) 二值化完成 (方法={BIN_METHOD})")
    return binary


# ---------- RapidOCR ----------
_rapid_ocr_instance = None


def get_rapid_ocr():
    """初始化 RapidOCR 实例（单例模式）"""
    global _rapid_ocr_instance
    if _rapid_ocr_instance is not None:
        return _rapid_ocr_instance

    try:
        from rapidocr import RapidOCR
    except ImportError as e:
        print(f"[致命错误] rapidocr 导入失败: {e}")
        print("请执行: pip install rapidocr")
        sys.exit(1)

    print("[..] 正在初始化 RapidOCR (默认ONNXRuntime CPU + PP-OCRv4)...")
    # 默认配置即可满足中文表格识别需求
    # 如需指定本地模型或切换引擎，可通过 config_path 或参数传入
    _rapid_ocr_instance = RapidOCR()
    print("[OK] RapidOCR 初始化完成")
    return _rapid_ocr_instance


def ocr_image(img_rgb: np.ndarray):
    """
    使用 RapidOCR 识别图像
    返回格式与原版兼容: [{"text": str, "box": [[x,y]*4], "conf": float}, ...]
    """
    ocr = get_rapid_ocr()
    
    # RapidOCR 直接接受 numpy array (BGR或RGB均可，内部会处理)
    # 返回值是 RapidOCROutput dataclass
    result = ocr(img_rgb)
    
    out = []
    # 根据知识库: result 可能为 None 或空
    if not result or not result.txts:
        return out

    # result.boxes: (N, 4, 2) ndarray
    # result.txts: Tuple[str]
    # result.scores: Tuple[float]
    for i in range(len(result.txts)):
        text = result.txts[i]
        score = float(result.scores[i])
        box = result.boxes[i]  # shape (4, 2)

        if score < OCR_MIN_CONF:
            continue
        
        text = (text or "").strip()
        if not text:
            continue

        # 转换为整数坐标列表 [[x1,y1], [x2,y2], [x3,y3], [x4,y4]]
        box4 = [
            (int(box[0][0]), int(box[0][1])),
            (int(box[1][0]), int(box[1][1])),
            (int(box[2][0]), int(box[2][1])),
            (int(box[3][0]), int(box[3][1])),
        ]
        
        out.append({"text": text, "box": box4, "conf": score})
    
    return out


# ---------- OCR 工具函数 ----------
def box_center(box):
    xs = [p[0] for p in box]
    ys = [p[1] for p in box]
    return (sum(xs) / 4.0, sum(ys) / 4.0)


def normalize_digits(s: str) -> str:
    return re.sub(r"\D", "", s)


def find_contract_token(tokens, contract_code: str):
    pat = re.compile(rf"(?<!\d){re.escape(contract_code)}(?!\d)")
    for tok in tokens:
        if pat.search(tok["text"]):
            return tok, "strict"
    if OCR_FUZZY_DIGITS >= 1:
        target_digits = contract_code
        for tok in tokens:
            d = normalize_digits(tok["text"])
            if not d:
                continue
            if abs(len(d) - len(target_digits)) > OCR_FUZZY_DIGITS:
                continue
            common = sum(a == b for a, b in zip(d, target_digits))
            need = max(len(d), len(target_digits)) - OCR_FUZZY_DIGITS
            if common >= need:
                return tok, f"fuzzy(d={d})"
    return None, None


def group_tokens_by_row(tokens, y_tol: int = 12):
    if not tokens:
        return []
    sorted_toks = sorted(
        tokens, key=lambda t: (box_center(t["box"])[1], box_center(t["box"])[0])
    )
    rows = [[sorted_toks[0]]]
    for tok in sorted_toks[1:]:
        cy = box_center(tok["box"])[1]
        last_cy = box_center(rows[-1][-1]["box"])[1]
        if abs(cy - last_cy) <= y_tol:
            rows[-1].append(tok)
        else:
            rows.append([tok])
    return rows


def detect_column_centers(tokens):
    code_cx, pos_cx = None, None
    for tok in tokens:
        text = tok["text"].replace(" ", "")
        cx, _ = box_center(tok["box"])
        if code_cx is None and ("合约代码" in text
                                 or ("合约" in text and "代码" in text)):
            code_cx = cx
        if pos_cx is None and ("持仓类别" in text
                               or ("持仓" in text and "类别" in text)):
            pos_cx = cx
    return code_cx, pos_cx


def find_target_in_columns(tokens, contract_code: str,
                            target_position_type: str,
                            x_tol: int = 80, y_tol: int = 15):
    code_cx, _pos_cx_unused = detect_column_centers(tokens)
    use_x_filter = (code_cx is not None)
    pat_strict = re.compile(rf"(?<!\d){re.escape(contract_code)}(?!\d)")
    target_digits = contract_code

    candidates = []
    # 严格匹配
    for tok in tokens:
        cx, _ = box_center(tok["box"])
        if use_x_filter and abs(cx - code_cx) > x_tol:
            continue
        if pat_strict.search(tok["text"]):
            candidates.append((tok, "column-strict"))

    # 模糊匹配
    if not candidates and OCR_FUZZY_DIGITS >= 1:
        for tok in tokens:
            cx, _ = box_center(tok["box"])
            if use_x_filter and abs(cx - code_cx) > x_tol:
                continue
            d = normalize_digits(tok["text"])
            if not d:
                continue
            if abs(len(d) - len(target_digits)) > OCR_FUZZY_DIGITS:
                continue
            common = sum(a == b for a, b in zip(d, target_digits))
            need = max(len(d), len(target_digits)) - OCR_FUZZY_DIGITS
            if common >= need:
                candidates.append((tok, f"column-fuzzy(d={d})"))

    if not candidates:
        return None, None, None

    candidates.sort(key=lambda item: box_center(item[0]["box"])[1])

    def nearest_position(cx_code: float, cy_code: float):
        actual_pos = None
        best_dist = None
        for tok in tokens:
            text = tok["text"]
            if not ("权利仓" in text or "义务仓" in text):
                continue
            cx, cy = box_center(tok["box"])
            if abs(cy - cy_code) > y_tol:
                continue
            dx = cx - cx_code
            if dx <= 0:
                continue
            if best_dist is None or dx < best_dist:
                best_dist = dx
                actual_pos = "权利仓" if "权利仓" in text else "义务仓"
        return actual_pos

    last_tok, last_mode, last_pos = None, None, None
    for tok, mode in candidates:
        cx, cy = box_center(tok["box"])
        actual_pos = nearest_position(cx, cy)
        last_tok, last_mode, last_pos = tok, mode, actual_pos
        if not target_position_type:
            return tok, actual_pos, mode
        if actual_pos == target_position_type:
            return tok, actual_pos, mode

    if last_tok is not None:
        return last_tok, last_pos, last_mode
    return None, None, None


def has_keyword(tokens, keyword: str) -> bool:
    for tok in tokens:
        if keyword in tok["text"]:
            return True
    all_text = "".join((t["text"] or "") for t in tokens)
    cleaned = re.sub(r"\s+", "", all_text)
    if keyword in cleaned:
        return True
    pattern = r"[^\s]*?".join(re.escape(c) for c in keyword)
    return re.search(pattern, cleaned) is not None


# ---------- 平仓操作 ----------
def click_table_first_row(win, auto_id: str = TABLE_OP_AUTO_ID,
                           coords: tuple = CLICK_COORDS):
    table = win.child_window(auto_id=auto_id, control_type="Button")
    table.wait("ready", timeout=10)
    rect = table.rectangle()
    print(f"[INFO] 操作表格: auto_id={auto_id}, "
          f"尺寸={rect.width()}x{rect.height()}")
    table.click_input(coords=coords)
    print(f"[OK] 已在表格内点击坐标 {coords},定位到第一行")


def click_position_button(win, name: str, auto_id: str):
    try:
        btn = win.child_window(title=name, auto_id=auto_id, control_type="Button")
        btn.wait("ready", timeout=2)
    except Exception as e:
        print(f"[--] 按钮 '{name}' (auto_id={auto_id}) 不存在,跳过: {e}")
        return False

    if not btn.exists():
        print(f"[--] 按钮 '{name}' 不存在,跳过")
        return False

    try:
        if not btn.is_visible():
            print(f"[--] 按钮 '{name}' 不可见,跳过")
            return False
    except Exception:
        pass

    try:
        enabled = btn.is_enabled()
    except Exception as e:
        print(f"[--] 按钮 '{name}' 状态读取失败,跳过: {e}")
        return False

    if not enabled:
        print(f"[--] 按钮 '{name}' 为灰色(不可用),跳过")
        return False

    btn.click_input()
    print(f"[OK] 已点击按钮: {name} (auto_id={auto_id})")
    return True


def press_enter_to_confirm(main_win=None, dialog_patterns=None, timeout: float = 3):
    """在主窗口所属进程内寻找确认弹窗并回车。

    通过进程 ID 过滤，避免把回车发送到其它程序窗口
    (例如打开着的 "平仓.xlsx - Excel" 也会命中 "平仓" 关键字)。
    """
    if dialog_patterns is None:
        dialog_patterns = ["提示", "确认", "期权下单", "平仓", "反手"]

    # 取主窗口句柄与进程 ID，用于过滤
    main_hwnd = None
    main_pid = None
    if main_win is not None:
        try:
            main_hwnd = main_win.handle
        except Exception:
            main_hwnd = None
        try:
            main_pid = main_win.process_id()
        except Exception:
            if main_hwnd is not None:
                pid = ctypes.c_ulong(0)
                ctypes.windll.user32.GetWindowThreadProcessId(
                    main_hwnd, ctypes.byref(pid))
                main_pid = pid.value or None

    # 高优先级关键字(明确是弹窗)优先于宽泛关键字(平仓/反手)
    high_priority = ("提示", "确认", "确定", "委托确认", "风险提示")

    end = time.time() + timeout
    while time.time() < end:
        try:
            elems = findwindows.find_elements(top_level_only=True)
        except Exception:
            elems = []

        high_candidates = []
        low_candidates = []
        for elem in elems:
            # 1) 进程过滤: 只处理与主窗口同进程的窗口
            if main_pid is not None:
                try:
                    if elem.process_id != main_pid:
                        continue
                except Exception:
                    continue
            hwnd = elem.handle
            # 2) 跳过主窗口本身
            if main_hwnd is not None and hwnd == main_hwnd:
                continue
            try:
                title = (elem.name or "").strip()
            except Exception:
                title = ""
            if not any(p in title for p in dialog_patterns):
                continue
            if any(p in title for p in high_priority):
                high_candidates.append((hwnd, title))
            else:
                low_candidates.append((hwnd, title))

        for hwnd, title in high_candidates + low_candidates:
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                dlg.set_focus()
                time.sleep(0.3)
                dlg.type_keys("{ENTER}", with_spaces=False)
                print(f"[OK] 回车确认 (hwnd={hwnd}, title='{title}', pid={main_pid})")
                return True
            except Exception as e:
                print(f"[--] 弹窗(hwnd={hwnd}, title='{title}')回车失败: {e}")
                continue

        time.sleep(1)
    print(f"[WARN] 等待弹窗超时({timeout}s)")
    return False


def confirm_all_dialogs(
    main_win=None,
    max_dialogs: int = 5,
    no_dialog_timeout: float = 2.0,
    per_dialog_timeout: float = 4.0,
):
    count = 0
    for i in range(1, max_dialogs + 1):
        print(f"[..] 等待第 {i} 个弹窗 (超时{no_dialog_timeout}s无新弹窗则结束)...")
        ok = press_enter_to_confirm(main_win=main_win, timeout=per_dialog_timeout)
        if ok:
            count += 1
            print(f"[OK] 已确认第 {count} 个弹窗")
            time.sleep(0.4)
        else:
            print(f"[OK] 无更多弹窗，共确认 {count} 个")
            break
    else:
        print(f"[WARN] 达到最大弹窗数量上限({max_dialogs})")


# ---------- Excel 读取 ----------
def read_excel(filepath: str) -> list:
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Excel 文件不存在: {filepath}")

    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip()

    required_cols = ["合约代码", "持仓类别", "平33%", "平50%", "平100%", "反手"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Excel 缺少必需列: {col}")

    rows = []
    truthy = {"TRUE", "1", "YES", "Y", "是", "√", "✓", "X", "T"}
    
    for idx, row in df.iterrows():
        contract_code = str(row["合约代码"]).strip()
        position_type = str(row["持仓类别"]).strip() if pd.notna(row["持仓类别"]) else ""

        action = None
        for col in ["平33%", "平50%", "平100%", "反手"]:
            val = row[col]
            if pd.isna(val):
                continue
            s = str(val).strip().upper()
            if s in truthy:
                action = col
                break
            if str(val).strip() in ("是", "√", "✓"):
                action = col
                break

        if not action:
            print(f"[WARN] 第 {idx + 1} 行: 未指定操作,跳过")
            continue
        if not contract_code or contract_code == "nan":
            print(f"[WARN] 第 {idx + 1} 行: 合约代码为空,跳过")
            continue

        rows.append({
            "contract_code": contract_code,
            "position_type": position_type,
            "action": action,
        })
    return rows


# ---------- 倒计时 ----------
def countdown(s: int):
    print(f"将在 {s} 秒后开始,请把焦点切到钱龙软件...")
    for i in range(s, 0, -1):
        print(f"  {i}...", end="\r")
        time.sleep(1)
    print(" " * 30, end="\r")


# ---------- 单行任务:定位 + 操作 ----------
def locate_and_click_contract(win, rect, contract_code: str,
                               position_type: str) -> bool:
    table_cx = (rect.left + rect.right) // 2
    table_cy = (rect.top + rect.bottom) // 2

    move_mouse(table_cx, table_cy)
    click_screen(table_cx, table_cy)
    time.sleep(0.3)

    scroll_wheel(delta=120, times=SCROLL_UP_TIMES)
    time.sleep(0.3)

    for round_idx in range(1, MAX_ROUNDS + 1):
        print(f"\n  --- 第 {round_idx}/{MAX_ROUNDS} 轮翻找 ---")

        # 1) 截图
        img_rgb = grab(rect)
        Image.fromarray(img_rgb).save(ORIGINAL_IMAGE_PATH)

        # 2) 预处理
        processed = preprocess(img_rgb)
        Image.fromarray(processed).save(DEBUG_IMAGE_PATH)

        # 3) OCR (RapidOCR)
        print("  [..] 正在 RapidOCR 识别...")
        tokens = ocr_image(processed)
        print(f"  [OK] 识别出 {len(tokens)} 条 token")

        # 4) 找目标合约
        target, actual_pos, mode = find_target_in_columns(
            tokens, contract_code, position_type
        )
        if target is None:
            print("  [WARN] 表头未识别,降级到全局匹配")
            target, mode = find_contract_token(tokens, contract_code)
            actual_pos = None

        if target:
            skip_reason = None
            if position_type:
                if actual_pos is None:
                    skip_reason = f"持仓类别未识别,目标要求 {position_type!r}"
                elif actual_pos != position_type:
                    skip_reason = f"持仓类别={actual_pos!r} ≠ 目标 {position_type!r}"

            if skip_reason:
                print(f"  [跳过] {skip_reason},继续翻找...")
            else:
                cx, cy = box_center(target["box"])
                cx_orig = cx / ENLARGE_FACTOR
                cy_orig = cy / ENLARGE_FACTOR
                screen_x = int(rect.left + 2 + cx_orig)
                screen_y = int(rect.top + 2 + cy_orig)
                print(f"  [OK] 找到: {target['text']!r} (mode={mode}), "
                      f"持仓类别={actual_pos!r}")
                print(f"  [OK] 屏幕坐标: ({screen_x}, {screen_y})")
                click_screen(screen_x, screen_y)
                return True

        # 5) 检测终止关键词
        if has_keyword(tokens, STOP_KEYWORD):
            print(f"  [INFO] 检测到 {STOP_KEYWORD!r},已到表格底部")
            return False

        # 6) 下滚
        print(f"  [..] 第 {round_idx} 轮未找到,下滚...")
        move_mouse(table_cx, table_cy)
        click_screen(table_cx, table_cy)
        time.sleep(0.2)
        scroll_wheel(delta=-120, times=SCROLL_DOWN_TIMES)
        time.sleep(0.4)

    return False


def execute_action(win, action_name: str):
    auto_id = POSITION_BUTTONS.get(action_name)
    if not auto_id:
        print(f"  [错误] 未知的操作: {action_name}")
        return False

    print(f"\n  --- 执行操作: {action_name} ---")

    ok = click_position_button(win, action_name, auto_id)
    if not ok:
        print(f"  [跳过] {action_name} 未执行")
        return False
    time.sleep(0.5)

    print(f"  [..] 等待确认弹窗...")
    confirm_all_dialogs(main_win=win)
    return True


# ---------- 主流程 ----------
def main():
    print(f"[INFO] 读取 Excel: {EXCEL_PATH}")
    try:
        tasks = read_excel(EXCEL_PATH)
    except FileNotFoundError as e:
        print(f"[错误] {e}")
        print("请修改 EXCEL_PATH 指向正确的 Excel 文件")
        sys.exit(1)

    if not tasks:
        print("[错误] 没有可执行的任务")
        sys.exit(1)

    print(f"[OK] 共读取 {len(tasks)} 个任务")
    for i, t in enumerate(tasks, 1):
        print(f"  {i}. 合约={t['contract_code']}, "
              f"持仓={t['position_type']!r}, 操作={t['action']}")

    countdown(COUNTDOWN)

    print("\n[INFO] 正在连接钱龙窗口...")
    hwnd = find_window(WINDOW_KEY)
    print(f"[OK] 已找到窗口,句柄 = {hwnd}")
    win = activate_window(hwnd)
    switch_panel(win, TREE_ITEM)
    time.sleep(0.5)

    rect, _ = get_table(win, TABLE_AUTO_ID)

    for i, task in enumerate(tasks, 1):
        print(f"\n{'=' * 50}")
        print(f"[任务 {i}/{len(tasks)}] "
              f"合约={task['contract_code']}, "
              f"持仓={task['position_type']!r}, "
              f"操作={task['action']}")

        found = locate_and_click_contract(
            win, rect,
            task["contract_code"],
            task["position_type"],
        )

        if not found:
            print(f"[错误] 未找到合约 {task['contract_code']},跳过操作")
            continue

        time.sleep(0.5)
        execute_action(win, task["action"])
        print(f"[OK] 任务 {i} 完成")
        time.sleep(1)

    print(f"\n{'=' * 50}")
    print(f"[完成] 全部 {len(tasks)} 个任务执行完毕")


if __name__ == "__main__":
    main()