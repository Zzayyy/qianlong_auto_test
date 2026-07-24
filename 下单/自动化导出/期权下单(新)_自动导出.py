# -*- coding: utf-8 -*-
import sys
import os
_here = os.path.dirname(os.path.abspath(__file__))
for _d in (_here, os.path.dirname(_here), os.path.dirname(os.path.dirname(_here))):
    if os.path.isdir(os.path.join(_d, "core")) and _d not in sys.path:
        sys.path.insert(0, _d)
        break
from core.window import find_window, activate_window, switch_panel as switch_main_panel

"""
钱龙期权交易 - 数据导出自动化(持仓 / 委托)
=====================================
功能流程:
    1. 查找"钱龙"主窗口
    2. 激活窗口并切换到"期权下单(新)"面板
    3. 点击"输出"按钮 (auto_id="1159"),弹出导出菜单
    4. 在弹出的菜单中选择"持仓"或"委托"
    5. 等待 Windows 系统的"另存为"窗口弹出
    6. 在文件名框内填入完整路径,点击"保存"
    7. 处理可能出现的"文件已存在"覆盖确认弹窗

运行环境:
    pip install pywinauto==0.6.8

使用方法:
    1. 打开钱龙旗舰版,登录交易账号,确保停留在"期权下单(新)"面板
    2. 修改下方 EXPORT_DIR / EXPORT_TARGETS / FILE_PREFIX
    3. 运行本脚本(将鼠标快速移到屏幕左上角可紧急停止)
    4. 倒计时内切换到钱龙窗口

注意:
    - "输出"按钮位于持仓/委托两个表格上方工具栏最右侧 (auto_id="1159")
    - 弹出的菜单是动态 MenuItem,静态组件树里抓不到,
      需要遍历顶级窗口实时匹配
    - 另存为窗口是 Windows 系统通用对话框,标题含"另存为"
"""

from pywinauto import Application, findwindows
import time
import sys
import os


# ====================== 可配置参数 ======================
TREE_ITEM       = "期权下单(新)"      # 左侧树节点名称
WINDOW_KEY      = "钱龙模拟期权宝"              # 窗口标题关键字
OUTPUT_AUTO_ID  = "1159"              # "输出"按钮 auto_id

# 从环境变量读取参数（GUI模式）
if os.environ.get("GUI_EXPORT_TARGETS"):
    # 从GUI传递的参数
    EXPORT_TARGETS = os.environ.get("GUI_EXPORT_TARGETS").split(",")
    EXPORT_DIR     = os.environ.get("GUI_EXPORT_DIR", r"D:\期权导出")
    FILE_EXT       = ".xls"  # 固定为.xls
    COUNTDOWN      = int(os.environ.get("GUI_COUNTDOWN", "3"))
else:
    # 默认参数（独立运行模式）
    # 要导出的表格子集;按顺序循环,例如 ["持仓"] / ["委托"] / ["持仓", "委托"]
    EXPORT_TARGETS  = [
        "持仓", 
        # "委托",
    ]

    # 另存为目录(不存在会自动创建)
    EXPORT_DIR      = r"D:\期权导出"

    # 扩展名,看另存为下拉框默认填什么(常见 .xlsx / .csv / .txt)
    FILE_EXT        = ".xls"

    COUNTDOWN       = 3      # 操作前倒计时秒数

WAIT_MENU_SEC   = 3      # 等"导出菜单"弹出的秒数
WAIT_SAVEAS_SEC = 3     # 等"另存为"窗口弹出的秒数
INTERVAL        = 1.0    # 两次导出之间的间隔

# 导出菜单项顺序(Win11 找不到 MenuItem 时,用于键盘方向键降级)
# 第一个是"持仓",第二个是"委托"
EXPORT_MENU_ORDER = ["持仓", "委托"]

# 键盘降级时每个目标需要按下的 Down 键次数(实测: 持仓按1次, 委托按2次)
# 说明: 菜单打开后首项并非直接高亮,需先下移相应次数再回车
EXPORT_MENU_DOWN_COUNT = {
    "持仓": 1,
    "委托": 2,
}
# ========================================================


# ====================== 窗口/面板基础操作 ======================



def switch_panel(win, tree_item: str):
    """点击左侧树节点,切换到指定面板。"""
    # tree = win.child_window(title=tree_item, control_type="TreeItem")
    # tree.wait("visible", timeout=10)
    # tree.click_input()
    tree = win.child_window(auto_id="1223", control_type="Tree")
    tree.wait("ready", timeout=10)
    tree.set_focus()
    tree.type_keys("{HOME}", with_spaces=False)
    time.sleep(0.2)
    print(f"[OK] 已切换到面板: {tree_item}")


# ====================== 导出主流程 ======================
def click_output_button(win, auto_id: str = OUTPUT_AUTO_ID):
    """点击右上角"输出"按钮,触发导出菜单(持仓 / 委托)。"""
    btn = win.child_window(title="输出", auto_id=auto_id, control_type="Button")
    btn.wait("ready", timeout=5)
    btn.click_input()
    print("[OK] 已点击'输出'按钮,等待导出菜单弹出...")


def select_export_target(target: str, win=None, timeout: float = WAIT_MENU_SEC) -> bool:
    """在弹出的导出菜单中选择目标项(持仓 / 委托)。

    该菜单是动态的,组件树里看不到,只能遍历顶级窗口实时匹配 MenuItem。
    弹出窗口可能是 Menu / Popup / 普通 Window,这里用宽松策略:
        - 遍历所有顶级窗口
        - 在每个窗口内查找 title==target 的 MenuItem
        - 找到就点击
        - 点击后检查是否有"提示"弹窗(表格无数据),如果有则关闭并返回失败

    Win11 兼容:
        部分 Win11 环境下,弹出的导出菜单无法被 UI Automation 识别为 MenuItem
        (find_elements 遍历不到 / child_window 匹配不到),此时降级为键盘方向键:
        菜单打开后第一个项默认高亮,按 Down 键按索引下移,再回车确认。
        菜单顺序见 EXPORT_MENU_ORDER(第一个持仓,第二个委托)。
    """
    # ---- 方案 A: UI Automation 查找 MenuItem 并点击 ----
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                item = dlg.child_window(
                    title=target,
                    control_type="MenuItem",
                )
                if item.exists(timeout=0.4):
                    item.wait("visible", timeout=1.5)
                    item.click_input()
                    print(f"[OK] 已选择导出项: {target} (hwnd={hwnd})")
                    return _after_select(target)
            except Exception:
                continue
        #time.sleep(0.15)

    # ---- 方案 B (Win11 降级): 键盘方向键选择 ----
    print(f"[WARN] 未找到 '{target}' 的 MenuItem(疑似 Win11 弹出菜单无法识别),"
          f"改用键盘方向键降级方案...")
    return _select_export_target_by_keyboard(target, win)


def _after_select(target: str) -> bool:
    """点击菜单项后的公共处理: 检测空数据提示。"""
    time.sleep(0.5)
    if _check_and_handle_empty_data_prompt():
        print(f"[WARN] {target} 表格无数据,已跳过")
        return False
    return True


def _select_export_target_by_keyboard(target: str, win=None) -> bool:
    """Win11 降级方案: 用键盘方向键在导出菜单中选择目标项。

    导出菜单打开后,第一项(持仓)默认处于高亮状态;按 Down 键按索引下移到位,
    再按 Enter 确认。菜单项顺序见 EXPORT_MENU_ORDER。
    """
    if win is None:
        print("[错误] 缺少主窗口句柄,无法使用键盘降级方案")
        return False
    try:
        # 把焦点给回主窗口(菜单已随"输出"点击而弹出并捕获键盘)
        win.set_focus()
        time.sleep(0.3)

        # 计算需要按下的 Down 次数(按目标显式配置)
        down_count = EXPORT_MENU_DOWN_COUNT.get(target, 0)
        keys = ""
        if down_count > 0:
            keys = f"{{DOWN {down_count}}}"
        win.type_keys(keys + "{ENTER}", with_spaces=False)
        print(f"[OK] 已用键盘选择导出项: {target} (下移 {down_count} 次)")
        return _after_select(target)
    except Exception as e:
        print(f"[WARN] 键盘降级方案失败: {e}")
        return False


def _check_and_handle_empty_data_prompt(timeout: float = 2) -> bool:
    """检查是否有"提示"弹窗(表格无数据),如果有则关闭它。
    
    Returns:
        bool: 是否有"提示"弹窗并已处理
    """
    end = time.time() + timeout
    while time.time() < end:
        for elem in findwindows.find_elements(top_level_only=True):
            hwnd = elem.handle
            try:
                dlg_app = Application(backend="uia").connect(handle=hwnd, timeout=0.5)
                dlg = dlg_app.window(handle=hwnd)
                title = dlg.window_text() or ""
                if "提示" in title:
                    print(f"[INFO] 检测到'提示'弹窗: {title}")
                    # 关闭弹窗(按回车或点击确定)
                    dlg.set_focus()
                    time.sleep(0.1)
                    dlg.type_keys("{ENTER}", with_spaces=False)
                    print(f"[OK] 已关闭'提示'弹窗")
                    return True
            except Exception:
                continue
        time.sleep(0.15)
    return False


def build_filename(target: str) -> str:
    """生成导出文件名: 期权下单(新)-持仓-20260629.xls"""
    date_str = time.strftime("%Y%m%d")
    return f"期权下单(新)-{target}-{date_str}{FILE_EXT}"


# ====================== 另存为对话框 ======================
# 直接复用 core.save_as_dialog：win32 主路径（方法二真实鼠标点击 + WM_SETTEXT），
# 失败自动回退 pywinauto。保持本脚本 handle_save_as_dialog 签名不变。
def handle_save_as_dialog(
    save_dir: str,
    filename: str,
    timeout: float = WAIT_SAVEAS_SEC,
) -> bool:
    """处理 Windows 系统的"另存为"窗口（委托 core，win32 优先）。"""
    from core.save_as_dialog import handle_save_as_dialog as _core_handle
    return _core_handle(save_dir=save_dir, filename=filename, timeout=timeout)


# ====================== 倒计时 ======================
def countdown(seconds: int):
    """操作前倒计时,让用户有时间切换窗口。"""
    print(f"将在 {seconds} 秒后开始,请把焦点切到钱龙软件...")
    try:
        for i in range(seconds, 0, -1):
            print(f"  {i}...", end="\r")
            time.sleep(1)
    except KeyboardInterrupt:
        raise
    print(" " * 30, end="\r")


# ====================== 入口 ======================
def main():
    try:
        if not EXPORT_TARGETS:
            print("[错误] EXPORT_TARGETS 不能为空")
            sys.exit(1)

        print(f"计划: 导出 {len(EXPORT_TARGETS)} 个表格 -> {EXPORT_DIR}")
        print(f"      目标: {', '.join(EXPORT_TARGETS)}")
        print(f"      扩展名: {FILE_EXT}")

        countdown(COUNTDOWN)

        hwnd = find_window(WINDOW_KEY)
        print(f"[OK] 已找到窗口,句柄 = {hwnd}")
        win = activate_window(hwnd)
        switch_main_panel(win, TREE_ITEM)
        time.sleep(0.5)  # 等待面板内容加载

        for i, target in enumerate(EXPORT_TARGETS, 1):
            print(f"\n--- [{i}/{len(EXPORT_TARGETS)}] 导出: {target} ---")
            click_output_button(win)

            if not select_export_target(target, win):
                print(f"[WARN] 跳过 {target},无法选择导出项")
                #time.sleep(INTERVAL)
                continue

            filename = build_filename(target)
            ok = handle_save_as_dialog(EXPORT_DIR, filename)
            if ok:
                print(f"[完成] {target} -> {os.path.join(EXPORT_DIR, filename)}")
            else:
                print(f"[失败] {target}")

            #time.sleep(INTERVAL)

        print(f"\n=== 全部完成 ===")

    except KeyboardInterrupt:
        print("\n[中断] 用户主动停止")
        sys.exit(0)
    except Exception as e:
        print(f"\n[错误] {type(e).__name__}: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
