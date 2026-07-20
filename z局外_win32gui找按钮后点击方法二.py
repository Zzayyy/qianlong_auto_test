import win32gui
import win32con
from win32gui import (
    EnumChildWindows,
    GetWindowText,
    IsWindowVisible
)
from pywinauto import Application,findwindows
import time


# 获取窗口

elements = findwindows.find_elements(title_re=f".*{"钱龙模拟"}.*")
if not elements:
    None
# 默认取第一个匹配的句柄即可
print(elements[0].handle)
target_hwnd = elements[0].handle

app = Application(backend="uia").connect(handle=target_hwnd)
window = app.window(handle=target_hwnd)

window.set_focus()  # 激活窗口
buttons = []

def cb(hwnd, _):
    if GetWindowText(hwnd) == "输出":
        buttons.append(hwnd)

EnumChildWindows(target_hwnd, cb, None)

print(buttons)

for hwnd in buttons:
    print(hwnd, IsWindowVisible(hwnd))


btn = next(h for h in buttons if IsWindowVisible(h))


print(btn)
def click_hwnd(hwnd):
    import win32api
    import win32con
    from win32gui import GetWindowRect

    l, t, r, b = GetWindowRect(hwnd)

    x = (l + r) // 2
    y = (t + b) // 2

    win32api.SetCursorPos((x, y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0)
click_hwnd(btn)