from pywinauto import Application,findwindows
from pywinauto import mouse
import time
import ctypes

# 获取屏幕分辨率
def get_screen_resolution():
    """获取屏幕分辨率"""
    user32 = ctypes.windll.user32
    width = user32.GetSystemMetrics(0)  # SM_CXSCREEN = 0
    height = user32.GetSystemMetrics(1)  # SM_CYSCREEN = 1
    return width, height

# 获取当前鼠标位置（用于调试）
def get_current_mouse_position():
    """获取当前鼠标位置"""
    point = ctypes.wintypes.POINT()
    ctypes.windll.user32.GetCursorPos(ctypes.byref(point))
    return point.x, point.y

# Windows API 用于更精确的鼠标控制
def click_at_position(x, y):
    """使用Windows API在指定位置点击"""
    # 移动鼠标
    ctypes.windll.user32.SetCursorPos(x, y)
    time.sleep(0.1)
    
    # 鼠标左键按下
    ctypes.windll.user32.mouse_event(2, 0, 0, 0, 0)  # LEFT_DOWN
    time.sleep(0.05)
    
    # 鼠标左键释放
    ctypes.windll.user32.mouse_event(4, 0, 0, 0, 0)  # LEFT_UP
    time.sleep(0.05)

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

# 新增功能：定位菜单并操作
def locate_and_click_menu():
    """定位窗口左侧菜单区域并双击，然后操作tree控件"""
    # 获取窗口矩形坐标
    rect = window.rectangle()
    print(f"窗口坐标: 左={rect.left}, 上={rect.top}, 右={rect.right}, 下={rect.bottom}")
    
    # 获取屏幕分辨率
    screen_width, screen_height = get_screen_resolution()
    print(f"屏幕分辨率: {screen_width} x {screen_height}")
    
    # 根据屏幕分辨率动态计算偏移量
    # 使用屏幕宽度的1%作为偏移量，确保不同分辨率下都能明显移动
    offset_x = int(screen_width * 0.01)  # 屏幕宽度的1%
    if offset_x < 50:  # 最小50像素
        offset_x = 50
    if offset_x > 200:  # 最大200像素
        offset_x = 200
    
    print(f"动态计算的水平偏移量: {offset_x} 像素")
    
    # 计算菜单位置：窗口最左边向右移动offset_x像素，垂直居中
    menu_x = rect.left + offset_x
    menu_y = rect.top + (rect.bottom - rect.top) // 2  # 垂直居中
    
    print(f"菜单定位坐标: x={menu_x}, y={menu_y}")
    
    # 获取当前鼠标位置（用于对比）
    old_x, old_y = get_current_mouse_position()
    print(f"移动前鼠标位置: x={old_x}, y={old_y}")
    
    # 移动鼠标到菜单位置
    print(f"移动鼠标从 ({old_x}, {old_y}) 到 ({menu_x}, {menu_y})...")
    mouse.move(coords=(menu_x, menu_y))
    time.sleep(0.3)
    
    # 验证鼠标是否移动成功
    new_x, new_y = get_current_mouse_position()
    print(f"移动后鼠标位置: x={new_x}, y={new_y}")
    
    # 双击菜单区域
    mouse.double_click(coords=(menu_x, menu_y))
    print("已双击菜单区域")
    time.sleep(0.5)
    
    # 通过键盘发送HOME键
    time.sleep(0.3)  # 等待菜单展开
    window.type_keys("{HOME}", with_spaces=False)
    print("已通过键盘发送HOME键")

# 执行菜单定位功能
locate_and_click_menu()
