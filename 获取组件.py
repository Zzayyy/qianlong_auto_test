from pywinauto import Application,findwindows
import time


# 获取窗口

elements = findwindows.find_elements(title_re=f".*{"国泰海通"}.*")
if not elements:
    None
# 默认取第一个匹配的句柄即可
print(elements[0].handle)
target_hwnd = elements[0].handle

app = Application(backend="uia").connect(handle=target_hwnd)
window = app.window(handle=target_hwnd)

window.set_focus()  # 激活窗口


# lock = window.child_window(title="解锁", auto_id="17004", control_type="Button")

# lock.click()

window.print_control_identifiers(depth=None,filename='国泰海通1.txt')  # 获取所有组件
# 尝试用gbk读取，用utf-8保存
with open('国泰海通1.txt', 'r', encoding='gbk', errors='ignore') as f:
    content = f.read()
with open('国泰海通1.txt', 'w', encoding='utf-8') as f:
    f.write(content)
