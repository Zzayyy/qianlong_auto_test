from pywinauto import Application,findwindows
import time

# 客户端名称（窗口标题关键字 / 输出文件名前缀）
CLIENT_NAME = "华宝证券"


# 获取窗口

elements = findwindows.find_elements(title_re=f".*{CLIENT_NAME}.*")
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

window.print_control_identifiers(depth=None,filename=f'{CLIENT_NAME}.txt')  # 获取所有组件
# 尝试用gbk读取，用utf-8保存
with open(f'{CLIENT_NAME}.txt', 'r', encoding='gbk', errors='ignore') as f:
    content = f.read()
with open(f'{CLIENT_NAME}.txt', 'w', encoding='utf-8') as f:
    f.write(content)
