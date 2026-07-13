import win32gui
import win32process
import ctypes
from ctypes import wintypes
import struct
import time
from pywinauto import findwindows

# ============================================================
# TreeView 消息常量
# ============================================================
TVM_GETNEXTITEM   = 0x110A
TVM_GETITEMA      = 0x110C
TVM_GETITEMW      = 0x113E
TVM_SELECTITEM    = 0x110B
TVM_EXPAND        = 0x1102
TVM_ENSUREVISIBLE = 0x1114
TVM_GETCOUNT      = 0x1105
TVM_GETITEMRECT   = 0x1104

TVGN_ROOT  = 0
TVGN_NEXT  = 1
TVGN_CHILD = 4
TVGN_CARET = 9

TVE_EXPAND = 2

TVIF_TEXT   = 0x0001
TVIF_HANDLE = 0x0010


# ============================================================
# 检测目标进程是 32 位还是 64 位
# ============================================================
def detect_target_bitness(hwnd):
    """检测 hwnd 所在进程的位数"""
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    h_proc = ctypes.windll.kernel32.OpenProcess(
        0x1000,  # PROCESS_QUERY_LIMITED_INFORMATION
        False, pid
    )
    if not h_proc:
        return struct.calcsize("P") * 8  # 无法检测，假设和 Python 一样

    is_wow64 = ctypes.c_int(0)
    ctypes.windll.kernel32.IsWow64Process(h_proc, ctypes.byref(is_wow64))
    ctypes.windll.kernel32.CloseHandle(h_proc)

    python_bits = struct.calcsize("P") * 8
    if python_bits == 64:
        # 64 位 Python: WOW64=True → 目标是 32 位; False → 64 位
        return 32 if is_wow64.value else 64
    else:
        return 32  # 32 位 Python 上所有进程都是 32 位


# ============================================================
# 远程内存管理
# ============================================================
class RemoteMem:
    """在目标进程中分配/读写内存"""

    def __init__(self, tree_hwnd, target_bits):
        _, self.pid = win32process.GetWindowThreadProcessId(tree_hwnd)
        self.h_proc = ctypes.windll.kernel32.OpenProcess(
            0x0008 | 0x0010 | 0x0020,  # VM_OPERATION | VM_READ | VM_WRITE
            False, self.pid
        )
        if not self.h_proc:
            raise OSError(f"无法打开进程 PID={self.pid}，请以管理员权限运行脚本")
        self.target_bits = target_bits

    def alloc(self, size):
        addr = ctypes.windll.kernel32.VirtualAllocEx(
            self.h_proc, 0, size, 0x1000, 0x04  # MEM_COMMIT, PAGE_READWRITE
        )
        if not addr:
            raise OSError("远程内存分配失败")
        return addr

    def free(self, addr):
        ctypes.windll.kernel32.VirtualFreeEx(self.h_proc, addr, 0, 0x8000)

    def write(self, addr, data):
        n = ctypes.c_size_t(0)
        ctypes.windll.kernel32.WriteProcessMemory(
            self.h_proc, addr, data, len(data), ctypes.byref(n)
        )

    def read_text(self, addr, max_bytes=512, unicode_mode=True):
        buf = ctypes.create_string_buffer(max_bytes)
        n = ctypes.c_size_t(0)
        ctypes.windll.kernel32.ReadProcessMemory(
            self.h_proc, addr, buf, max_bytes, ctypes.byref(n)
        )
        raw = buf.raw
        if unicode_mode:
            return raw.decode('utf-16-le', errors='replace').rstrip('\x00')
        else:
            # 中文 ANSI 程序通常使用 GBK 编码
            try:
                return raw.split(b'\x00')[0].decode('gbk')
            except:
                return raw.split(b'\x00')[0].decode('latin-1')

    def close(self):
        if self.h_proc:
            ctypes.windll.kernel32.CloseHandle(self.h_proc)
            self.h_proc = None

    def __del__(self):
        self.close()


# ============================================================
# 核心：按目标进程位数构造 TVITEM 结构体
# ============================================================
def pack_tvitem(h_item, text_buf_addr, cch_max, target_bits):
    """
    按目标进程位数构造 TVITEM(A/W 通用) 字节序列

    32 位布局:                        64 位布局:
    +0  mask       (4)               +0  mask       (4)
    +4  hItem      (4)               +4  (pad)      (4)
    +8  state      (4)               +8  hItem      (8)
    +12 stateMask  (4)               +16 state      (4)
    +16 pszText    (4)               +20 stateMask  (4)
    +20 cchTextMax (4)               +24 pszText    (8)
    总计 24 字节                      +32 cchTextMax (4)
                                      总计 40 字节
    """
    buf = bytearray(64)  # 预留足够空间

    if target_bits == 32:
        struct.pack_into("=I", buf, 0,  TVIF_TEXT | TVIF_HANDLE)
        struct.pack_into("=I", buf, 4,  h_item & 0xFFFFFFFF)
        struct.pack_into("=I", buf, 8,  0)  # state
        struct.pack_into("=I", buf, 12, 0)  # stateMask
        struct.pack_into("=I", buf, 16, text_buf_addr & 0xFFFFFFFF)
        struct.pack_into("=i", buf, 20, cch_max)
    else:
        struct.pack_into("=I", buf, 0,  TVIF_TEXT | TVIF_HANDLE)
        # +4 自然是 0（bytearray 初始化为零）= padding
        struct.pack_into("=Q", buf, 8,  h_item)
        struct.pack_into("=I", buf, 16, 0)  # state
        struct.pack_into("=I", buf, 20, 0)  # stateMask
        struct.pack_into("=Q", buf, 24, text_buf_addr)
        struct.pack_into("=i", buf, 32, cch_max)

    return bytes(buf)


# ============================================================
# 获取节点文本（自动尝试 Unicode 和 ANSI）
# ============================================================
def get_node_text(tree_hwnd, h_item, mem):
    """
    获取 TreeView 节点文本
    先尝试 TVM_GETITEMW (Unicode), 失败再尝试 TVM_GETITEMA (ANSI)
    """
    TEXT_BYTES = 512
    PAD = 64  # TVITEM 结构体空间
    total = PAD + TEXT_BYTES

    remote_base = mem.alloc(total)
    text_addr = remote_base + PAD

    try:
        # ── 尝试 1: Unicode (TVM_GETITEMW) ──
        cch_max = TEXT_BYTES // 2  # 字符数
        tvitem_data = pack_tvitem(h_item, text_addr, cch_max, mem.target_bits)
        mem.write(remote_base, tvitem_data)

        result = win32gui.SendMessage(tree_hwnd, TVM_GETITEMW, 0, remote_base)
        if result:
            return mem.read_text(text_addr, TEXT_BYTES, unicode_mode=True)

        # ── 尝试 2: ANSI (TVM_GETITEMA) ──
        cch_max = TEXT_BYTES  # 字节数
        tvitem_data = pack_tvitem(h_item, text_addr, cch_max, mem.target_bits)
        mem.write(remote_base, tvitem_data)

        result = win32gui.SendMessage(tree_hwnd, TVM_GETITEMA, 0, remote_base)
        if result:
            return mem.read_text(text_addr, TEXT_BYTES, unicode_mode=False)

        return None

    finally:
        mem.free(remote_base)


# ============================================================
# 查找 TreeView 句柄
# ============================================================
def find_treeview(parent_hwnd, control_id=1223):
    """通过控件 ID 查找 TreeView"""
    # 方法1: GetDlgItem（最快）
    hwnd = ctypes.windll.user32.GetDlgItem(parent_hwnd, control_id)
    if hwnd:
        cls = win32gui.GetClassName(hwnd)
        print(f"✓ TreeView 句柄: {hex(hwnd)} (ID={control_id}, 类名={cls})")
        return hwnd

    # 方法2: 遍历
    found = []
    def _enum(h, _):
        if win32gui.GetClassName(h) == "SysTreeView32":
            found.append(h)
        return True
    win32gui.EnumChildWindows(parent_hwnd, _enum, None)

    if found:
        for h in found:
            cid = ctypes.windll.user32.GetDlgCtrlID(h)
            print(f"✓ TreeView 句柄: {hex(h)} (ID={cid})")
        return found[0]

    print("✗ 未找到 TreeView")
    return 0


# ============================================================
# 遍历整棵树
# ============================================================
def dump_tree(tree_hwnd, mem):
    """遍历并打印树结构"""
    count = win32gui.SendMessage(tree_hwnd, TVM_GETCOUNT, 0, 0)
    print(f"节点总数: {count}\n")

    def _walk(h_item, level):
        while h_item:
            text = get_node_text(tree_hwnd, h_item, mem)
            name = text.strip() if text else "(无文本)"
            indent = "  " * level
            print(f"{indent}{name}")

            h_child = win32gui.SendMessage(
                tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, h_item
            )
            if h_child:
                _walk(h_child, level + 1)

            h_item = win32gui.SendMessage(
                tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, h_item
            )

    h_root = win32gui.SendMessage(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    if h_root:
        _walk(h_root, 0)
    else:
        print("(树为空)")


# ============================================================
# 按路径选择节点
# ============================================================
def select_node(tree_hwnd, path, mem):
    """
    按路径选择节点并高亮
    path: ["单向下单", "买入平仓"]
    返回: 节点句柄 或 0
    """
    h_item = win32gui.SendMessage(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)

    for depth, target in enumerate(path):
        current = h_item
        found = False

        while current:
            text = get_node_text(tree_hwnd, current, mem)
            if text and text.strip() == target.strip():
                found = True
                h_item = current
                break
            current = win32gui.SendMessage(
                tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, current
            )

        if not found:
            print(f"✗ 第{depth}层未找到: '{target}'")
            return 0

        # 展开并进入子层
        if depth < len(path) - 1:
            win32gui.SendMessage(tree_hwnd, TVM_EXPAND, TVE_EXPAND, h_item)
            time.sleep(0.15)
            h_item = win32gui.SendMessage(
                tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, h_item
            )
            if not h_item:
                print(f"✗ '{target}' 没有子节点")
                return 0

    # 确保可见 + 选中
    win32gui.SendMessage(tree_hwnd, TVM_ENSUREVISIBLE, 0, h_item)
    time.sleep(0.05)
    win32gui.SendMessage(tree_hwnd, TVM_SELECTITEM, TVGN_CARET, h_item)

    text = get_node_text(tree_hwnd, h_item, mem)
    print(f"✓ 已选中: {' > '.join(path)}  [{text}]")
    return h_item


# ============================================================
# 选中 + 双击激活（如果仅选中不触发业务，用这个）
# ============================================================
def select_and_click(tree_hwnd, path, mem):
    """选中节点后模拟双击，确保触发业务逻辑"""
    h_item = select_node(tree_hwnd, path, mem)
    if not h_item:
        return False

    time.sleep(0.2)

    # 获取节点矩形 (TVM_GETITEMRECT)
    # lParam 指向一个 RECT，但输入时 RECT.left = hItem
    remote_size = 64
    remote_buf = mem.alloc(remote_size)

    # 写入: 前 4/8 字节 = hItem（作为输入），后面是 RECT 输出空间
    if mem.target_bits == 32:
        data = struct.pack("=I", h_item & 0xFFFFFFFF) + b'\x00' * 16
    else:
        data = struct.pack("=Q", h_item) + b'\x00' * 32

    mem.write(remote_buf, data)

    # TVM_GETITEMRECT, wParam=True (只返回文本区域)
    result = win32gui.SendMessage(tree_hwnd, TVM_GETITEMRECT, True, remote_buf)

    if result:
        rect_data = mem.read_text(remote_buf, 32, unicode_mode=False)
        # 直接读原始字节来解析 RECT
        buf = ctypes.create_string_buffer(32)
        n = ctypes.c_size_t(0)
        ctypes.windll.kernel32.ReadProcessMemory(
            mem.h_proc, remote_buf, buf, 32, ctypes.byref(n)
        )

        if mem.target_bits == 32:
            left, top, right, bottom = struct.unpack("=iiii", buf.raw[:16])
        else:
            left, top, right, bottom = struct.unpack("=iiii", buf.raw[:16])

        cx = (left + right) // 2
        cy = (top + bottom) // 2
        print(f"  节点位置: ({left},{top})-({right},{bottom}), 中心=({cx},{cy})")

        mem.free(remote_buf)

        # 转为屏幕坐标并双击
        pt = win32gui.ClientToScreen(tree_hwnd, (cx, cy))

        # 激活窗口
        parent = win32gui.GetParent(tree_hwnd) or tree_hwnd
        win32gui.SetForegroundWindow(parent)
        time.sleep(0.1)

        # 移动鼠标到节点位置
        ctypes.windll.user32.SetCursorPos(pt[0], pt[1])
        time.sleep(0.05)

        # 双击
        for _ in range(2):
            ctypes.windll.user32.mouse_event(0x0002, 0, 0, 0, 0)  # LEFTDOWN
            ctypes.windll.user32.mouse_event(0x0004, 0, 0, 0, 0)  # LEFTUP
            time.sleep(0.05)

        print(f"  ✓ 已双击")
        return True
    else:
        mem.free(remote_buf)
        print(f"  ✗ 获取节点矩形失败，仅选中")
        return False


# ============================================================
# 主程序
# ============================================================
if __name__ == "__main__":
    # ── 1. 找到目标窗口 ──
    elements = findwindows.find_elements(title_re=".*钱龙模拟.*")
    if not elements:
        print("未找到钱龙模拟窗口，请确认程序已打开")
        exit(1)

    target_hwnd = elements[0].handle
    print(f"主窗口句柄: {hex(target_hwnd)}")

    # ── 2. 检测目标进程位数（关键！）──
    target_bits = detect_target_bitness(target_hwnd)
    python_bits = struct.calcsize("P") * 8
    print(f"Python: {python_bits} 位  |  目标进程: {target_bits} 位")

    if python_bits != target_bits:
        print(f"⚠  位数不同！结构体将按 {target_bits} 位布局构造")

    # ── 3. 找 TreeView 控件 ──
    win32gui.SetForegroundWindow(target_hwnd)
    time.sleep(0.3)

    tree_hwnd = find_treeview(target_hwnd, control_id=1223)
    if not tree_hwnd:
        exit(1)

    # ── 4. 创建远程内存管理器 ──
    mem = RemoteMem(tree_hwnd, target_bits)

    # ── 5. 遍历整棵树 ──
    print("\n" + "=" * 50)
    print("遍历树结构:")
    print("=" * 50)
    dump_tree(tree_hwnd, mem)

    # ── 6. 按路径选择 ──
    print("\n" + "=" * 50)
    print("选择节点:")
    print("=" * 50)

    # select_node(tree_hwnd, ["买入平仓"], mem)
    select_node(tree_hwnd, ["查询", "资金查询"], mem)
    # select_node(tree_hwnd, ["组合申报", "拆分申报"], mem)

    # ── 7. 如果选中不触发业务，改用双击 ──
    # print("\n选中 + 双击激活:")
    # select_and_click(tree_hwnd, ["单向下单", "买入平仓"], mem)

    # ── 清理 ──
    mem.close()
    print("\n完成")