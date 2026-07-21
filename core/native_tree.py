# -*- coding: utf-8 -*-
"""Native Win32 TreeView access used when UI Automation hides TreeItem nodes.

The Guotai Haitong client still owns a normal ``SysTreeView32`` on Windows 11,
but the OS UIA bridge may expose only its scroll bars.  This module talks to the
control with TVM_* messages and reads item text through a small buffer allocated
in the target process.
"""

from __future__ import annotations

import ctypes
from ctypes import wintypes
import struct
import time
import unicodedata

import win32gui
import win32process


TV_FIRST = 0x1100
TVM_EXPAND = TV_FIRST + 2
TVM_GETCOUNT = TV_FIRST + 5
TVM_GETNEXTITEM = TV_FIRST + 10
TVM_SELECTITEM = TV_FIRST + 11
TVM_GETITEMA = TV_FIRST + 12
TVM_GETITEMW = TV_FIRST + 62
TVM_ENSUREVISIBLE = TV_FIRST + 20

TVGN_ROOT = 0
TVGN_NEXT = 1
TVGN_CHILD = 4
TVGN_CARET = 9
TVE_EXPAND = 2

TVIF_TEXT = 0x0001
TVIF_HANDLE = 0x0010

PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

SMTO_BLOCK = 0x0001
SMTO_ABORTIFHUNG = 0x0002
ERROR_ACCESS_DENIED = 5


class NativeTreeError(RuntimeError):
    """Base error for native TreeView operations."""


class NativeTreeAccessError(NativeTreeError):
    """The automation process cannot read the target process."""


class NativeTreePathError(NativeTreeError):
    """A configured menu path does not exist in the target tree."""


_kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
_user32 = ctypes.WinDLL("user32", use_last_error=True)

_kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
_kernel32.OpenProcess.restype = wintypes.HANDLE
_kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
_kernel32.CloseHandle.restype = wintypes.BOOL
_kernel32.VirtualAllocEx.argtypes = [
    wintypes.HANDLE,
    ctypes.c_void_p,
    ctypes.c_size_t,
    wintypes.DWORD,
    wintypes.DWORD,
]
_kernel32.VirtualAllocEx.restype = ctypes.c_void_p
_kernel32.VirtualFreeEx.argtypes = [
    wintypes.HANDLE,
    ctypes.c_void_p,
    ctypes.c_size_t,
    wintypes.DWORD,
]
_kernel32.VirtualFreeEx.restype = wintypes.BOOL
_kernel32.WriteProcessMemory.argtypes = [
    wintypes.HANDLE,
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]
_kernel32.WriteProcessMemory.restype = wintypes.BOOL
_kernel32.ReadProcessMemory.argtypes = [
    wintypes.HANDLE,
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]
_kernel32.ReadProcessMemory.restype = wintypes.BOOL

_user32.SendMessageTimeoutW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    ctypes.c_size_t,
    ctypes.c_ssize_t,
    wintypes.UINT,
    wintypes.UINT,
    ctypes.POINTER(ctypes.c_size_t),
]
_user32.SendMessageTimeoutW.restype = ctypes.c_size_t


def _windows_error(prefix: str, error: int | None = None) -> OSError:
    code = ctypes.get_last_error() if error is None else error
    return OSError(code, f"{prefix}: {ctypes.FormatError(code).strip()}")


def _send_message(hwnd: int, message: int, wparam: int = 0, lparam: int = 0,
                  timeout_ms: int = 3000) -> int:
    """Send a control message without allowing a hung client to hang the tool."""
    result = ctypes.c_size_t()
    ctypes.set_last_error(0)
    ok = _user32.SendMessageTimeoutW(
        hwnd,
        message,
        wparam,
        lparam,
        SMTO_BLOCK | SMTO_ABORTIFHUNG,
        timeout_ms,
        ctypes.byref(result),
    )
    if not ok:
        raise NativeTreeError(str(_windows_error("TreeView 消息超时或失败")))
    return int(result.value)


def normalize_panel_path(panel_path: str) -> list[str]:
    path = panel_path.replace("/", "\\")
    parts = [part.strip() for part in path.split("\\") if part.strip()]
    if not parts:
        raise NativeTreePathError(f"无效的菜单路径: {panel_path!r}")
    return parts


def _normalized_text(value: str) -> str:
    return unicodedata.normalize("NFKC", value).replace("\u3000", " ").strip()


def find_treeview(parent_hwnd: int, control_id: int = 1223) -> int:
    """Find the requested descendant ``SysTreeView32`` by control ID."""
    matches: list[int] = []

    def _enum(hwnd, _):
        try:
            if (
                win32gui.GetDlgCtrlID(hwnd) == control_id
                and win32gui.GetClassName(hwnd) == "SysTreeView32"
            ):
                matches.append(hwnd)
        except win32gui.error:
            pass

    win32gui.EnumChildWindows(parent_hwnd, _enum, None)
    if not matches:
        raise NativeTreeError(
            f"未找到菜单树(control_id={control_id}, class=SysTreeView32)"
        )
    if len(matches) > 1:
        raise NativeTreeError(
            f"找到 {len(matches)} 个相同 ID 的菜单树，拒绝选择不确定目标"
        )
    return matches[0]


def get_tree_count(tree_hwnd: int) -> int:
    return _send_message(tree_hwnd, TVM_GETCOUNT)


def get_tree_root_child_counts(tree_hwnd: int) -> list[int]:
    """Return the number of direct children for every root item.

    This fingerprint uses only TreeView navigation messages and therefore does
    not need target-process memory access.  It is stronger than a total node
    count: two client versions can contain the same number of nodes but arrange
    them under different root menus.
    """
    counts: list[int] = []
    root = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    while root:
        child_count = 0
        child = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, root)
        while child:
            child_count += 1
            child = _send_message(
                tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, child
            )
        counts.append(child_count)
        root = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, root)
    return counts


def _get_process_bitness(process_handle: int) -> int:
    """Return 32 or 64 for an already opened process handle."""
    is_wow64_process2 = getattr(_kernel32, "IsWow64Process2", None)
    if is_wow64_process2 is not None:
        process_machine = wintypes.USHORT()
        native_machine = wintypes.USHORT()
        is_wow64_process2.argtypes = [
            wintypes.HANDLE,
            ctypes.POINTER(wintypes.USHORT),
            ctypes.POINTER(wintypes.USHORT),
        ]
        is_wow64_process2.restype = wintypes.BOOL
        if is_wow64_process2(
            process_handle,
            ctypes.byref(process_machine),
            ctypes.byref(native_machine),
        ):
            if process_machine.value != 0:
                return 32
            return 64 if native_machine.value in (0x8664, 0xAA64) else 32

    is_wow64 = wintypes.BOOL()
    _kernel32.IsWow64Process.argtypes = [
        wintypes.HANDLE,
        ctypes.POINTER(wintypes.BOOL),
    ]
    _kernel32.IsWow64Process.restype = wintypes.BOOL
    if not _kernel32.IsWow64Process(process_handle, ctypes.byref(is_wow64)):
        raise NativeTreeError(str(_windows_error("无法检测目标进程位数")))
    if struct.calcsize("P") == 4:
        return 32
    return 32 if is_wow64.value else 64


class RemoteProcessMemory:
    """Small, checked remote-memory buffer used by TreeView text messages."""

    def __init__(self, tree_hwnd: int):
        _, self.pid = win32process.GetWindowThreadProcessId(tree_hwnd)
        access = (
            PROCESS_QUERY_LIMITED_INFORMATION
            | PROCESS_VM_OPERATION
            | PROCESS_VM_READ
            | PROCESS_VM_WRITE
        )
        ctypes.set_last_error(0)
        self.handle = _kernel32.OpenProcess(access, False, self.pid)
        if not self.handle:
            error = ctypes.get_last_error()
            message = (
                f"目标进程 PID={self.pid} 禁止远程内存读取"
            )
            if error == ERROR_ACCESS_DENIED:
                raise NativeTreeAccessError(f"{message}（Windows 拒绝访问）")
            raise NativeTreeAccessError(f"{message}；错误码={error}")
        self.target_bits = _get_process_bitness(self.handle)
        self._allocations: set[int] = set()

    def allocate(self, size: int) -> int:
        ctypes.set_last_error(0)
        address = _kernel32.VirtualAllocEx(
            self.handle,
            None,
            size,
            MEM_COMMIT | MEM_RESERVE,
            PAGE_READWRITE,
        )
        if not address:
            raise NativeTreeAccessError(str(_windows_error("目标进程内存分配失败")))
        value = int(address)
        self._allocations.add(value)
        return value

    def free(self, address: int) -> None:
        if address in self._allocations:
            _kernel32.VirtualFreeEx(self.handle, address, 0, MEM_RELEASE)
            self._allocations.discard(address)

    def write(self, address: int, data: bytes) -> None:
        buffer = ctypes.create_string_buffer(data)
        written = ctypes.c_size_t()
        if not _kernel32.WriteProcessMemory(
            self.handle,
            address,
            buffer,
            len(data),
            ctypes.byref(written),
        ) or written.value != len(data):
            raise NativeTreeAccessError(str(_windows_error("写入目标进程内存失败")))

    def read(self, address: int, size: int) -> bytes:
        buffer = ctypes.create_string_buffer(size)
        read = ctypes.c_size_t()
        if not _kernel32.ReadProcessMemory(
            self.handle,
            address,
            buffer,
            size,
            ctypes.byref(read),
        ):
            raise NativeTreeAccessError(str(_windows_error("读取目标进程内存失败")))
        return buffer.raw[: read.value]

    def close(self) -> None:
        if getattr(self, "handle", None):
            for address in tuple(self._allocations):
                self.free(address)
            _kernel32.CloseHandle(self.handle)
            self.handle = None

    def __enter__(self):
        return self

    def __exit__(self, *_):
        self.close()


def _pack_tvitem(h_item: int, text_address: int, char_count: int,
                 target_bits: int) -> bytes:
    """Pack TVITEMW using the target process pointer layout."""
    buffer = bytearray(40 if target_bits == 64 else 24)
    struct.pack_into("<I", buffer, 0, TVIF_TEXT | TVIF_HANDLE)
    if target_bits == 32:
        struct.pack_into("<I", buffer, 4, h_item & 0xFFFFFFFF)
        struct.pack_into("<I", buffer, 16, text_address & 0xFFFFFFFF)
        struct.pack_into("<i", buffer, 20, char_count)
    elif target_bits == 64:
        struct.pack_into("<Q", buffer, 8, h_item)
        struct.pack_into("<Q", buffer, 24, text_address)
        struct.pack_into("<i", buffer, 32, char_count)
    else:
        raise ValueError(f"不支持的目标进程位数: {target_bits}")
    return bytes(buffer)


class TreeTextReader:
    STRUCT_SIZE = 64
    TEXT_SIZE = 1024

    def __init__(self, tree_hwnd: int, memory: RemoteProcessMemory):
        self.tree_hwnd = tree_hwnd
        self.memory = memory
        self.base = memory.allocate(self.STRUCT_SIZE + self.TEXT_SIZE)
        self.text_address = self.base + self.STRUCT_SIZE

    def get_text(self, h_item: int) -> str:
        package = bytearray(self.STRUCT_SIZE + self.TEXT_SIZE)
        item = _pack_tvitem(
            h_item,
            self.text_address,
            self.TEXT_SIZE // 2,
            self.memory.target_bits,
        )
        package[: len(item)] = item
        self.memory.write(self.base, bytes(package))
        if _send_message(self.tree_hwnd, TVM_GETITEMW, 0, self.base):
            raw = self.memory.read(self.text_address, self.TEXT_SIZE)
            return raw.decode("utf-16-le", errors="replace").split("\0", 1)[0]

        item = _pack_tvitem(
            h_item,
            self.text_address,
            self.TEXT_SIZE,
            self.memory.target_bits,
        )
        package[:] = b"\0" * len(package)
        package[: len(item)] = item
        self.memory.write(self.base, bytes(package))
        if _send_message(self.tree_hwnd, TVM_GETITEMA, 0, self.base):
            raw = self.memory.read(self.text_address, self.TEXT_SIZE)
            return raw.split(b"\0", 1)[0].decode("gb18030", errors="replace")
        raise NativeTreeError(f"无法读取 TreeView 节点文本: hItem={h_item}")

    def close(self) -> None:
        self.memory.free(self.base)


def select_tree_path(parent_hwnd: int, panel_path: str,
                     control_id: int = 1223) -> dict:
    """Select a menu path and return diagnostic information."""
    parts = normalize_panel_path(panel_path)
    tree_hwnd = find_treeview(parent_hwnd, control_id)
    count = get_tree_count(tree_hwnd)
    if count <= 0:
        raise NativeTreeError(f"菜单树存在但节点数为 {count}")

    with RemoteProcessMemory(tree_hwnd) as memory:
        reader = TreeTextReader(tree_hwnd, memory)
        try:
            h_item = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
            for depth, target in enumerate(parts):
                expected = _normalized_text(target)
                current = h_item
                sibling_names: list[str] = []
                matched = 0
                while current:
                    text = reader.get_text(current)
                    sibling_names.append(text)
                    if _normalized_text(text) == expected:
                        matched = current
                        break
                    current = _send_message(
                        tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, current
                    )
                if not matched:
                    visible = "、".join(name for name in sibling_names if name)
                    raise NativeTreePathError(
                        f"第 {depth + 1} 层未找到菜单 {target!r}；"
                        f"同层节点: {visible or '(无可读文本)'}"
                    )

                h_item = matched
                if depth < len(parts) - 1:
                    _send_message(tree_hwnd, TVM_EXPAND, TVE_EXPAND, h_item)
                    time.sleep(0.1)
                    h_item = _send_message(
                        tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, h_item
                    )
                    if not h_item:
                        raise NativeTreePathError(f"菜单 {target!r} 没有子节点")

            _send_message(tree_hwnd, TVM_ENSUREVISIBLE, 0, h_item)
            _send_message(tree_hwnd, TVM_SELECTITEM, TVGN_CARET, h_item)
            selected = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_CARET, 0)
            if selected != h_item:
                raise NativeTreeError("TreeView 未确认目标节点为当前选中项")
            return {
                "tree_hwnd": tree_hwnd,
                "node_count": count,
                "target_bits": memory.target_bits,
                "selected_item": h_item,
                "path": parts,
            }
        finally:
            reader.close()


def select_tree_path_by_position(parent_hwnd: int, panel_path: str,
                                 profile: dict, control_id: int = 1223) -> dict:
    """Select a path from a verified positional profile without process memory.

    This fallback is intentionally fail-closed: it is used only when the live
    TreeView node count exactly matches the profile fingerprint.  It is useful
    for clients that block ``OpenProcess`` and whose Windows 11 accessibility
    provider does not expose item names.
    """
    parts = normalize_panel_path(panel_path)
    normalized_path = "\\" + "\\".join(parts)
    positions = profile.get("positions") or {}
    indexes = positions.get(normalized_path)
    if not isinstance(indexes, list) or not indexes:
        raise NativeTreePathError(
            f"当前客户端没有菜单 {normalized_path!r} 的安全位置配置"
        )
    if len(indexes) != len(parts) or any(
        not isinstance(index, int) or index < 0 for index in indexes
    ):
        raise NativeTreePathError(
            f"菜单 {normalized_path!r} 的位置配置无效: {indexes!r}"
        )

    expected_count = profile.get("expected_node_count")
    if not isinstance(expected_count, int) or expected_count <= 0:
        raise NativeTreeError("位置配置缺少 expected_node_count 安全指纹")

    tree_hwnd = find_treeview(parent_hwnd, control_id)
    actual_count = get_tree_count(tree_hwnd)
    if actual_count != expected_count:
        raise NativeTreeError(
            f"菜单树节点数已变化（期望 {expected_count}，实际 {actual_count}），"
            "为避免误点，已停用位置定位"
        )

    expected_root_child_counts = profile.get("expected_root_child_counts")
    if expected_root_child_counts is not None:
        if not isinstance(expected_root_child_counts, list) or any(
            not isinstance(value, int) or value < 0
            for value in expected_root_child_counts
        ):
            raise NativeTreeError(
                "位置配置的 expected_root_child_counts 无效"
            )
        actual_root_child_counts = get_tree_root_child_counts(tree_hwnd)
        if actual_root_child_counts != expected_root_child_counts:
            raise NativeTreeError(
                "菜单树拓扑已变化（根节点直属子项数量不一致）："
                f"期望 {expected_root_child_counts}，"
                f"实际 {actual_root_child_counts}；"
                "为避免误点，已停用位置定位"
            )
    else:
        actual_root_child_counts = None

    h_item = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    for depth, sibling_index in enumerate(indexes):
        if not h_item:
            raise NativeTreePathError(
                f"菜单 {normalized_path!r} 在第 {depth + 1} 层没有节点"
            )
        for _ in range(sibling_index):
            h_item = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, h_item)
            if not h_item:
                raise NativeTreePathError(
                    f"菜单 {normalized_path!r} 在第 {depth + 1} 层越界"
                )
        if depth < len(indexes) - 1:
            _send_message(tree_hwnd, TVM_EXPAND, TVE_EXPAND, h_item)
            time.sleep(0.1)
            h_item = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, h_item)

    _send_message(tree_hwnd, TVM_ENSUREVISIBLE, 0, h_item)
    _send_message(tree_hwnd, TVM_SELECTITEM, TVGN_CARET, h_item)
    selected = _send_message(tree_hwnd, TVM_GETNEXTITEM, TVGN_CARET, 0)
    if selected != h_item:
        raise NativeTreeError("TreeView 未确认位置目标为当前选中项")
    return {
        "tree_hwnd": tree_hwnd,
        "node_count": actual_count,
        "selected_item": h_item,
        "path": parts,
        "position": indexes,
        "root_child_counts": actual_root_child_counts,
    }
