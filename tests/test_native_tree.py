import struct
import unittest
from unittest.mock import patch
import ctypes
from ctypes import wintypes

from core import native_tree


class PackTvItemTests(unittest.TestCase):
    def test_pack_32_bit_layout(self):
        data = native_tree._pack_tvitem(0x11223344, 0x55667788, 256, 32)
        self.assertEqual(len(data), 24)
        self.assertEqual(struct.unpack_from("<I", data, 4)[0], 0x11223344)
        self.assertEqual(struct.unpack_from("<I", data, 16)[0], 0x55667788)
        self.assertEqual(struct.unpack_from("<i", data, 20)[0], 256)

    def test_pack_64_bit_layout(self):
        data = native_tree._pack_tvitem(
            0x1122334455667788,
            0x1020304050607080,
            512,
            64,
        )
        self.assertEqual(len(data), 40)
        self.assertEqual(
            struct.unpack_from("<Q", data, 8)[0], 0x1122334455667788
        )
        self.assertEqual(
            struct.unpack_from("<Q", data, 24)[0], 0x1020304050607080
        )
        self.assertEqual(struct.unpack_from("<i", data, 32)[0], 512)

    def test_normalize_panel_path(self):
        self.assertEqual(
            native_tree.normalize_panel_path("/ 查询 /资金持仓/"),
            ["查询", "资金持仓"],
        )


class _FakeMemory:
    target_bits = 64

    def __init__(self, _):
        pass

    def __enter__(self):
        return self

    def __exit__(self, *_):
        pass


class _FakeReader:
    names = {1: "查询", 2: "撤单", 3: "资金持仓", 4: "策略持仓"}

    def __init__(self, *_):
        pass

    def get_text(self, item):
        return self.names[item]

    def close(self):
        pass


class SelectTreePathTests(unittest.TestCase):
    def _messages(self):
        state = {"selected": 0}

        def send(_hwnd, message, wparam=0, lparam=0, timeout_ms=3000):
            del timeout_ms
            if message == native_tree.TVM_GETNEXTITEM:
                if wparam == native_tree.TVGN_ROOT:
                    return 1
                if wparam == native_tree.TVGN_NEXT:
                    return {1: 2, 2: 0, 3: 4, 4: 0}.get(lparam, 0)
                if wparam == native_tree.TVGN_CHILD:
                    return {1: 3}.get(lparam, 0)
                if wparam == native_tree.TVGN_CARET:
                    return state["selected"]
            if message == native_tree.TVM_SELECTITEM:
                state["selected"] = lparam
                return 1
            return 1

        return state, send

    def test_selects_hierarchical_path(self):
        state, send = self._messages()
        with (
            patch.object(native_tree, "find_treeview", return_value=99),
            patch.object(native_tree, "get_tree_count", return_value=78),
            patch.object(native_tree, "RemoteProcessMemory", _FakeMemory),
            patch.object(native_tree, "TreeTextReader", _FakeReader),
            patch.object(native_tree, "_send_message", side_effect=send),
        ):
            result = native_tree.select_tree_path(10, r"\查询\策略持仓")

        self.assertEqual(state["selected"], 4)
        self.assertEqual(result["path"], ["查询", "策略持仓"])
        self.assertEqual(result["node_count"], 78)

    def test_missing_path_lists_siblings(self):
        _, send = self._messages()
        with (
            patch.object(native_tree, "find_treeview", return_value=99),
            patch.object(native_tree, "get_tree_count", return_value=78),
            patch.object(native_tree, "RemoteProcessMemory", _FakeMemory),
            patch.object(native_tree, "TreeTextReader", _FakeReader),
            patch.object(native_tree, "_send_message", side_effect=send),
        ):
            with self.assertRaises(native_tree.NativeTreePathError) as caught:
                native_tree.select_tree_path(10, r"\查询\不存在")

        self.assertIn("资金持仓", str(caught.exception))
        self.assertIn("策略持仓", str(caught.exception))


class NativeControlIntegrationTests(unittest.TestCase):
    """Exercise remote memory and TVM_* calls against a real local control."""

    def test_real_sys_treeview_round_trip(self):
        user32 = ctypes.WinDLL("user32", use_last_error=True)
        kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
        comctl32 = ctypes.WinDLL("comctl32", use_last_error=True)

        class InitCommonControlsEx(ctypes.Structure):
            _fields_ = [("dwSize", wintypes.DWORD), ("dwICC", wintypes.DWORD)]

        init = InitCommonControlsEx(ctypes.sizeof(InitCommonControlsEx), 0x0002)
        self.assertTrue(comctl32.InitCommonControlsEx(ctypes.byref(init)))

        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, ctypes.c_void_p, wintypes.HINSTANCE, ctypes.c_void_p,
        ]
        user32.CreateWindowExW.restype = wintypes.HWND
        user32.SendMessageW.argtypes = [
            wintypes.HWND, wintypes.UINT, ctypes.c_size_t, ctypes.c_ssize_t
        ]
        user32.SendMessageW.restype = ctypes.c_ssize_t

        instance = kernel32.GetModuleHandleW(None)
        parent = user32.CreateWindowExW(
            0, "STATIC", "NativeTreeTest", 0x00CF0000,
            0, 0, 400, 400, None, None, instance, None,
        )
        tree = user32.CreateWindowExW(
            0, "SysTreeView32", "", 0x50000000,
            0, 0, 300, 300, parent, ctypes.c_void_p(1223), instance, None,
        )
        self.assertTrue(parent)
        self.assertTrue(tree)

        class TvItemW(ctypes.Structure):
            _fields_ = [
                ("mask", wintypes.UINT), ("hItem", wintypes.HANDLE),
                ("state", wintypes.UINT), ("stateMask", wintypes.UINT),
                ("pszText", wintypes.LPWSTR), ("cchTextMax", ctypes.c_int),
                ("iImage", ctypes.c_int), ("iSelectedImage", ctypes.c_int),
                ("cChildren", ctypes.c_int), ("lParam", ctypes.c_ssize_t),
            ]

        class TvInsertStructW(ctypes.Structure):
            _fields_ = [
                ("hParent", wintypes.HANDLE),
                ("hInsertAfter", wintypes.HANDLE),
                ("item", TvItemW),
            ]

        def insert(parent_item, text):
            text_buffer = ctypes.create_unicode_buffer(text)
            item = TvItemW()
            item.mask = native_tree.TVIF_TEXT
            item.pszText = ctypes.cast(text_buffer, wintypes.LPWSTR)
            item.cchTextMax = len(text) + 1
            insertion = TvInsertStructW(
                parent_item, ctypes.c_void_p(-65534), item  # TVI_LAST
            )
            return user32.SendMessageW(
                tree, 0x1132, 0, ctypes.addressof(insertion)  # TVM_INSERTITEMW
            )

        try:
            query = insert(None, "查询")
            expected = insert(query, "资金详细信息")
            insert(None, "撤单")
            result = native_tree.select_tree_path(
                parent, r"\查询\资金详细信息"
            )
            self.assertEqual(result["selected_item"], expected)
            self.assertEqual(result["node_count"], 3)
            self.assertEqual(result["target_bits"], struct.calcsize("P") * 8)

            positional = native_tree.select_tree_path_by_position(
                parent,
                r"\查询\资金详细信息",
                {
                    "expected_node_count": 3,
                    "expected_root_child_counts": [1, 0],
                    "positions": {r"\查询\资金详细信息": [0, 0]},
                },
            )
            self.assertEqual(positional["selected_item"], expected)
            self.assertEqual(positional["position"], [0, 0])
            self.assertEqual(positional["root_child_counts"], [1, 0])

            with self.assertRaises(native_tree.NativeTreeError):
                native_tree.select_tree_path_by_position(
                    parent,
                    r"\查询\资金详细信息",
                    {
                        "expected_node_count": 3,
                        "expected_root_child_counts": [0, 1],
                        "positions": {r"\查询\资金详细信息": [0, 0]},
                    },
                )

            with self.assertRaises(native_tree.NativeTreeError):
                native_tree.select_tree_path_by_position(
                    parent,
                    r"\查询\资金详细信息",
                    {
                        "expected_node_count": 4,
                        "positions": {r"\查询\资金详细信息": [0, 0]},
                    },
                )
        finally:
            user32.DestroyWindow(parent)


if __name__ == "__main__":
    unittest.main()
