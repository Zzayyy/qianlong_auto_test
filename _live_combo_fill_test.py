# -*- coding: utf-8 -*-
"""组合申报安全批量填表验证（不提交、不报单）。

演示新复选框模式：市场(上证) × 全部 6 个策略，逐个打开组合申报对话框、
按持仓派生并选合约一/合约二、填组合数量、回读验证、关闭（不点击“申报”）。
绝不下单。运行前需 国泰海通证券期权宝 已登录。
"""

import os
import sys
import ctypes
import ctypes.wintypes as wt

sys.path.insert(0, r"E:\Code\6.GUI模块化")

os.environ["GUI_CLIENT_ID"] = "guotai_haitong"
os.environ["GUI_COUNTDOWN"] = "1"

from core.window import find_window
from core.clients import get_client
from core import super_strategy as ss
from core import combination_order as cm

user32 = ctypes.windll.user32
user32.SendMessageW.argtypes = [wt.HWND, wt.UINT, wt.WPARAM, wt.LPARAM]
user32.SendMessageW.restype = ctypes.c_ssize_t
CB_GETCURSEL = 0x0147
CB_GETLBTEXT = 0x0148
CB_GETLBTEXTLEN = 0x0149


def combo_cur_text(hwnd):
    idx = user32.SendMessageW(hwnd, CB_GETCURSEL, 0, 0)
    if idx < 0:
        return cm.get_edit_text(hwnd)
    ln = user32.SendMessageW(hwnd, CB_GETLBTEXTLEN, idx, 0)
    buf = ctypes.create_unicode_buffer(ln + 1)
    user32.SendMessageW(hwnd, CB_GETLBTEXT, idx, ctypes.cast(ctypes.addressof(buf), wt.LPARAM))
    return buf.value


def main():
    markets = ["上证"]
    strategies = list(ss.COMBO_STRATEGIES)  # 全部 6 个策略
    print(f"批量验证：市场={markets} × 策略={len(strategies)} 个 = {len(markets)*len(strategies)} 个组合（均不提交）")
    results = []
    try:
        results = ss.run_combination_declare_all(
            markets=markets,
            strategies=strategies,
            qty=1,
            execute=False,          # 安全：不点击申报
            cleanup_leftover=True,
        )
        for r in results:
            dlg = r.get("dialog_hwnd")
            if not dlg:
                print("  - 跳过（未打开对话框）:", r)
                continue
            c_market = combo_cur_text(cm.find_visible_child(dlg, ss.COMBO_DLG_MARKET_CID))
            c_strat = combo_cur_text(cm.find_visible_child(dlg, ss.COMBO_DLG_STRATEGY_CID))
            c_c1 = combo_cur_text(cm.find_visible_child(dlg, ss.COMBO_DLG_CONTRACT1_CID))
            c_c2 = combo_cur_text(cm.find_visible_child(dlg, ss.COMBO_DLG_CONTRACT2_CID))
            print(f"  [{r.get('market')}/{r.get('strategy')}] "
                  f"回填 => 合约一={c_c1!r} 合约二={c_c2!r} (市场={c_market!r} 策略={c_strat!r})")
    finally:
        try:
            client = get_client(os.environ["GUI_CLIENT_ID"])
            mw = find_window(client.get("window_key") or client.get("name"))
            ss._close_combo_leftover_dialogs(mw)
            print("[OK] 已关闭所有组合申报对话框（未提交报单）")
        except Exception as exc:  # noqa: BLE001
            print("[WARN] 清理对话框失败:", exc)
    ok = sum(1 for r in results if r.get("closed"))
    print(f"[汇总] 共 {len(results)} 个组合，安全填表并关闭 {ok} 个")


if __name__ == "__main__":
    main()
