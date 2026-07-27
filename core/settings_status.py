# -*- coding: utf-8 -*-
"""交易系统设置检查使用的统一状态定义与兼容映射。"""

from __future__ import annotations

from collections import Counter
from typing import Iterable, Mapping


STATUS_PASS = "通过"
STATUS_DIFFERENCE = "差异"
STATUS_DISABLED = "未启用"
STATUS_UNVERIFIED = "未验证"
STATUS_ADDED = "新增"
STATUS_CONFLICT = "冲突"
STATUS_EXECUTION_FAILED = "执行失败"
STATUS_NOT_APPLICABLE = "不适用"

CANONICAL_STATUSES = frozenset({
    STATUS_PASS,
    STATUS_DIFFERENCE,
    STATUS_DISABLED,
    STATUS_UNVERIFIED,
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_EXECUTION_FAILED,
    STATUS_NOT_APPLICABLE,
})

DIFFERENCE_STATUSES = frozenset({
    STATUS_DIFFERENCE,
    STATUS_ADDED,
    STATUS_CONFLICT,
})

REVIEW_STATUSES = frozenset({STATUS_UNVERIFIED})

_STATUS_ALIASES: Mapping[str, str] = {
    "✓": STATUS_PASS,
    "一致": STATUS_PASS,
    "pass": STATUS_PASS,
    "passed": STATUS_PASS,
    "✗": STATUS_DIFFERENCE,
    "✗ 差异": STATUS_DIFFERENCE,
    "不一致": STATUS_DIFFERENCE,
    "diff": STATUS_DIFFERENCE,
    "different": STATUS_DIFFERENCE,
    "○ 未启用": STATUS_DISABLED,
    "未生效": STATUS_DISABLED,
    "disabled": STATUS_DISABLED,
    "unverified": STATUS_UNVERIFIED,
    "added": STATUS_ADDED,
    "conflict": STATUS_CONFLICT,
    "失败": STATUS_EXECUTION_FAILED,
    "异常": STATUS_EXECUTION_FAILED,
    "error": STATUS_EXECUTION_FAILED,
    "failed": STATUS_EXECUTION_FAILED,
    "跳过": STATUS_NOT_APPLICABLE,
    "不支持": STATUS_NOT_APPLICABLE,
    "skip": STATUS_NOT_APPLICABLE,
    "skipped": STATUS_NOT_APPLICABLE,
}


def normalize_status(status: object) -> str:
    """将历史显示值或英文状态转换为统一中文状态。

    未知状态会直接报错，避免汇总时悄悄漏掉新增状态。
    """
    value = str(status).strip()
    if value in CANONICAL_STATUSES:
        return value
    normalized = _STATUS_ALIASES.get(value)
    if normalized is None:
        normalized = _STATUS_ALIASES.get(value.lower())
    if normalized is None:
        raise ValueError(f"未知的交易系统设置状态: {status!r}")
    return normalized


def is_difference_status(status: object) -> bool:
    """返回该状态是否应计入总差异。"""
    return normalize_status(status) in DIFFERENCE_STATUSES


def count_statuses(rows: Iterable[Mapping[str, object]]) -> Counter:
    """按统一的 ``状态`` 字段统计结果行。"""
    return Counter(normalize_status(row["状态"]) for row in rows)
