# -*- coding: utf-8 -*-
"""Pure parsing and validation helpers for one-click trading settings.
    用于交易系统设置的 一键炒单脚本 OCR后解析+比对
"""

from __future__ import annotations

from collections import defaultdict
import re
import unicodedata
from typing import Any, Iterable

from core.settings import (
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_DIFFERENCE,
    STATUS_PASS,
    STATUS_UNVERIFIED,
)


HEADER_TEXTS = {"序号", "选项名称", "当前快捷键(点击设置)"}

# 小键盘后的按键位只可能是 0-9 / + - . 中的一个；用于找回被 OCR 单独成行的细按键。
HOTKEY_KEY_RE = re.compile(r"^[0-9+\-.]$")


def normalize_text(value: Any) -> str:
    """Normalize client/OCR text without changing its business meaning."""
    text = unicodedata.normalize("NFKC", "" if value is None else str(value))
    return re.sub(r"\s+", " ", text).strip()


# 仅作用于“小键盘”之后的按键字符；这些位置只可能是 0-9 / + - . 。
# OCR 常见的形似混淆（NFKC 不会转换的希腊/西里尔字母、中文“一”误作减号）。
# 注意：绝不映射 CTRL+S 中的字母（如 S），只对“小键盘”按键位做上下文修正。
_HOTKEY_KEY_CORRECTIONS = {
    "\u0395": "3",  # GREEK CAPITAL LETTER EPSILON -> 3
    "\u0415": "3",  # CYRILLIC CAPITAL LETTER IE -> 3
    "\u039F": "0",  # GREEK CAPITAL LETTER OMICRON -> 0
    "\u041E": "0",  # CYRILLIC CAPITAL LETTER O -> 0
    "一": "-",      # 中文“一”常被误识为减号
}


def canonical_hotkey(value: Any) -> str:
    """Normalize harmless spacing/case differences in displayed hotkeys.

    对“小键盘”后的按键位额外做上下文字形修正（例如 OCR 把 ``3`` 识成
    希腊大写 Epsilon ``Ε`` 时归正为 ``小键盘3``），避免由此产生的假差异。
    """
    text = normalize_text(value).upper()
    text = re.sub(r"\s*([+\-.])\s*", r"\1", text)
    text = re.sub(r"小键盘\s+([0-9])", r"小键盘\1", text)
    text = text.replace("CTRL +", "CTRL+")
    if text.startswith("小键盘"):
        prefix, _, key = text.partition("小键盘")
        corrected_key = "".join(
            _HOTKEY_KEY_CORRECTIONS.get(ch, ch) for ch in key
        )
        text = "小键盘" + corrected_key
    return text


def _recover_missing_key(positioned, group, column_boundaries, tolerance) -> str:
    """OCR 常把“小键盘”后的细按键(-/+/./数字)单独成行，导致该行只剩“小键盘”而
    丢失按键。在同一视觉行的邻近右侧(容差放宽)找回孤立的按键令牌并拼回。"""
    if not group:
        return ""
    group_y = sum(t["y"] for t in group) / len(group)
    for tok in positioned:
        if tok in group:
            continue
        if not HOTKEY_KEY_RE.match(tok["text"]):
            continue
        if abs(tok["y"] - group_y) > tolerance:
            continue
        if tok["x"] < column_boundaries[1]:
            continue
        return tok["text"]
    return ""


def filter_phantom_rows(rows: Iterable[dict[str, Any]], row_count: int) -> list[dict[str, Any]]:
    """丢弃序号不在 1..row_count 的 OCR 幻影行。

    表格的真实行数由调用方通过 LVM_GETITEMCOUNT 取得；任何超出该范围的
    “序号”(如误识出的 40)都来自界面噪点，绝不可能对应真实行，应在比对前剔除，
    避免污染结果(否则会显示成「标准配置中不存在」的未验证项)。
    """
    if not isinstance(row_count, int) or row_count <= 0:
        return list(rows)
    return [r for r in rows if 1 <= int(r.get("sequence", -1)) <= row_count]


def _box_center(box: Iterable[Iterable[float]]) -> tuple[float, float]:
    points = list(box)
    return (
        sum(float(point[0]) for point in points) / len(points),
        sum(float(point[1]) for point in points) / len(points),
    )


def parse_shortcut_ocr_tokens(
    tokens: list[dict[str, Any]],
    column_boundaries: tuple[float, float] = (50.0, 165.0),
    y_tolerance: float = 6.0,
    name_aliases: dict[str, str] | None = None,
) -> list[dict[str, Any]]:
    """Convert positioned OCR tokens into shortcut table rows.

    The first boundary ends the sequence column and the second ends the name
    column. Tokens on the same visual row are joined from left to right, which
    also repairs OCR output such as ``小键盘`` + ``4``.
    """
    aliases = {
        normalize_text(key): normalize_text(value)
        for key, value in (name_aliases or {}).items()
    }
    positioned: list[dict[str, Any]] = []
    for token in tokens:
        text = normalize_text(token.get("text"))
        box = token.get("box")
        if not text or not box:
            continue
        x, y = _box_center(box)
        positioned.append(
            {
                "text": text,
                "score": float(token.get("score", 0.0)),
                "x": x,
                "y": y,
            }
        )

    groups: list[list[dict[str, Any]]] = []
    for token in sorted(positioned, key=lambda item: (item["y"], item["x"])):
        if not groups:
            groups.append([token])
            continue
        center_y = sum(item["y"] for item in groups[-1]) / len(groups[-1])
        if abs(token["y"] - center_y) <= y_tolerance:
            groups[-1].append(token)
        else:
            groups.append([token])

    rows: list[dict[str, Any]] = []
    for group in groups:
        columns = {"sequence": [], "name": [], "shortcut": []}
        for token in sorted(group, key=lambda item: item["x"]):
            if token["x"] < column_boundaries[0]:
                columns["sequence"].append(token)
            elif token["x"] < column_boundaries[1]:
                columns["name"].append(token)
            else:
                columns["shortcut"].append(token)

        sequence_text = "".join(item["text"] for item in columns["sequence"])
        match = re.search(r"\d+", sequence_text)
        if not match:
            continue
        sequence = int(match.group())
        name = "".join(item["text"] for item in columns["name"])
        raw_shortcut = "".join(item["text"] for item in columns["shortcut"])
        # OCR 可能把“小键盘”后的细按键(-/+/.)单独成组，导致该行只剩“小键盘”。
        # 在同一视觉行右侧(放宽容差)找回孤立的按键令牌并拼回。
        if raw_shortcut == "小键盘":
            recovered = _recover_missing_key(
                positioned, group, column_boundaries, y_tolerance * 2
            )
            if recovered:
                raw_shortcut = "小键盘" + recovered
        if normalize_text(name) in HEADER_TEXTS:
            continue
        normalized_name = aliases.get(normalize_text(name), normalize_text(name))
        used = columns["sequence"] + columns["name"] + columns["shortcut"]
        # 记录“小键盘”令牌的中心坐标(与传入 tokens 同一坐标系)，供调用方在
        # 细按键漏检时做“紧裁高倍复核”定位使用(内部字段，比对时不使用)。
        anchor = None
        for tok in columns["shortcut"]:
            if tok["text"].startswith("小键盘"):
                anchor = (tok["x"], tok["y"])
                break
        if anchor is None and columns["shortcut"]:
            anchor = (columns["shortcut"][0]["x"], columns["shortcut"][0]["y"])
        row = {
            "sequence": sequence,
            "name": normalized_name,
            "shortcut": canonical_hotkey(raw_shortcut),
            "confidence": min((item["score"] for item in used), default=0.0),
            "source": "OCR",
        }
        if anchor is not None:
            row["_anchor"] = anchor
        rows.append(row)
    return rows


def merge_shortcut_pages(pages: Iterable[list[dict[str, Any]]]) -> list[dict[str, Any]]:
    """Merge overlapping screenshots by sequence, keeping the best row."""
    best: dict[int, dict[str, Any]] = {}
    for page in pages:
        for row in page:
            sequence = int(row["sequence"])
            rank = (
                bool(row.get("name")) + bool(row.get("shortcut")),
                len(str(row.get("name", ""))) + len(str(row.get("shortcut", ""))),
                float(row.get("confidence", 0.0)),
            )
            current = best.get(sequence)
            if current is None:
                best[sequence] = row
                continue
            current_rank = (
                bool(current.get("name")) + bool(current.get("shortcut")),
                len(str(current.get("name", "")))
                + len(str(current.get("shortcut", ""))),
                float(current.get("confidence", 0.0)),
            )
            if rank > current_rank:
                best[sequence] = row
    return [best[key] for key in sorted(best)]


def _rank_actual_row(row: dict[str, Any]) -> tuple[int, int, float]:
    """业务完整度排序：键越完整、文本越长、置信度越高越优先（与合并同口径）。"""
    name = row.get("name") or ""
    shortcut = row.get("shortcut") or ""
    return (
        bool(name) + bool(shortcut),
        len(str(name)) + len(str(shortcut)),
        float(row.get("confidence", 0.0)),
    )


def _evaluate_one_shortcut(
    checks: list[dict[str, Any]],
    expected: dict[str, Any],
    actual: dict[str, Any],
    label: str,
    source: str,
    is_ocr: bool,
    min_ocr_confidence: float,
) -> None:
    """对单个“期望值行”与“已匹配到的 actual 行”产出一条比对结果。"""
    confidence = float(actual.get("confidence", 1.0))
    if is_ocr and confidence < min_ocr_confidence:
        checks.append(
            {
                "name": label,
                "expected": canonical_hotkey(expected["shortcut"]),
                "actual": f"{actual.get('name', '')} -> {actual.get('shortcut', '')}",
                "status": STATUS_UNVERIFIED,
                "detail": f"OCR最低置信度 {confidence:.3f}",
            }
        )
        return

    expected_hotkey = canonical_hotkey(expected["shortcut"])
    actual_hotkey = canonical_hotkey(actual.get("shortcut"))
    if not actual_hotkey or actual_hotkey == "小键盘":
        status = STATUS_UNVERIFIED if is_ocr else STATUS_DIFFERENCE
        detail = "OCR未识别到完整按键" if is_ocr else "快捷键为空或不完整"
    elif actual_hotkey == expected_hotkey:
        status = STATUS_PASS
        detail = source
    else:
        status = STATUS_DIFFERENCE
        detail = source
    checks.append(
        {
            "name": label,
            "expected": expected_hotkey,
            "actual": actual_hotkey or "(空白)",
            "status": status,
            "detail": detail,
        }
    )


def evaluate_shortcuts(
    expected_rows: list[dict[str, Any]],
    actual_rows: list[dict[str, Any]],
    *,
    source: str,
    min_ocr_confidence: float = 0.75,
    expected_row_count: int | None = None,
) -> list[dict[str, Any]]:
    """Return pass/difference/unverified checks for shortcut rows.

    匹配策略：选项名称是业务主键，序号只是显示顺序。因此**优先按名称**把
    actual 行对齐到期望值；只有名称无法匹配时才回退按序号。这样可避免 OCR
    把某一行内容误识成另一行的名称/按键时产生的“假差异”和“假冲突”，也避免
    重复 OCR 行（同一期望值被多行命中）造成的冲突。
    """
    checks: list[dict[str, Any]] = []
    expected_by_sequence = {int(row["sequence"]): row for row in expected_rows}
    expected_by_name = {normalize_text(row["name"]): row for row in expected_rows}
    is_ocr = source.upper().startswith("OCR")

    # 1) 按名称收集候选，每个期望值只保留一条最佳 actual（其余视为重复行）。
    name_candidates: dict[int, list[int]] = defaultdict(list)
    name_matched_indices: set[int] = set()
    for idx, actual in enumerate(actual_rows):
        exp = expected_by_name.get(normalize_text(actual.get("name")))
        if exp is None:
            continue
        name_candidates[int(exp["sequence"])].append(idx)
        name_matched_indices.add(idx)

    representative: dict[int, int] = {}
    for eseq, idxs in name_candidates.items():
        best_idx = idxs[0]
        best_rank = _rank_actual_row(actual_rows[best_idx])
        for other in idxs[1:]:
            other_rank = _rank_actual_row(actual_rows[other])
            if other_rank > best_rank:
                best_idx, best_rank = other, other_rank
        representative[eseq] = best_idx

    resolved_indices = set(representative.values())

    # 2) 评估按名称匹配到的行（标签使用期望值的序号与名称）。
    for eseq, idx in representative.items():
        exp = expected_by_sequence[eseq]
        label = f"快捷键[{eseq}]_{exp['name']}"
        _evaluate_one_shortcut(
            checks, exp, actual_rows[idx], label, source, is_ocr, min_ocr_confidence
        )

    # 3) 其余行按序号回退匹配；名称不符则显式报差异。
    for idx, actual in enumerate(actual_rows):
        if idx in name_matched_indices:
            continue
        aseq = int(actual["sequence"])
        exp = expected_by_sequence.get(aseq)
        if exp is None:
            if expected_row_count is not None and aseq > expected_row_count:
                checks.append(
                    {
                        "name": f"快捷键[{aseq}]_{actual.get('name', '未知项目')}",
                        "expected": "标准配置中不存在",
                        "actual": canonical_hotkey(actual.get("shortcut")),
                        "status": STATUS_UNVERIFIED,
                        "detail": (
                            f"疑似OCR幻影行(序号{aseq}超出表格范围"
                            f"{expected_row_count})"
                        ),
                    }
                )
            else:
                checks.append(
                    {
                        "name": f"快捷键[{aseq}]_{actual.get('name', '未知项目')}",
                        "expected": "标准配置中不存在",
                        "actual": canonical_hotkey(actual.get("shortcut")),
                        "status": STATUS_ADDED,
                        "detail": "客户端出现未配置的新项目",
                    }
                )
            continue
        if normalize_text(actual.get("name")) != normalize_text(exp["name"]):
            checks.append(
                {
                    "name": f"快捷键[{aseq}]_{exp['name']}_名称",
                    "expected": normalize_text(exp["name"]),
                    "actual": normalize_text(actual.get("name")) or "(空白)",
                    "status": STATUS_DIFFERENCE,
                    "detail": "快捷键项目名称不一致(OCR或客户端变更)",
                }
            )
        label = f"快捷键[{aseq}]_{exp['name']}"
        _evaluate_one_shortcut(
            checks, exp, actual, label, source, is_ocr, min_ocr_confidence
        )

    # 4) 期望值中无任何 actual 命中 -> 未验证（OCR 未能确认该行）。
    for eseq, exp in expected_by_sequence.items():
        if eseq in representative:
            continue
        label = f"快捷键[{eseq}]_{exp['name']}"
        checks.append(
            {
                "name": label,
                "expected": canonical_hotkey(exp["shortcut"]),
                "actual": "未读取到该行",
                "status": STATUS_UNVERIFIED if is_ocr else STATUS_DIFFERENCE,
                "detail": "OCR分页未能确认该行" if is_ocr else "标准行缺失",
            }
        )

    # 5) 按键冲突检测（仅在去重后的 actual 行上进行，忽略重复 OCR 行）。
    resolved_rows = [actual_rows[i] for i in resolved_indices]
    by_hotkey: dict[str, list[str]] = defaultdict(list)
    for row in resolved_rows:
        hotkey = canonical_hotkey(row.get("shortcut"))
        if hotkey and hotkey != "小键盘":
            by_hotkey[hotkey].append(normalize_text(row.get("name")))
    conflicts = {key: names for key, names in by_hotkey.items() if len(names) > 1}
    if conflicts:
        for hotkey, names in conflicts.items():
            checks.append(
                {
                    "name": f"快捷键冲突_{hotkey}",
                    "expected": "无冲突",
                    "actual": "、".join(names),
                    "status": STATUS_CONFLICT,
                    "detail": "同一个按键分配给多个项目",
                }
            )
    else:
        checks.append(
            {
                "name": "快捷键冲突检查",
                "expected": "无冲突",
                "actual": "无冲突",
                "status": STATUS_PASS,
                "detail": source,
            }
        )
    return checks
