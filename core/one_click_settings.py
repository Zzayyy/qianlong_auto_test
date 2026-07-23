# -*- coding: utf-8 -*-
"""Pure parsing and validation helpers for one-click trading settings."""

from __future__ import annotations

from collections import defaultdict
import re
import unicodedata
from typing import Any, Iterable


HEADER_TEXTS = {"序号", "选项名称", "当前快捷键(点击设置)"}


def normalize_text(value: Any) -> str:
    """Normalize client/OCR text without changing its business meaning."""
    text = unicodedata.normalize("NFKC", "" if value is None else str(value))
    return re.sub(r"\s+", " ", text).strip()


def canonical_hotkey(value: Any) -> str:
    """Normalize harmless spacing/case differences in displayed hotkeys."""
    text = normalize_text(value).upper()
    text = re.sub(r"\s*([+\-.])\s*", r"\1", text)
    text = re.sub(r"小键盘\s+([0-9])", r"小键盘\1", text)
    return text.replace("CTRL +", "CTRL+")


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
        shortcut = "".join(item["text"] for item in columns["shortcut"])
        if normalize_text(name) in HEADER_TEXTS:
            continue
        normalized_name = aliases.get(normalize_text(name), normalize_text(name))
        used = columns["sequence"] + columns["name"] + columns["shortcut"]
        rows.append(
            {
                "sequence": sequence,
                "name": normalized_name,
                "shortcut": canonical_hotkey(shortcut),
                "confidence": min((item["score"] for item in used), default=0.0),
                "source": "OCR",
            }
        )
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


def evaluate_shortcuts(
    expected_rows: list[dict[str, Any]],
    actual_rows: list[dict[str, Any]],
    *,
    source: str,
    min_ocr_confidence: float = 0.75,
) -> list[dict[str, Any]]:
    """Return pass/difference/unverified checks for shortcut rows."""
    checks: list[dict[str, Any]] = []
    expected_by_sequence = {int(row["sequence"]): row for row in expected_rows}
    actual_by_sequence = {int(row["sequence"]): row for row in actual_rows}
    is_ocr = source.upper().startswith("OCR")

    for sequence, expected in expected_by_sequence.items():
        actual = actual_by_sequence.get(sequence)
        label = f"快捷键[{sequence}]_{expected['name']}"
        if actual is None:
            checks.append(
                {
                    "name": label,
                    "expected": canonical_hotkey(expected["shortcut"]),
                    "actual": "未读取到该行",
                    "status": "未验证" if is_ocr else "差异",
                    "detail": (
                        "OCR分页未能确认该行"
                        if is_ocr else "标准行缺失"
                    ),
                }
            )
            continue

        confidence = float(actual.get("confidence", 1.0))
        if is_ocr and confidence < min_ocr_confidence:
            checks.append(
                {
                    "name": label,
                    "expected": canonical_hotkey(expected["shortcut"]),
                    "actual": f"{actual.get('name', '')} -> {actual.get('shortcut', '')}",
                    "status": "未验证",
                    "detail": f"OCR最低置信度 {confidence:.3f}",
                }
            )
            continue

        expected_name = normalize_text(expected["name"])
        actual_name = normalize_text(actual.get("name"))
        expected_hotkey = canonical_hotkey(expected["shortcut"])
        actual_hotkey = canonical_hotkey(actual.get("shortcut"))
        if actual_name != expected_name:
            checks.append(
                {
                    "name": f"{label}_名称",
                    "expected": expected_name,
                    "actual": actual_name or "(空白)",
                    "status": "差异",
                    "detail": "快捷键项目名称不一致",
                }
            )

        if not actual_hotkey or actual_hotkey == "小键盘":
            status = "未验证" if is_ocr else "差异"
            detail = "OCR未识别到完整按键" if is_ocr else "快捷键为空或不完整"
        elif actual_hotkey == expected_hotkey:
            status = "通过"
            detail = source
        else:
            status = "差异"
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

    for sequence, actual in actual_by_sequence.items():
        if sequence not in expected_by_sequence:
            checks.append(
                {
                    "name": f"快捷键[{sequence}]_{actual.get('name', '未知项目')}",
                    "expected": "标准配置中不存在",
                    "actual": canonical_hotkey(actual.get("shortcut")),
                    "status": "新增",
                    "detail": "客户端出现未配置的新项目",
                }
            )

    by_hotkey: dict[str, list[str]] = defaultdict(list)
    for row in actual_rows:
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
                    "status": "冲突",
                    "detail": "同一个按键分配给多个项目",
                }
            )
    else:
        checks.append(
            {
                "name": "快捷键冲突检查",
                "expected": "无冲突",
                "actual": "无冲突",
                "status": "通过",
                "detail": source,
            }
        )
    return checks
