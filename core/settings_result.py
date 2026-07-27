# -*- coding: utf-8 -*-
"""交易系统设置模块共用的结果收集、展示和结构化输出。"""

from __future__ import annotations

from datetime import datetime
import json
import os
from pathlib import Path
import re
from typing import Any, Callable, Dict, List, Optional

from core.settings_status import (
    DIFFERENCE_STATUSES,
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_DIFFERENCE,
    STATUS_DISABLED,
    STATUS_EXECUTION_FAILED,
    STATUS_NOT_APPLICABLE,
    STATUS_PASS,
    STATUS_UNVERIFIED,
    count_statuses,
    normalize_status,
)


STATUS_ORDER = (
    STATUS_PASS,
    STATUS_DIFFERENCE,
    STATUS_ADDED,
    STATUS_CONFLICT,
    STATUS_UNVERIFIED,
    STATUS_DISABLED,
    STATUS_EXECUTION_FAILED,
    STATUS_NOT_APPLICABLE,
)


class SettingsTestResult:
    """统一的设置检查结果模型。

    所有结果行固定包含 ``名称/期望值/实际值/状态/说明`` 五个字段。
    ``to_file`` 在写入文本报告的同时生成同名 JSON，供批次总报告读取。
    """

    def __init__(
        self,
        panel_name: str,
        *,
        normalizer: Optional[Callable[[Any], Any]] = None,
    ):
        self.panel_name = panel_name
        self.normalizer = normalizer
        self.results: List[Dict[str, Any]] = []
        self.observations: List[Dict[str, Any]] = []

    @property
    def differences(self) -> List[Dict[str, Any]]:
        return [row for row in self.results if row["状态"] in DIFFERENCE_STATUSES]

    @property
    def unverified(self) -> List[Dict[str, Any]]:
        return [row for row in self.results if row["状态"] == STATUS_UNVERIFIED]

    @property
    def not_enabled(self) -> int:
        return sum(row["状态"] == STATUS_DISABLED for row in self.results)

    def _values_match(self, actual_value: Any, expected_value: Any) -> bool:
        if self.normalizer is None:
            return actual_value == expected_value
        return self.normalizer(actual_value) == self.normalizer(expected_value)

    def add_result(
        self,
        name: str,
        actual_value: Any,
        expected_value: Any,
        detail: str = "",
    ):
        matched = self._values_match(actual_value, expected_value)
        self.add_status(
            name,
            actual_value,
            expected_value,
            STATUS_PASS if matched else STATUS_DIFFERENCE,
            detail,
        )

    def add_status(
        self,
        name: str,
        actual_value: Any,
        expected_value: Any,
        status: str,
        detail: str = "",
    ):
        self.results.append({
            "名称": name,
            "期望值": expected_value,
            "实际值": actual_value,
            "状态": normalize_status(status),
            "说明": detail,
        })

    def add_not_enabled(self, name: str, detail: str = "(未启用)"):
        self.add_status(name, detail, "—", STATUS_DISABLED)

    def add_unverified(
        self,
        name: str,
        expected_value: Any,
        detail: str,
        actual_value: Any = "(无法确认)",
    ):
        self.add_status(
            name, actual_value, expected_value, STATUS_UNVERIFIED, detail
        )

    def add_observation(self, name: str, value: Any, detail: str = ""):
        self.observations.append({
            "名称": name,
            "采集值": value,
            "说明": detail,
        })

    def summary(self) -> Dict[str, int]:
        counts = count_statuses(self.results)
        summary = {"总项目数": len(self.results)}
        summary.update({status: counts[status] for status in STATUS_ORDER})
        summary["差异合计"] = sum(counts[status] for status in DIFFERENCE_STATUSES)
        summary["采集项"] = len(self.observations)
        return summary

    def print_summary(self):
        summary = self.summary()
        print(f"\n{'=' * 60}")
        print("测试结果汇总")
        print(f"{'=' * 60}")
        self._write_summary(print, summary)
        self._write_rows(print)

    def _write_summary(self, writer, summary: Dict[str, int]):
        writer(f"总项目数: {summary['总项目数']}")
        for status in STATUS_ORDER:
            writer(f"{status}: {summary[status]}")
        writer(f"差异合计: {summary['差异合计']}")
        writer(f"采集项: {summary['采集项']}")

    def _write_rows(self, writer):
        if self.results:
            writer("")
        for row in self.results:
            writer(f"[{row['状态']}] {row['名称']}")
            writer(f"  期望值: {row['期望值']}")
            writer(f"  实际值: {row['实际值']}")
            if row["说明"]:
                writer(f"  说明: {row['说明']}")
        if self.observations:
            writer("")
            writer("采集项（不计入差异）:")
            for row in self.observations:
                writer(f"  - {row['名称']}: {row['采集值']}")
                if row["说明"]:
                    writer(f"    说明: {row['说明']}")

    def to_payload(self, report_path: str = "") -> Dict[str, Any]:
        generated_at = datetime.now().isoformat(timespec="seconds")
        return {
            "schema_version": 1,
            "run_id": os.environ.get("GUI_SETTINGS_RUN_ID", ""),
            "module": self.panel_name,
            "client_id": os.environ.get("GUI_CLIENT_ID", ""),
            "execution_status": STATUS_PASS,
            "generated_at": generated_at,
            "report_path": os.path.abspath(report_path) if report_path else "",
            "summary": self.summary(),
            "items": [
                {
                    "name": row["名称"],
                    "expected": row["期望值"],
                    "actual": row["实际值"],
                    "status": row["状态"],
                    "detail": row["说明"],
                }
                for row in self.results
            ],
            "observations": [
                {
                    "name": row["名称"],
                    "value": row["采集值"],
                    "detail": row["说明"],
                }
                for row in self.observations
            ],
        }

    def to_file(self, filepath: str):
        run_id = os.environ.get("GUI_SETTINGS_RUN_ID", "").strip()
        run_dir = os.environ.get("GUI_SETTINGS_RUN_DIR", "").strip()
        if run_id and run_dir:
            safe_panel_name = re.sub(r'[<>:"/\\|?*]+', "_", self.panel_name).strip()
            filepath = os.path.join(run_dir, f"{safe_panel_name}.txt")

        output_dir = os.path.dirname(os.path.abspath(filepath))
        os.makedirs(output_dir, exist_ok=True)
        generated_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        summary = self.summary()
        text_temp = f"{filepath}.tmp"
        with open(text_temp, "w", encoding="utf-8") as report:
            report.write(f"{self.panel_name}测试报告\n")
            report.write(f"生成时间: {generated_at}\n")
            report.write(f"{'=' * 60}\n\n")

            def write_line(value: str):
                report.write(f"{value}\n")

            self._write_summary(write_line, summary)
            self._write_rows(write_line)
        os.replace(text_temp, filepath)

        json_path = str(Path(filepath).with_suffix(".json"))
        json_temp = f"{json_path}.tmp"
        with open(json_temp, "w", encoding="utf-8") as structured:
            json.dump(
                self.to_payload(filepath),
                structured,
                ensure_ascii=False,
                indent=2,
                default=_json_default,
            )
            structured.write("\n")
        os.replace(json_temp, json_path)

        print(f"[OK] 测试报告已保存: {filepath}")
        print(f"[OK] 结构化结果已保存: {json_path}")
        return json_path


def _json_default(value: Any):
    """兼容 numpy 标量、Path 以及少数控件包装值。"""
    if hasattr(value, "item"):
        try:
            return value.item()
        except Exception:
            pass
    if isinstance(value, Path):
        return str(value)
    return str(value)
