# -*- coding: utf-8 -*-

import json
from datetime import datetime
from pathlib import Path
import tempfile
import unittest

from core.settings_report import (
    BATCH_SUMMARY_JSON,
    BATCH_RUNNING,
    OVERALL_FAIL,
    OVERALL_RUNNING,
    OVERALL_REVIEW,
    TOTAL_REPORT_TXT,
    TOTAL_REPORT_XLSX,
    create_run_id,
    discover_batches,
    generate_batch_reports,
)
from core.settings_status import STATUS_DIFFERENCE, STATUS_PASS, STATUS_UNVERIFIED


class SettingsBatchReportTests(unittest.TestCase):
    def test_run_id_contains_timestamp_and_dedicated_code(self):
        run_id = create_run_id(
            datetime(2026, 7, 23, 15, 30, 45),
            code="a7k9q2",
        )
        self.assertEqual("20260723_153045_A7K9Q2", run_id)

    def _write_module(self, batch_dir, module, items):
        counts = {
            STATUS_PASS: 0,
            STATUS_DIFFERENCE: 0,
            "新增": 0,
            "冲突": 0,
            STATUS_UNVERIFIED: 0,
            "未启用": 0,
            "不适用": 0,
        }
        for item in items:
            counts[item["status"]] += 1
        payload = {
            "schema_version": 1,
            "run_id": "run-1",
            "module": module,
            "client_id": "client-a",
            "generated_at": "2026-07-23T10:00:00",
            "report_path": str(Path(batch_dir) / f"{module}.txt"),
            "summary": {
                "总项目数": len(items),
                **counts,
                "执行失败": 0,
                "差异合计": counts[STATUS_DIFFERENCE] + counts["新增"] + counts["冲突"],
                "采集项": 0,
            },
            "items": items,
            "observations": [],
        }
        path = Path(batch_dir) / f"{module}.json"
        path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")

    @staticmethod
    def _task(script_name, status="成功", return_code=0):
        return {
            "category": "交易系统设置",
            "script_name": script_name,
            "script_path": "dummy.py",
            "params": {},
            "status": status,
            "return_code": return_code,
            "elapsed": 1.25,
            "error": "",
        }

    def test_generates_total_reports_and_discovers_batch(self):
        with tempfile.TemporaryDirectory() as output_dir:
            batch_dir = Path(output_dir) / "批次" / "run-1"
            batch_dir.mkdir(parents=True)
            self._write_module(
                batch_dir,
                "委托设置",
                [
                    {
                        "name": "默认价格",
                        "expected": "对手价",
                        "actual": "最新价",
                        "status": STATUS_DIFFERENCE,
                        "detail": "",
                    }
                ],
            )

            summary = generate_batch_reports(
                run_id="run-1",
                batch_dir=str(batch_dir),
                client_id="client-a",
                task_records=[self._task("1.委托设置")],
            )

            self.assertEqual(OVERALL_FAIL, summary["overall_status"])
            self.assertEqual(1, summary["totals"]["差异合计"])
            self.assertEqual(1, len(summary["problems"]))
            self.assertTrue((batch_dir / BATCH_SUMMARY_JSON).is_file())
            self.assertTrue((batch_dir / TOTAL_REPORT_TXT).is_file())
            self.assertTrue((batch_dir / TOTAL_REPORT_XLSX).is_file())
            self.assertEqual("run-1", discover_batches(output_dir)[0]["run_id"])

    def test_running_summary_records_source_without_marking_pending_failures(self):
        with tempfile.TemporaryDirectory() as output_dir:
            batch_dir = Path(output_dir) / "批次" / "running"
            batch_dir.mkdir(parents=True)
            summary = generate_batch_reports(
                run_id="running",
                batch_dir=str(batch_dir),
                client_id="client-a",
                task_records=[],
                source="任务中心",
                batch_status=BATCH_RUNNING,
            )
            self.assertEqual(OVERALL_RUNNING, summary["overall_status"])
            self.assertEqual(BATCH_RUNNING, summary["batch_status"])
            self.assertEqual("任务中心", summary["source"])
            self.assertEqual(0, summary["totals"]["执行失败"])

    def test_unverified_is_review_and_missing_json_is_failure(self):
        with tempfile.TemporaryDirectory() as output_dir:
            review_dir = Path(output_dir) / "批次" / "review"
            review_dir.mkdir(parents=True)
            self._write_module(
                review_dir,
                "期权设置",
                [
                    {
                        "name": "快捷键",
                        "expected": "F1",
                        "actual": "无法确认",
                        "status": STATUS_UNVERIFIED,
                        "detail": "OCR 不完整",
                    }
                ],
            )
            review = generate_batch_reports(
                run_id="review",
                batch_dir=str(review_dir),
                client_id="client-a",
                task_records=[self._task("2.期权设置")],
            )
            self.assertEqual(OVERALL_REVIEW, review["overall_status"])

            failed_dir = Path(output_dir) / "批次" / "failed"
            failed_dir.mkdir(parents=True)
            failed = generate_batch_reports(
                run_id="failed",
                batch_dir=str(failed_dir),
                client_id="client-a",
                task_records=[self._task("3.自动拆单设置")],
            )
            self.assertEqual(OVERALL_FAIL, failed["overall_status"])
            self.assertEqual(1, failed["totals"]["执行失败"])
            self.assertIn("没有生成结构化结果", failed["modules"][0]["detail"])


class SettingsTaskEnvironmentTests(unittest.TestCase):
    def test_batch_directory_and_run_id_override_regular_output(self):
        import sys

        gui_dir = str(Path(__file__).resolve().parents[1] / "GUI自动化工具2")
        if gui_dir not in sys.path:
            sys.path.insert(0, gui_dir)
        from engine.task import Task

        task = Task(
            {"name": "1.委托设置", "path": "dummy.py"},
            "交易系统设置",
            {
                "settings_output_dir": r"D:\普通输出",
                "settings_run_dir": r"D:\普通输出\批次\run-1",
                "settings_run_id": "run-1",
            },
        )
        env = task.build_env(r"D:\自动测试")
        self.assertEqual(r"D:\普通输出\批次\run-1", env["GUI_OUTPUT_DIR"])
        self.assertEqual("run-1", env["GUI_SETTINGS_RUN_ID"])
        self.assertEqual(r"D:\普通输出\批次\run-1", env["GUI_SETTINGS_RUN_DIR"])


if __name__ == "__main__":
    unittest.main()
