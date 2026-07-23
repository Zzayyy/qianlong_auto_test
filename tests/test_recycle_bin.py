# -*- coding: utf-8 -*-

import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch


GUI_DIR = str(Path(__file__).resolve().parents[1] / "GUI自动化工具2")
if GUI_DIR not in sys.path:
    sys.path.insert(0, GUI_DIR)

from gui.recycle_bin import (
    move_directory_to_recycle_bin,
    validate_batch_directory,
)


class BatchRecycleBinTests(unittest.TestCase):
    @staticmethod
    def _make_batch(output_dir, run_id="20260723_153045_A7K9Q2"):
        batch_dir = Path(output_dir) / "批次" / run_id
        batch_dir.mkdir(parents=True)
        (batch_dir / "批次汇总.json").write_text(
            json.dumps({"run_id": run_id}, ensure_ascii=False),
            encoding="utf-8",
        )
        return batch_dir

    def test_validation_accepts_only_matching_direct_batch_directory(self):
        with tempfile.TemporaryDirectory() as output_dir:
            run_id = "20260723_153045_A7K9Q2"
            batch_dir = self._make_batch(output_dir, run_id)
            validated = validate_batch_directory(
                batch_dir, output_dir, run_id
            )
            self.assertEqual(batch_dir.resolve(), validated)

            with self.assertRaises(ValueError):
                validate_batch_directory(batch_dir, output_dir, "OTHER")

    def test_validation_rejects_directory_outside_batch_root(self):
        with tempfile.TemporaryDirectory() as output_dir:
            run_id = "20260723_153045_A7K9Q2"
            outside = Path(output_dir) / run_id
            outside.mkdir()
            (outside / "批次汇总.json").write_text(
                json.dumps({"run_id": run_id}),
                encoding="utf-8",
            )
            with self.assertRaises(ValueError):
                validate_batch_directory(outside, output_dir, run_id)

    def test_recycle_operation_runs_in_external_process(self):
        with tempfile.TemporaryDirectory() as output_dir:
            batch_dir = self._make_batch(output_dir)

            def emulate_recycle(*args, **kwargs):
                shutil.rmtree(batch_dir)
                return subprocess.CompletedProcess(args[0], 0, "", "")

            with patch(
                "gui.recycle_bin.subprocess.run",
                side_effect=emulate_recycle,
            ) as run:
                move_directory_to_recycle_bin(batch_dir)

            command = run.call_args.args[0]
            self.assertEqual("powershell.exe", command[0])
            self.assertIn("SendToRecycleBin", command[-1])
            self.assertEqual(
                str(batch_dir.resolve()),
                run.call_args.kwargs["env"]["GUI_REPORT_BATCH_TO_RECYCLE"],
            )

    def test_failed_recycle_operation_keeps_batch_and_raises(self):
        with tempfile.TemporaryDirectory() as output_dir:
            batch_dir = self._make_batch(output_dir)
            failed = subprocess.CompletedProcess(
                ["powershell.exe"], 1, "", "access denied"
            )
            with patch(
                "gui.recycle_bin.subprocess.run",
                return_value=failed,
            ):
                with self.assertRaisesRegex(OSError, "access denied"):
                    move_directory_to_recycle_bin(batch_dir)
            self.assertTrue(batch_dir.is_dir())


if __name__ == "__main__":
    unittest.main()
