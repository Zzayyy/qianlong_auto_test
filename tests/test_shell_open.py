# -*- coding: utf-8 -*-

from pathlib import Path
import sys
import tempfile
import unittest
from unittest.mock import patch


GUI_DIR = str(Path(__file__).resolve().parents[1] / "GUI自动化工具2")
if GUI_DIR not in sys.path:
    sys.path.insert(0, GUI_DIR)

from gui.shell_open import open_path


class SafeShellOpenTests(unittest.TestCase):
    def test_directory_uses_external_explorer_process(self):
        with tempfile.TemporaryDirectory() as directory:
            with patch("gui.shell_open.subprocess.Popen") as popen:
                open_path(directory)
            command = popen.call_args.args[0]
            self.assertEqual("explorer.exe", command[0])
            self.assertEqual(str(Path(directory).resolve()), command[1])

    def test_file_association_runs_outside_gui_process(self):
        with tempfile.TemporaryDirectory() as directory:
            report = Path(directory) / "总差异报告.txt"
            report.write_text("test", encoding="utf-8")
            with patch("gui.shell_open.subprocess.Popen") as popen:
                open_path(report)
            command = popen.call_args.args[0]
            self.assertEqual("rundll32.exe", command[0])
            self.assertEqual("url.dll,FileProtocolHandler", command[1])
            self.assertEqual(str(report.resolve()), command[2])

    def test_missing_path_is_rejected_before_launch(self):
        with patch("gui.shell_open.subprocess.Popen") as popen:
            with self.assertRaises(FileNotFoundError):
                open_path(r"C:\definitely-missing\report.txt")
        popen.assert_not_called()


if __name__ == "__main__":
    unittest.main()
