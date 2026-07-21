import unittest

from core.one_click_settings import (
    canonical_hotkey,
    evaluate_shortcuts,
    merge_shortcut_pages,
    normalize_text,
    parse_shortcut_ocr_tokens,
)


def token(text, x, y, score=0.99):
    return {
        "text": text,
        "score": score,
        "box": [[x - 2, y - 2], [x + 2, y - 2],
                [x + 2, y + 2], [x - 2, y + 2]],
    }


class NormalizeTests(unittest.TestCase):
    def test_normalizes_hotkey_spacing_and_width(self):
        self.assertEqual(normalize_text("  ＣＴＲＬ＋Ｓ  "), "CTRL+S")
        self.assertEqual(canonical_hotkey("ctrl + s"), "CTRL+S")
        self.assertEqual(canonical_hotkey("小键盘 4"), "小键盘4")


class OcrParserTests(unittest.TestCase):
    def test_groups_split_shortcut_tokens_by_row_and_column(self):
        rows = parse_shortcut_ocr_tokens(
            [
                token("5", 30, 50),
                token("平合约1义务仓", 100, 50),
                token("小键盘", 250, 50),
                token("4", 310, 50),
            ]
        )
        self.assertEqual(
            rows,
            [{
                "sequence": 5,
                "name": "平合约1义务仓",
                "shortcut": "小键盘4",
                "confidence": 0.99,
                "source": "OCR",
            }],
        )

    def test_applies_explicit_ocr_name_alias(self):
        rows = parse_shortcut_ocr_tokens(
            [token("10", 30, 50), token("全撒", 100, 50, 0.8),
             token("小键盘.", 250, 50, 0.9)],
            name_aliases={"全撒": "全撤"},
        )
        self.assertEqual(rows[0]["name"], "全撤")
        self.assertEqual(rows[0]["shortcut"], "小键盘.")
        self.assertEqual(rows[0]["confidence"], 0.8)

    def test_merges_overlapping_pages_using_more_complete_row(self):
        merged = merge_shortcut_pages(
            [
                [{"sequence": 12, "name": "减少张数", "shortcut": "小键盘",
                  "confidence": 0.99}],
                [{"sequence": 12, "name": "减少张数", "shortcut": "小键盘-",
                  "confidence": 0.90}],
            ]
        )
        self.assertEqual(merged[0]["shortcut"], "小键盘-")


class ShortcutEvaluationTests(unittest.TestCase):
    def setUp(self):
        self.expected = [
            {"sequence": 1, "name": "启用一键炒单", "shortcut": "CTRL+S"},
            {"sequence": 2, "name": "买合约1", "shortcut": "小键盘1"},
        ]

    def test_matching_rows_and_no_conflicts_pass(self):
        actual = [
            {**self.expected[0], "confidence": 1.0},
            {**self.expected[1], "confidence": 1.0},
        ]
        checks = evaluate_shortcuts(self.expected, actual, source="原生ListView")
        self.assertTrue(all(row["status"] == "通过" for row in checks))

    def test_incomplete_ocr_hotkey_is_unverified(self):
        actual = [
            {"sequence": 1, "name": "启用一键炒单", "shortcut": "CTRL+S",
             "confidence": 0.99},
            {"sequence": 2, "name": "买合约1", "shortcut": "小键盘",
             "confidence": 0.99},
        ]
        checks = evaluate_shortcuts(self.expected, actual, source="OCR分页")
        target = next(row for row in checks if row["name"] == "快捷键[2]_买合约1")
        self.assertEqual(target["status"], "未验证")

    def test_missing_ocr_row_is_unverified_not_difference(self):
        checks = evaluate_shortcuts(
            self.expected, [], source="OCR分页"
        )
        missing = [row for row in checks if row["name"].startswith("快捷键[")]
        self.assertTrue(missing)
        self.assertTrue(all(row["status"] == "未验证" for row in missing))

    def test_detects_missing_rows_and_hotkey_conflicts(self):
        expected = self.expected + [
            {"sequence": 3, "name": "卖合约1", "shortcut": "小键盘2"}
        ]
        actual = [
            {"sequence": 1, "name": "启用一键炒单", "shortcut": "CTRL+S"},
            {"sequence": 2, "name": "买合约1", "shortcut": "CTRL+S"},
        ]
        checks = evaluate_shortcuts(expected, actual, source="原生ListView")
        statuses = {row["status"] for row in checks}
        self.assertIn("差异", statuses)
        self.assertIn("冲突", statuses)


if __name__ == "__main__":
    unittest.main()
