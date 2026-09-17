"""
test_basic.py
基础测试：异常场景 + 5类意图正常通路
执行：python -m unittest tests.test_basic -v
"""
import unittest
import subprocess
import json
import os

PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
TEST_DATA_DIR = os.path.join(PROJECT_ROOT, "tests", "test_data")

env = os.environ.copy()
env["PYTHONPATH"] = PROJECT_ROOT


def run_cli(args_list):
    cmd = ["python", CLI_SCRIPT] + args_list
    result = subprocess.run(cmd, capture_output=True, text=False, env=env)
    try:
        stdout_text = result.stdout.decode("utf-8")
    except UnicodeDecodeError:
        stdout_text = result.stdout.decode("gbk", errors="replace")
    json_start = stdout_text.find("{")
    if json_start == -1:
        raise RuntimeError(f"输出中未找到JSON对象，stdout_text:\n{stdout_text}")
    return json.loads(stdout_text[json_start:])


class TestCliErrorCases(unittest.TestCase):
    """异常用例 E01-E04"""
    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_E01_file_not_found(self):
        """E01 文件不存在 → FILE_NOT_FOUND"""
        resp = run_cli(["--file", "no_file.xlsx", "--intent", "overview"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "FILE_NOT_FOUND")

    def test_E02_invalid_intent(self):
        """E02 非法意图 → INVALID_INTENT（使用真实存在的文件）"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "wrong_intent"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "INVALID_INTENT")

    def test_E03_trend_missing_arg(self):
        """E03 trend 缺少必填参数 → MISSING_ARG（使用真实存在的文件）"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "trend"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "MISSING_ARG")

    def test_E04_invalid_suffix(self):
        """E04 非法后缀 → INVALID_SUFFIX"""
        bad_file = os.path.join(TEST_DATA_DIR, "bad_suffix.txt")
        os.makedirs(TEST_DATA_DIR, exist_ok=True)
        if not os.path.exists(bad_file):
            with open(bad_file, "w", encoding="utf-8") as f:
                f.write("not excel")
        resp = run_cli(["--file", bad_file, "--intent", "overview"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "INVALID_SUFFIX")


class TestCliFunctionCases(unittest.TestCase):
    """功能测试 F01-F05 + E05"""

    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_E05_sheet_not_exist(self):
        """E05 指定不存在 sheet → SHEET_NOT_EXIST"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "overview",
                        "--sheet", "NotExistSheet"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "SHEET_NOT_EXIST")

    def test_F01_overview(self):
        """F01 数据概览"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "overview"])
        self.assertTrue(resp["success"])
        self.assertIn("row_count", resp["result"])

    def test_F02_stats(self):
        """F02 描述性统计（含 cv/skew）"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "stats"])
        self.assertTrue(resp["success"])
        stats_result = resp["result"]["stats_result"]
        if stats_result:
            any_col = next(iter(stats_result.values()))
            self.assertIn("cv", any_col)
            self.assertIn("skew", any_col)

    def test_F03_sort_filter(self):
        """F03 排序筛选"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "sort_filter",
            "--sort_col", "销售额",
            "--filter_condition", "销售额>500",
        ])
        self.assertTrue(resp["success"])
        self.assertIn("hit_rows", resp["result"])

    def test_F04_trend_autoclean(self):
        """F04 趋势分析：样本含非法时间行，auto_clean 应自动剔除并成功"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--time_col", "日期",
            "--value_col", "销售额",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        self.assertIn("clean_report", resp["result"])
        self.assertGreaterEqual(resp["result"]["clean_report"]["invalid_time_rows"], 1)
        self.assertIn("regression_slope", resp["result"])
        self.assertIn("moving_average", resp["result"])

    def test_F04b_trend_no_clean(self):
        """F04b 趋势分析：关闭自动清洗，应报 DATA_TYPE_ERR"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--time_col", "日期",
            "--value_col", "销售额",
            "--no_clean",
        ])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "DATA_TYPE_ERR")

    def test_F05_anomaly(self):
        """F05 异常检测（含 Z-Score 计数）"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "anomaly"])
        self.assertTrue(resp["success"])
        self.assertIn("duplicate_row_count", resp["result"])
        outlier_stats = resp["result"]["outlier_statistics"]
        if outlier_stats:
            any_col = next(iter(outlier_stats.values()))
            self.assertIn("zscore_outlier_count", any_col)


if __name__ == "__main__":
    unittest.main()