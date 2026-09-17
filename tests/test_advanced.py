"""
test_advanced.py
进阶测试：图表生成、自动清洗、算法增强
执行：python -m unittest tests.test_advanced -v
"""
import unittest
import subprocess
import json
import os
from pathlib import Path

PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
TEST_DATA_DIR = os.path.join(PROJECT_ROOT, "tests", "test_data")
OUTPUT_DIR = os.path.join(PROJECT_ROOT, "outputs")

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


class TestChartGeneration(unittest.TestCase):
    """图表生成测试（可选功能）"""

    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_C01_bar_chart(self):
        """C01 柱状图：按区域统计销售额均值"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "bar",
            "--x_col", "区域",
            "--y_col", "销售额",
            "--title", "各区域销售额均值",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        out_path = resp["result"]["output_path"]
        self.assertTrue(os.path.exists(out_path))
        self.assertGreater(resp["result"]["file_size_kb"], 0)

    def test_C02_line_chart(self):
        """C02 折线图：销售额随日期变化"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "line",
            "--x_col", "日期",
            "--y_col", "销售额",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        self.assertTrue(os.path.exists(resp["result"]["output_path"]))

    def test_C03_box_chart(self):
        """C03 箱线图：多数值列"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "box",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        self.assertTrue(os.path.exists(resp["result"]["output_path"]))

    def test_C04_hist_chart(self):
        """C04 直方图：销售额分布"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "hist",
            "--y_col", "销售额",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        self.assertTrue(os.path.exists(resp["result"]["output_path"]))

    def test_C05_chart_missing_arg(self):
        """C05 缺少参数：bar 图未提供 y_col"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "bar",
            "--x_col", "区域",
        ])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "MISSING_ARG")

    def test_C06_chart_col_not_found(self):
        """C06 列不存在"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "chart",
            "--chart_type", "hist",
            "--y_col", "不存在的列",
        ])
        self.assertFalse(resp["success"])
        self.assertIn(resp["error_code"], ("COLUMN_NOT_FOUND", "NO_NUMERIC_COL"))


class TestDataClean(unittest.TestCase):
    """数据清洗测试"""

    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_D01_trend_autoclean_multi_bad(self):
        """D01 多行脏时间：auto_clean 应剔除全部脏行"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--sheet", "脏时间数据",
            "--time_col", "日期",
            "--value_col", "指标",
        ])
        self.assertTrue(resp["success"], msg=str(resp))
        clean = resp["result"]["clean_report"]
        self.assertEqual(clean["invalid_time_rows"], 2)
        self.assertTrue(clean["dropped_invalid"])
        self.assertEqual(clean["remaining_rows"], 4)

    def test_D02_trend_no_clean_multi_bad(self):
        """D02 多行脏时间，关闭清洗，应报错"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--sheet", "脏时间数据",
            "--time_col", "日期",
            "--value_col", "指标",
            "--no_clean",
        ])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "DATA_TYPE_ERR")


class TestAlgorithmEnhancements(unittest.TestCase):
    """算法增强测试"""

    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_A01_stats_cv_skew(self):
        """A01 统计含变异系数与偏度"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "stats"])
        self.assertTrue(resp["success"])
        sr = resp["result"]["stats_result"]
        self.assertIn("销售额", sr)
        self.assertIn("cv", sr["销售额"])
        self.assertIn("skew", sr["销售额"])

    def test_A02_trend_regression(self):
        """A02 趋势含回归斜率与移动平均"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--time_col", "日期",
            "--value_col", "销售额",
        ])
        self.assertTrue(resp["success"])
        self.assertIn("regression_slope", resp["result"])
        self.assertIn("moving_average", resp["result"])
        self.assertIsInstance(resp["result"]["moving_average"], list)

    def test_A03_anomaly_zscore(self):
        """A03 异常检测含 Z-Score 计数"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "anomaly"])
        self.assertTrue(resp["success"])
        os_ = resp["result"]["outlier_statistics"]
        self.assertIn("销售额", os_)
        self.assertIn("zscore_outlier_count", os_["销售额"])

    def test_A04_stats_no_numeric(self):
        """A04 纯文本sheet统计应返回提示"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "stats",
            "--sheet", "员工信息",
        ])
        self.assertTrue(resp["success"])
        self.assertEqual(resp["result"]["numeric_columns"], [])
        self.assertIn("hint", resp["result"])


if __name__ == "__main__":
    unittest.main()