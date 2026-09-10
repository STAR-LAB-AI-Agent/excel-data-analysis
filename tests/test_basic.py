"""
test_basic.py
AI‑Excel‑Agent 测试用例
执行方式：在项目根目录运行 python -m unittest tests.test_basic -v
"""
import unittest
import subprocess
import json
import os

# =========新增开始========
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from src.excel_analyzer.whitelist_manager import load_whitelist, PROJECT_ROOT

TEST_DATA_DIR = str((PROJECT_ROOT / "tests" / "test_data").resolve())

# 单元测试前置校验：确认测试目录在白名单中
_whitelist = load_whitelist()
if TEST_DATA_DIR not in _whitelist:
    raise RuntimeError(
        f"【单元测试前置检查失败】\n"
        f"测试目录 {TEST_DATA_DIR} 不在白名单列表 {_whitelist}\n"
        f"删除 config/excel_agent_config.json，让程序自动生成默认白名单后再运行单元测试。"
    )
# =========新增结束========

# 项目根目录
PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
# 测试数据目录
TEST_DATA_DIR = os.path.join(PROJECT_ROOT, "tests", "test_data")

# 设置环境变量，保证src包可被找到
env = os.environ.copy()
env["PYTHONPATH"] = PROJECT_ROOT


def run_cli(args_list: list[str]) -> dict:
    """封装调用cli，返回解析后的response字典
    自动过滤混杂在stdout里的日志输出，提取JSON片段，不再修改logger配置
    """
    cmd = ["python", CLI_SCRIPT] + args_list
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=False,
        env=env
    )
    # 多编码兼容解码
    try:
        stdout_text = result.stdout.decode("utf-8")
    except UnicodeDecodeError:
        stdout_text = result.stdout.decode("gbk", errors="replace")

    # 从输出中找到第一个'{'，截取后面全部，过滤掉前面的日志打印行
    json_start = stdout_text.find("{")
    if json_start == -1:
        raise RuntimeError(f"输出中未找到JSON对象，stdout_text:\n{stdout_text}")
    json_text = stdout_text[json_start:]
    return json.loads(json_text)


class TestCliErrorCases(unittest.TestCase):
    """异常用例 E01‑E04，不需要真实Excel文件"""

    # E01 文件不存在
    def test_E01_file_not_found(self):
        # 原："no_file.xlsx"
        resp = run_cli(["--file", "tests/test_data/no_file.xlsx", "--intent", "overview"])
        self.assertEqual(resp["error_code"], "FILE_NOT_FOUND")

    # E02 非法意图
    def test_E02_invalid_intent(self):
        # 原："no_file.xlsx"
        resp = run_cli(["--file", "tests/test_data/no_file.xlsx", "--intent", "wrong_intent"])
        self.assertEqual(resp["error_code"], "INVALID_INTENT")

    # E03 trend缺少必填参数
    def test_E03_trend_missing_arg(self):
        # 原："no_file.xlsx"
        resp = run_cli(["--file", "tests/test_data/no_file.xlsx", "--intent", "trend"])
        self.assertEqual(resp["error_code"], "MISSING_ARG")

    def test_E04_invalid_suffix(self):
        """E04 文件后缀非法，传入真实存在的txt文件"""
        bad_file = os.path.join(TEST_DATA_DIR, "bad_suffix.txt")
        resp = run_cli(["--file", bad_file, "--intent", "overview"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "INVALID_SUFFIX")


#@unittest.skip("需要准备测试Excel后，注释掉skip执行下面功能测试")
class TestCliFunctionCases(unittest.TestCase):
    """功能测试 F01‑F05 + E05；需要你把测试excel放入tests/test_data"""
    TEST_EXCEL = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

    def test_E05_sheet_not_exist(self):
        """E05 指定不存在sheet"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "overview", "--sheet", "NotExistSheet"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "SHEET_NOT_EXIST")

    def test_F01_overview(self):
        """F01 数据概览"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "overview"])
        self.assertTrue(resp["success"])
        self.assertIn("row_count", resp["result"])

    def test_F02_stats(self):
        """F02 描述性统计"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "stats"])
        self.assertTrue(resp["success"])

    def test_F03_sort_filter(self):
        """F03 排序筛选"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "sort_filter",
            "--sort_col", "销售额",
            "--filter_condition", "销售额>500"
        ])
        self.assertTrue(resp["success"])
        self.assertIn("hit_rows", resp["result"])

    def test_F04_trend(self):
        """F04 趋势分析：样本包含非法时间行，预期触发时间解析DataTypeError"""
        resp = run_cli([
            "--file", self.TEST_EXCEL,
            "--intent", "trend",
            "--time_col", "日期",
            "--value_col", "销售额"
        ])
        # 当前测试数据含有非法日期，应当返回DataTypeError
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "DATA_TYPE_ERR")

    def test_F05_anomaly(self):
        """F05 异常检测"""
        resp = run_cli(["--file", self.TEST_EXCEL, "--intent", "anomaly"])
        self.assertTrue(resp["success"])
        self.assertIn("duplicate_row_count", resp["result"])


if __name__ == "__main__":
    unittest.main()