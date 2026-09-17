"""
test_report.py
一次性完整报告（intent=report）专项测试

覆盖目标 —— 防止历史故障回归：
「六个意图需要多次调用 + 智能体自行拼装 + 图表单独成气泡」
导致「只发了一张图、其余问题没有回复」的交付失败。

执行：python -m unittest tests.test_report -v
"""
import unittest
import subprocess
import json
import os
import shutil
import tempfile

PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
TEST_DATA_DIR = os.path.join(PROJECT_ROOT, "tests", "test_data")
SAMPLE = os.path.join(TEST_DATA_DIR, "sample_test.xlsx")

# SKILL.md §9 要求的分段标题（缺一即视为交付不完整）
REQUIRED_SECTIONS = [
    "## overview",
    "## stats",
    "## sort_filter",
    "## trend",
    "## anomaly",
    "## chart",
]

# 分段子集（--intents）专项：未请求的分段不得出现在正文里
ALL_SECTIONS = ("overview", "stats", "sort_filter", "trend", "anomaly", "chart")

# anomaly 段强制四要素
ANOMALY_REQUIRED = ["重复行数", "高缺失列", "IQR 离群数", "Z-Score 离群数"]

env = os.environ.copy()
env["PYTHONPATH"] = PROJECT_ROOT


def _raw_stdout(args_list):
    """按字节捕获 stdout 后严格按 UTF-8 解码。

    cli.py 已把 stdout 固定为 UTF-8（见 SKILL.md §5 输出编码契约），
    因此这里不再需要 gbk 兼容分支；保留兼容分支会掩盖编码回归故障。
    """
    result = subprocess.run(["python", CLI_SCRIPT] + args_list,
                            capture_output=True, text=False, env=env)
    try:
        return result.stdout.decode("utf-8"), result.returncode
    except UnicodeDecodeError as e:
        raise AssertionError(
            f"stdout 不是合法 UTF-8（{e}）：编码回归！前 80 字节 = {result.stdout[:80]!r}"
        )


def run_cli(args_list):
    text, _ = _raw_stdout(args_list)
    idx = text.find("{")
    if idx == -1:
        raise RuntimeError(f"输出中未找到JSON对象：\n{text}")
    return json.loads(text[idx:])


class TestReportOneShot(unittest.TestCase):
    """report 意图：一次调用必须产出完整交付物"""

    @classmethod
    def setUpClass(cls):
        cls.resp = run_cli(["--file", SAMPLE, "--intent", "report"])
        cls.res = cls.resp.get("result", {})
        cls.md = cls.res.get("markdown", "")

    # ---------- 基础通路 ----------
    def test_R01_success(self):
        self.assertTrue(self.resp["success"], msg=str(self.resp)[:500])

    def test_R02_all_six_sections_present(self):
        """六个分段必须全部出现在同一份报告里"""
        for header in REQUIRED_SECTIONS:
            with self.subTest(section=header):
                self.assertIn(header, self.md)

    def test_R03_no_section_failed(self):
        """任何一段都不应带 __error__"""
        sections = self.res.get("sections", {})
        self.assertEqual(len(sections), 6, msg=f"分段数不足：{list(sections)}")
        for name, body in sections.items():
            with self.subTest(section=name):
                self.assertNotIn("__error__", body, msg=f"{name} 段执行失败：{body}")

    # ---------- §9 强制内容 ----------
    def test_R04_anomaly_four_elements(self):
        """anomaly 段必须包含重复行/高缺失列/IQR/Z-Score 四要素"""
        for token in ANOMALY_REQUIRED:
            with self.subTest(token=token):
                self.assertIn(token, self.md)

    def test_R05_chart_output_path_and_numbers(self):
        """chart 段必须给出 output_path，且报告内要有分组数值（不能只有图）"""
        chart = self.res["sections"]["chart"]
        self.assertIn("output_path", chart)
        self.assertTrue(os.path.exists(chart["output_path"]),
                        msg=f"图表文件不存在：{chart['output_path']}")
        self.assertIn("output_path", self.md)
        self.assertIn("分组数值", self.md)

    def test_R06_attachments_exposed(self):
        """附件路径必须显式给出，供智能体做附件交付"""
        self.assertEqual(len(self.res.get("attachments", [])), 1)
        self.assertTrue(os.path.exists(self.res["attachments"][0]))

    def test_R07_delivery_checklist_requires_single_message(self):
        """交付清单必须明确要求「正文与附件同一条消息」"""
        checklist = "\n".join(self.res.get("delivery_checklist", []))
        self.assertIn("同一条消息", checklist)
        self.assertIn("禁止只发送图表", checklist)

    # ---------- 列自动探测 ----------
    def test_R08_autodetect_without_any_column_args(self):
        """不传任何列参数也能跑出完整报告"""
        det = self.res.get("auto_detected", {})
        self.assertEqual(det.get("time_col"), "日期")
        self.assertEqual(det.get("value_col"), "销售额")
        self.assertEqual(det.get("x_col"), "区域")

    def test_R09_report_respects_explicit_filter_and_sort(self):
        """显式过滤+升序参数要被 report 采纳"""
        resp = run_cli([
            "--file", SAMPLE, "--intent", "report",
            "--sort_col", "销售额", "--sort_asc",
            "--filter_condition", "销售额>500",
        ])
        self.assertTrue(resp["success"], msg=str(resp)[:500])
        sf = resp["result"]["sections"]["sort_filter"]
        self.assertEqual(sf["hit_rows"], 26)
        self.assertTrue(sf["sort_asc"])

    def test_R10_trend_flags_outlier_driven_change(self):
        """趋势段必须点明「表面暴跌其实由离群点驱动」"""
        self.assertIn("剔除离群点后真实走势", self.md)


class TestReportSubsetScope(unittest.TestCase):
    """report 分段子集（--intents）：用户只问其中几件事时不得把六段全倒出去

    锁定「冗余交付」回归：全量报告解决了漏答，但用户只请求 2 个意图时，
    若仍跑满六段，就是答非所问。因此要求：
        · 只计算、只渲染、只出图请求的分段；
        · 范围写进 markdown 与 delivery_checklist；
        · 未知分段名报错，不静默丢弃。
    """

    def test_R13_subset_runs_only_requested_sections(self):
        """--intents stats,trend：正文与 sections 只含这 2 段，其余不得出现"""
        resp = run_cli(["--file", SAMPLE, "--intent", "report", "--intents", "stats,trend"])
        self.assertTrue(resp["success"], msg=str(resp)[:500])
        res = resp["result"]
        md = res["markdown"]
        self.assertEqual(res["requested_intents"], ["stats", "trend"])
        self.assertEqual(res["scope"], "subset")
        self.assertEqual(sorted(res["sections"].keys()), ["stats", "trend"])
        self.assertIn("## stats", md)
        self.assertIn("## trend", md)
        for absent in ("## overview", "## sort_filter", "## anomaly", "## chart"):
            with self.subTest(absent=absent):
                self.assertNotIn(absent, md, msg=f"未请求的 {absent} 出现在正文里（冗余交付）")
        self.assertEqual(res["excluded_intents"],
                         ["overview", "sort_filter", "anomaly", "chart"])

    def test_R14_subset_without_chart_produces_no_attachment(self):
        """未请求 chart 段时不生成图片，且自查表必须声清「不会出图」"""
        resp = run_cli(["--file", SAMPLE, "--intent", "report", "--intents", "anomaly"])
        self.assertTrue(resp["success"], msg=str(resp)[:500])
        res = resp["result"]
        self.assertEqual(res["requested_intents"], ["anomaly"])
        self.assertEqual(res["attachments"], [])
        self.assertIn("## anomaly", res["markdown"])
        checklist = "\n".join(res["delivery_checklist"])
        self.assertIn("未请求 chart", checklist)
        self.assertIn("不得出现", checklist)

    def test_R15_subset_chart_keeps_output_path_and_numbers(self):
        """只请求 chart 时仍必须给 output_path 与分组数值，且不带其他分段"""
        resp = run_cli(["--file", SAMPLE, "--intent", "report", "--intents", "chart"])
        self.assertTrue(resp["success"], msg=str(resp)[:500])
        res = resp["result"]
        self.assertEqual(len(res["attachments"]), 1)
        self.assertTrue(os.path.exists(res["attachments"][0]))
        self.assertIn("## chart", res["markdown"])
        self.assertIn("output_path", res["markdown"])
        self.assertIn("分组数值", res["markdown"])
        self.assertNotIn("## overview", res["markdown"])

    def test_R16_intents_order_and_dup_normalized(self):
        """子集书写顺序无关、可重复，输出恒为规范顺序（stats 在 chart 之前）"""
        resp = run_cli(["--file", SAMPLE, "--intent", "report", "--intents", "chart,stats,stats"])
        self.assertTrue(resp["success"], msg=str(resp)[:500])
        res = resp["result"]
        self.assertEqual(res["requested_intents"], ["stats", "chart"])
        md = res["markdown"]
        self.assertLess(md.index("## stats"), md.index("## chart"))

    def test_R17_full_scope_default_and_all(self):
        """默认不传 --intents 与 --intents all 都等于全量六段（向后兼容）"""
        for extra in ([], ["--intents", "all"]):
            with self.subTest(extra=extra):
                resp = run_cli(["--file", SAMPLE, "--intent", "report"] + extra)
                self.assertTrue(resp["success"], msg=str(resp)[:500])
                res = resp["result"]
                self.assertEqual(len(res["sections"]), 6)
                self.assertEqual(res["excluded_intents"], [])
                self.assertEqual(res["scope"], "full")
                for header in REQUIRED_SECTIONS:
                    self.assertIn(header, res["markdown"])

    def test_R18_unknown_section_rejected(self):
        """未知分段名必须报错（INVALID_ARG），不得静默丢掉用户问过的分段"""
        resp = run_cli(["--file", SAMPLE, "--intent", "report", "--intents", "overview,foo"])
        self.assertFalse(resp["success"])
        self.assertEqual(resp["error_code"], "INVALID_ARG")
        self.assertIn("foo", resp["error_msg"])


class TestReportMarkdownFormat(unittest.TestCase):
    """--format md：stdout 必须是可直接交付的 Markdown 正文"""

    def test_R11_md_format_is_pure_markdown(self):
        text, code = _raw_stdout(["--file", SAMPLE, "--intent", "report", "--format", "md"])
        self.assertEqual(code, 0)
        self.assertTrue(text.lstrip().startswith("# "), msg=text[:200])
        for header in REQUIRED_SECTIONS:
            with self.subTest(section=header):
                self.assertIn(header, text)
        # 不能是 JSON 包装
        self.assertNotIn('"success"', text)

    def test_R12_md_format_keeps_output_json_dual_write(self):
        """md 模式与 --output_json 可以同时使用，互不干扰（写到系统临时目录，不污染仓库）"""
        out_dir = tempfile.mkdtemp(prefix="excel_report_test_")
        out_json = os.path.join(out_dir, "report_test_dual.json")
        try:
            text, code = _raw_stdout([
                "--file", SAMPLE, "--intent", "report",
                "--format", "md", "--output_json", out_json,
            ])
            self.assertEqual(code, 0)
            self.assertTrue(text.lstrip().startswith("# "))
            self.assertTrue(os.path.exists(out_json))
            with open(out_json, encoding="utf-8") as f:
                payload = json.load(f)
            self.assertTrue(payload["success"])
            self.assertEqual(len(payload["result"]["attachments"]), 1)
        finally:
            shutil.rmtree(out_dir, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
