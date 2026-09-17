"""
test_encoding.py
输出编码契约回归测试

覆盖目标 —— 防止历史故障回归：
「Windows 控制台/管道默认代码页为 GBK 时，CLI 的中文报告按 GBK 写出，
  上层按 UTF-8 解析后整段乱码；只能靠智能体临场想到『把 stdout 改写进 UTF-8 文件再读』
  这类额外步骤才拿到可用结果」

契约（详见 SKILL.md §5）：cli.py 启动时把 stdout/stderr 固定为 UTF-8，
无论平台代码页与 PYTHONIOENCODING 如何设置，stdout 字节流恒为 UTF-8。

执行：python -m unittest tests.test_encoding -v
"""
import os
import subprocess
import unittest

PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
SAMPLE = os.path.join(PROJECT_ROOT, "tests", "test_data", "sample_test.xlsx")


def _env(extra=None, drop=None):
    env = os.environ.copy()
    env["PYTHONPATH"] = PROJECT_ROOT
    for key in (drop or []):
        env.pop(key, None)
    env.update(extra or {})
    return env


def _run_bytes(args, env):
    """按字节捕获 stdout：测试中故意不做任何编码兜底，否则会掩盖乱码故障"""
    return subprocess.run(["python", CLI_SCRIPT] + args,
                          capture_output=True, text=False, env=env)


class TestStdoutIsAlwaysUtf8(unittest.TestCase):
    """CLI 的 stdout 必须与平台代码页无关地恒为 UTF-8"""

    def _assert_utf8_cjk(self, raw: bytes, label: str, tokens):
        try:
            text = raw.decode("utf-8")  # 严格解码：不是 UTF-8 直接失败
        except UnicodeDecodeError as e:
            self.fail(f"{label}：stdout 不是合法 UTF-8（{e}）；前 80 字节 = {raw[:80]!r}")
        for token in tokens:
            self.assertIn(token, text, msg=f"{label}：中文未正确还原，缺少 {token}")
        return text

    def test_E01_md_under_legacy_console(self):
        """模拟 GBK 控制台（PYTHONIOENCODING=gbk）：md 正文仍必须是 UTF-8"""
        proc = _run_bytes(["--file", SAMPLE, "--intent", "report", "--format", "md"],
                          _env({"PYTHONIOENCODING": "gbk"}))
        self.assertEqual(proc.returncode, 0, msg=proc.stderr[:500])
        text = self._assert_utf8_cjk(proc.stdout, "report --format md", ["销售额", "区域", "备注"])
        self.assertIn("## anomaly", text)

    def test_E02_json_under_legacy_console(self):
        """JSON 模式同样锁定 UTF-8（中文列名/结论可读）"""
        proc = _run_bytes(["--file", SAMPLE, "--intent", "report"],
                          _env({"PYTHONIOENCODING": "gbk"}))
        self.assertEqual(proc.returncode, 0, msg=proc.stderr[:500])
        text = self._assert_utf8_cjk(proc.stdout, "report json", ["销售额", "区域"])
        self.assertIn('"success": true', text)

    def test_E03_default_env_is_utf8(self):
        """不设任何编码变量（父进程 locale 可能是 cp936）时也必须是 UTF-8"""
        proc = _run_bytes(["--file", SAMPLE, "--intent", "overview"],
                          _env(drop=["PYTHONIOENCODING"]))
        self.assertEqual(proc.returncode, 0, msg=proc.stderr[:500])
        self._assert_utf8_cjk(proc.stdout, "overview 默认环境", ["销售额", "产品名称"])

    def test_E04_error_path_is_utf8_too(self):
        """错误分支（如文件不存在）的中文提示同样不能乱码"""
        proc = _run_bytes(["--file", "no_such_file.xlsx", "--intent", "overview"],
                          _env({"PYTHONIOENCODING": "gbk"}))
        text = self._assert_utf8_cjk(proc.stdout, "FILE_NOT_FOUND", ["文件不存在"])
        self.assertIn("FILE_NOT_FOUND", text)


class TestCallerSideContract(unittest.TestCase):
    """适配层 skill.py 必须显式按 UTF-8 解码，不得依赖平台 locale"""

    SKILL_CANDIDATES = [
        os.path.join(PROJECT_ROOT, "docs", "skill.py"),
        os.path.join(os.path.dirname(PROJECT_ROOT), "skills", "excel_analyzer_skill", "skill.py"),
    ]

    def test_E05_skill_declares_utf8_decoding(self):
        checked = 0
        for path in self.SKILL_CANDIDATES:
            if not os.path.exists(path):
                continue
            checked += 1
            with open(path, encoding="utf-8") as f:
                src = f.read()
            with self.subTest(file=path):
                self.assertIn('encoding="utf-8"', src)
                self.assertEqual(src.count("PYTHONIOENCODING"), 2,
                                 msg="两个 subprocess 入口都应设置 PYTHONIOENCODING")
                self.assertEqual(src.count('errors="replace"'), 2,
                                 msg="两个 subprocess.run 都应显式 encoding/errors")
                self.assertEqual(src.count("subprocess.run(cmd, capture_output=True, text=True,"), 2,
                                 msg="subprocess.run 调用点数量异常")
        self.assertGreater(checked, 0, msg="未找到任何 skill.py 适配层文件")


if __name__ == "__main__":
    unittest.main(verbosity=2)
