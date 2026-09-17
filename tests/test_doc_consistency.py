"""
test_doc_consistency.py
文档一致性契约测试

覆盖目标 —— 防止历史故障回归：
「改完代码/文档后，文档副本与汇总文档静默过期」：
  · docs/SKILL.md 与 Nanobot 技能目录下的 SKILL.md 内容分叉
  · docs/skill.py 与技能目录下的 skill.py 内容分叉
  · README.md / SKILL.md 声明的用例数、测试文件清单与实际不符
  · README.md 的「目录结构」树与实际文件树不符
  · 文档里写的 CLI 参数与 cli.py 的 argparse 定义不对应

设计原则：所有断言都从**文件实际内容**推导（哈希、正则计数、argparse 定义、目录遍历），
不依赖人工同步的自觉性；文档过期时本测试直接失败。

执行：python -m unittest tests.test_doc_consistency -v
"""
import os
import re
import unittest

PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
TESTS_DIR = os.path.join(PROJECT_ROOT, "tests")
DOC_SKILL_MD = os.path.join(PROJECT_ROOT, "docs", "SKILL.md")
DOC_SKILL_PY = os.path.join(PROJECT_ROOT, "docs", "skill.py")
README = os.path.join(PROJECT_ROOT, "README.md")
CLI_PY = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")

# Nanobot 技能目录下的部署副本（工作区布局不同时跳过，不误报）
SKILLS_DIR = os.path.join(os.path.dirname(PROJECT_ROOT), "skills", "excel_analyzer_skill")

# 用例汇总句的唯一格式：测试用例共 **N 个**（test_x 11 + test_y 12）
SUMMARY_RE = re.compile(r"测试用例共\s*\*\*(\d+)\s*个\*\*\s*（(.*?)）")
PART_RE = re.compile(r"(test_[a-z_]+)\s+(\d+)")
FLAG_RE = re.compile(r"--([a-z][a-z_\-]*)")
TEST_DEF_RE = re.compile(r"^\s*def test_", re.M)


def _read(path):
    with open(path, encoding="utf-8") as f:
        return f.read()


def _read_bytes(path):
    with open(path, "rb") as f:
        return f.read()


def _actual_test_counts():
    """实际用例数：按 tests/test_*.py 中的 test_ 方法定义逐个统计"""
    counts = {}
    for name in sorted(os.listdir(TESTS_DIR)):
        if name.startswith("test_") and name.endswith(".py"):
            counts[name[:-3]] = len(TEST_DEF_RE.findall(_read(os.path.join(TESTS_DIR, name))))
    return counts


def _cli_flags():
    """cli.py 中真实存在的长参数名（下划线归一，便于与文档对比）"""
    raw = re.findall(r'add_argument\("--([a-z][a-z_\-]*)"', _read(CLI_PY))
    return {f.replace("-", "_") for f in raw}


def _norm_flags(text):
    return {f.replace("-", "_") for f in FLAG_RE.findall(text)}


def _readme_tree_stems():
    """README 「目录结构」代码块中列出的全部 *.py 文件名（去后缀）"""
    block = re.search(r"```text\n(.*?)```", _read(README), re.S).group(1)
    return set(re.findall(r"([A-Za-z_][\w\-]*)\.py", block))


def _disk_py_stems():
    """仓库中实际存在的全部 *.py 文件名（去后缀，排除运行产物目录）"""
    skip = {".git", "__pycache__", "logs", "outputs", "config", ".vscode"}
    stems = set()
    for dirpath, dirnames, filenames in os.walk(PROJECT_ROOT):
        dirnames[:] = [d for d in dirnames if d not in skip]
        for fn in filenames:
            if fn.endswith(".py"):
                stems.add(fn[:-3])
    return stems


class TestDocCopiesStayInSync(unittest.TestCase):
    """docs/ 下的源文档与技能目录下的部署副本必须逐字节一致"""

    def _assert_identical(self, doc_path, deployed_name):
        deployed = os.path.join(SKILLS_DIR, deployed_name)
        if not os.path.exists(deployed):
            self.skipTest(f"未找到部署副本 {deployed}，跳过")
        self.assertEqual(
            _read_bytes(doc_path), _read_bytes(deployed),
            msg=f"{deployed} 与 docs/{deployed_name} 内容分叉：改完 docs/ 后必须同步到技能目录（副本必须逐字节一致）",
        )

    def test_D01_skill_md_copies_identical(self):
        self._assert_identical(DOC_SKILL_MD, "SKILL.md")

    def test_D02_skill_py_copies_identical(self):
        self._assert_identical(DOC_SKILL_PY, "skill.py")


class TestDocMatchesReality(unittest.TestCase):
    """文档声明的清单/计数/参数必须与仓库实际状态一致"""

    def test_D03_doc_lists_every_test_module(self):
        actual = set(_actual_test_counts())
        self.assertGreater(len(actual), 0, msg="未发现任何 tests/test_*.py")
        for doc_path in (DOC_SKILL_MD, README):
            doc = _read(doc_path)
            listed = set(re.findall(r"tests\.(test_[a-z_]+)", doc))
            listed |= {m for m in re.findall(r"(?<![\w\-])(test_[a-z_]+)\.py", doc)}
            with self.subTest(doc=os.path.basename(doc_path)):
                self.assertEqual(
                    sorted(actual - listed), [],
                    msg=f"{os.path.basename(doc_path)} 漏列了测试文件",
                )
                self.assertEqual(
                    sorted(listed - actual), [],
                    msg=f"{os.path.basename(doc_path)} 列出了不存在的测试文件",
                )

    def test_D04_doc_case_counts_match_actual(self):
        actual = _actual_test_counts()
        total = sum(actual.values())
        for doc_path in (DOC_SKILL_MD, README):
            doc = _read(doc_path)
            m = SUMMARY_RE.search(doc)
            self.assertIsNotNone(
                m,
                msg=f"{os.path.basename(doc_path)} 缺少可解析的汇总句："
                    "测试用例共 **N 个**（test_xxx 11 + test_yyy 12 …）",
            )
            declared_total = int(m.group(1))
            declared_parts = {k: int(v) for k, v in PART_RE.findall(m.group(2))}
            with self.subTest(doc=os.path.basename(doc_path)):
                self.assertEqual(declared_total, total,
                                 msg=f"{os.path.basename(doc_path)} 声明的总用例数({declared_total}) ≠ 实际({total})")
                self.assertEqual(declared_parts, actual,
                                 msg=f"{os.path.basename(doc_path)} 的逐文件用例数与实际不符")

    def test_D05_doc_cli_flags_match_argparse(self):
        flags = _cli_flags()
        self.assertGreater(len(flags), 0, msg="未能从 cli.py 解析出任何参数")
        skill_flags = _norm_flags(_read(DOC_SKILL_MD))
        self.assertEqual(sorted(skill_flags - flags), [],
                         msg="SKILL.md 描述了 cli.py 中不存在的参数")
        self.assertEqual(sorted(flags - skill_flags), [],
                         msg="cli.py 新增/改名了参数，但 SKILL.md 未同步")
        readme_flags = _norm_flags(_read(README))
        self.assertEqual(sorted(readme_flags - flags), [],
                         msg="README.md 描述了 cli.py 中不存在的参数")

    def test_D06_readme_directory_tree_matches_disk(self):
        """README 的「目录结构」树必须与实际文件树对得上（不漏列、不写幽灵文件）"""
        tree = _readme_tree_stems()
        disk = _disk_py_stems()
        self.assertGreater(len(tree), 0, msg="未能从 README 解析出目录结构树")
        self.assertEqual(sorted(tree - disk), [],
                         msg="README 目录结构里列出了磁盘上不存在的 .py 文件")
        for rel_dir in ("src/excel_analyzer", "tests"):
            actual_stems = {
                f[:-3] for f in os.listdir(os.path.join(PROJECT_ROOT, rel_dir))
                if f.endswith(".py")
            }
            with self.subTest(dir=rel_dir):
                self.assertEqual(sorted(actual_stems - tree), [],
                                 msg=f"README 目录结构漏列了 {rel_dir} 下的模块")


if __name__ == "__main__":
    unittest.main(verbosity=2)
