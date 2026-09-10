# tests/debug_cli.py
"""调试脚本：直接调用 CLI 并打印原始 stdout/stderr"""

import subprocess
import os
import sys

# 项目根目录
PROJECT_ROOT = os.path.abspath(os.path.dirname(os.path.dirname(__file__)))
CLI_SCRIPT = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")
TEST_EXCEL = os.path.join(PROJECT_ROOT, "tests", "test_data", "sample_test.xlsx")

# 方案A：继承系统环境变量
env_a = os.environ.copy()
env_a["PYTHONPATH"] = PROJECT_ROOT

# 方案B：最小环境变量
env_b = {
    "PYTHONPATH": PROJECT_ROOT,
    "PATH": os.environ.get("PATH", ""),
    "SYSTEMROOT": os.environ.get("SYSTEMROOT", ""),
}

def test_run(env, desc):
    print(f"\n{'='*60}")
    print(f"测试方案: {desc}")
    print(f"环境变量: {list(env.keys())}")
    print(f"命令: python {CLI_SCRIPT} --file {TEST_EXCEL} --intent overview")
    print('-'*60)
    
    cmd = ["python", CLI_SCRIPT, "--file", TEST_EXCEL, "--intent", "overview"]
    
    # 使用 text=True 让 subprocess 自动处理编码
    result = subprocess.run(
        cmd,
        capture_output=True,
        text=True,
        env=env,
        encoding="utf-8",
        errors="replace"
    )
    
    print(f"返回码 (returncode): {result.returncode}")
    print(f"stdout 长度: {len(result.stdout)}")
    print(f"stdout 内容 (repr): {repr(result.stdout)}")
    print(f"stdout 前200字符: {result.stdout[:200] if result.stdout else '(空)'}")
    print(f"\nstderr 长度: {len(result.stderr)}")
    print(f"stderr 内容 (repr): {repr(result.stderr)}")
    print(f"stderr 前200字符: {result.stderr[:200] if result.stderr else '(空)'}")
    
    # 如果 stdout 有内容，尝试解析 JSON
    if result.stdout.strip():
        try:
            import json
            json_start = result.stdout.find("{")
            if json_start != -1:
                data = json.loads(result.stdout[json_start:])
                print(f"\n✅ JSON 解析成功! success={data.get('success')}")
            else:
                print("\n⚠️ stdout 中没有找到 '{'")
        except json.JSONDecodeError as e:
            print(f"\n❌ JSON 解析失败: {e}")
    
    return result

if __name__ == "__main__":
    # 检查文件是否存在
    print(f"CLI 脚本存在: {os.path.exists(CLI_SCRIPT)}")
    print(f"测试 Excel 存在: {os.path.exists(TEST_EXCEL)}")
    
    # 测试方案A
    result_a = test_run(env_a, "继承系统环境变量")
    
    # 测试方案B
    result_b = test_run(env_b, "最小环境变量 + PATH")
    
    # 测试方案C：直接在当前进程中导入并执行（绕过 subprocess）
    print(f"\n{'='*60}")
    print("测试方案C: 直接导入 CLI 模块执行")
    print('-'*60)
    try:
        sys.path.insert(0, PROJECT_ROOT)
        # 模拟命令行参数
        sys.argv = [
            "cli.py",
            "--file", TEST_EXCEL,
            "--intent", "overview"
        ]
        # 导入并执行 main
        from src.excel_analyzer import cli
        # 注意：需要捕获 print 输出
        import io
        from contextlib import redirect_stdout
        
        f = io.StringIO()
        with redirect_stdout(f):
            cli.main()
        output = f.getvalue()
        print(f"捕获到输出长度: {len(output)}")
        print(f"输出内容 (repr): {repr(output[:500])}")
        if output.strip():
            json_start = output.find("{")
            if json_start != -1:
                import json
                data = json.loads(output[json_start:])
                print(f"✅ JSON 解析成功! success={data.get('success')}")
            else:
                print("⚠️ 输出中没有找到 '{'")
    except Exception as e:
        print(f"❌ 直接导入执行失败: {e}")
        import traceback
        traceback.print_exc()