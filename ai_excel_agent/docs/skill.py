"""
nanobot skill插件: excel_analyzer_skill
适配 ai‑excel‑agent，遵循 docs/SKILL.md 安全契约
存放位置：项目 docs/skill.py（样例）
使用时：将 skill.py + SKILL.md 复制到 Nanobot 的自建技能目录
注意：不要放入 ai_excel_agent/src 源码目录
"""
import os
import json
import subprocess
from pathlib import Path
from typing import Dict, Any, Optional

# --------自动推导ai_excel_agent项目根目录，消除硬编码绝对路径--------
# 策略：向上搜索，找到包含 src/excel_analyzer 的目录
def _detect_agent_project_root() -> Path:
    """自动探测ai_excel_agent项目根目录，向上最多回溯6层"""
    current = Path(__file__).resolve()
    for _ in range(6):
        candidate = current / "src" / "excel_analyzer"
        if candidate.is_dir():
            return current
        current = current.parent
    # 探测失败回退：使用环境变量 AGENT_PROJECT_ROOT
    env_root = os.getenv("AGENT_PROJECT_ROOT")
    if env_root:
        return Path(env_root).resolve()
    raise RuntimeError(
        "无法自动探测ai_excel_agent项目根目录。"
        "请设置环境变量 AGENT_PROJECT_ROOT 指向项目根目录。"
    )

try:
    PROJECT_ROOT = _detect_agent_project_root()
except RuntimeError as e:
    raise SystemExit(e)

# 导入底层白名单确认API
import sys
sys.path.insert(0, str(PROJECT_ROOT))
from src.excel_analyzer.whitelist_manager import agent_apply_confirm

# ----------------------------------------------------------------

def excel_analyze(
    file_path: str,
    intent: str,
    sheet=0,
    time_col: Optional[str] = None,
    value_col: Optional[str] = None,
    sort_col: Optional[str] = None,
    sort_asc: bool = False,
    filter_condition: Optional[str] = None,
    session_ctx: Optional[Dict[str, Any]] = None
):
    """
    读取本地Excel文件进行数据分析。
    支持intent枚举：
    overview：数据概览；
    stats：数值描述统计；
    sort_filter：排序过滤；
    trend：时序趋势（time_col、value_col必须同时传入）；
    anomaly：异常检测。

    【安全契约说明】
    白名单校验逻辑下沉至cli/whitelist_manager；
    文件不在信任目录会返回 NEED_USER_CONFIRM，需要用户在对话中确认是否添加目录；
    禁止绕过底层白名单机制自行实现权限拦截。
    【强制约束】：所有Excel分析任务仅调用本工具，禁止调用shell/python/file‑read等其他工具读取Excel。

    :param session_ctx: Nanobot会话上下文，用于保存待确认任务状态
    """
    if session_ctx is None:
        session_ctx = {}

    env = os.environ.copy()
    env["PYTHONPATH"] = str(PROJECT_ROOT)
    cli_path = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")

    cmd = [
        "python",
        cli_path,
        "--file", file_path,
        "--intent", intent,
        "--sheet", str(sheet)
    ]
    if time_col is not None:
        cmd.extend(["--time_col", time_col])
    if value_col is not None:
        cmd.extend(["--value_col", value_col])
    if sort_col is not None:
        cmd.extend(["--sort_col", sort_col])
    if sort_asc:
        cmd.append("--sort_asc")
    if filter_condition is not None:
        cmd.extend(["--filter_condition", filter_condition])

    proc = subprocess.run(cmd, capture_output=True, text=True, env=env)
    try:
        resp = json.loads(proc.stdout)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error_code": "JSON_PARSE_ERR",
            "error_msg": f"JSON解析失败:{str(e)}",
            "stdout_raw": proc.stdout,
            "result": {}
        }

    # 处理需要用户确认白名单场景：把待确认信息存入会话上下文
    if resp.get("error_code") == "NEED_USER_CONFIRM":
        confirm_info = resp["result"]["confirm_info"]
        session_ctx["_excel_pending_confirm"] = {
            "candidate_dir": confirm_info["candidate_dir"],
            "target_file": confirm_info["file_abs"],
            "origin_args": {
                "file_path": file_path,
                "intent": intent,
                "sheet": sheet,
                "time_col": time_col,
                "value_col": value_col,
                "sort_col": sort_col,
                "sort_asc": sort_asc,
                "filter_condition": filter_condition
            }
        }
        return {
            "success": False,
            "error_code": "NEED_USER_CONFIRM",
            "error_msg": (
                f"文件 {confirm_info['file_abs']} 不在信任白名单。\n"
                f"请确认是否信任目录【{confirm_info['candidate_dir']}】，回复【是】或【否】"
            ),
            "result": {}
        }

    return resp


def excel_handle_user_confirm(user_input: str, session_ctx: Dict[str, Any]) -> Dict[str, Any]:
    """
    处理用户回复【是/否】，完成白名单确认；确认通过后重新执行原先分析任务
    由Nanobot上层逻辑检测到session_ctx存在待确认任务时调用
    """
    pending = session_ctx.pop("_excel_pending_confirm", None)
    if pending is None:
        return {"success": False, "error_msg": "无待确认的Excel访问任务", "result": {}}

    user_approve = user_input.strip() == "是"
    agent_apply_confirm(pending["candidate_dir"], user_approve)

    if not user_approve:
        return {"success": False, "error_msg": "用户拒绝信任目录，本次分析终止。", "result": {}}

    # 用户同意，重新执行原始分析参数
    args = pending["origin_args"]
    return excel_analyze(
        file_path=args["file_path"],
        intent=args["intent"],
        sheet=args["sheet"],
        time_col=args["time_col"],
        value_col=args["value_col"],
        sort_col=args["sort_col"],
        sort_asc=args["sort_asc"],
        filter_condition=args["filter_condition"],
        session_ctx=session_ctx
    )


# nanobot skill 导出声明：暴露工具给框架发现
SKILL_TOOLS = [
    {
        "name": "excel_analyze",
        "callable": excel_analyze,
        "description": excel_analyze.__doc__,
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel本地文件路径"},
                "intent": {
                    "type": "string",
                    "enum": ["overview", "stats", "sort_filter", "trend", "anomaly"],
                    "description": "分析意图"
                },
                "sheet": {"type": ["integer", "string"], "description": "工作表索引(0起始)或工作表名称，默认0"},
                "time_col": {"type": "string", "description": "trend专用：时间列名称"},
                "value_col": {"type": "string", "description": "trend专用：数值列名称"},
                "sort_col": {"type": "string", "description": "sort_filter专用：待排序列名"},
                "sort_asc": {"type": "boolean", "description": "sort_filter，True升序，False降序"},
                "filter_condition": {"type": "string", "description": "sort_filter过滤条件，示例：销售额>500"}
            },
            "required": ["file_path", "intent"]
        }
    }
]

SKILL_METADATA = {
    "name": "Excel‑Analyzer‑Skill",
    "version": "1.0.0",
    "description": "AI‑Excel‑Agent项目Nanobot Skill，调用独立CLI完成Excel数据分析；白名单安全校验下沉至底层cli",
    "readme": os.path.join(os.path.dirname(__file__), "SKILL.md")
}