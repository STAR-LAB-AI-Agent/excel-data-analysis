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
    # 子进程 stdout 已被 cli.py 固定为 UTF-8；此处再显式声明一遍，
    # 避免父进程 locale 为非 UTF-8（如中文 Windows 的 cp936）时解码错位。
    env["PYTHONIOENCODING"] = "utf-8"
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

    # 显式按 UTF-8 解码，禁止依赖平台 locale（否则中文列名/报告会乱码）
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
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


def excel_analyze_report(
    file_path: str,
    sheet=0,
    intents: Optional[str] = None,
    chart_type: str = "bar",
    time_col: Optional[str] = None,
    value_col: Optional[str] = None,
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
    sort_col: Optional[str] = None,
    sort_asc: bool = True,
    filter_condition: Optional[str] = None,
    title: Optional[str] = None,
    session_ctx: Optional[Dict[str, Any]] = None,
):
    """
    ✨ 一次性报告（**推荐入口**）。

    单次调用内跑完**请求的分段**（默认全部六段），
    直接返回一份**可直接作为回复正文的 Markdown 报告**，
    同时给出 attachments（图表路径，仅请求 chart 段时非空）与 delivery_checklist（交付自查表）。

    参数 intents：分段子集，逗号分隔，例 "stats,trend"；不传或 "all" = 全部六段。
    用户只问其中几件事时**必须传**这个参数，未请求的分段不计算、不渲染、不出图。

    返回:
        {
          "success": bool,
          "markdown": str,             # 直接作为回复正文，不要改写、不要摘要
          "attachments": [str],        # 非空时必须与 markdown 放在**同一条消息**中发出
          "delivery_checklist": [str], # 发送前逐条自检（第一条是本次范围红线）
          "sections": {...},           # 本次请求的分段的原始结构化结果
          "requested_intents": [str],  # 本次真正跑的分段
          "excluded_intents": [str],   # 未请求的分段：正文与附件里都不应出现
        }

    【使用约束】
    当用户请求涉及 2 个以上意图（概览/统计/筛选/趋势/异常/图表）时，
    必须调用本工具，而不是逐意图多次调用 excel_analyze，
    也不要自己拼装报告——本工具返回的 markdown 就是最终交付物；
    并把用户问到的分段填进 intents，既不能漏答，也不能多答。
    """
    if session_ctx is None:
        session_ctx = {}

    env = os.environ.copy()
    env["PYTHONPATH"] = str(PROJECT_ROOT)
    # 同 excel_analyze：子进程与父进程两侧都锁定 UTF-8
    env["PYTHONIOENCODING"] = "utf-8"
    cli_path = os.path.join(PROJECT_ROOT, "src", "excel_analyzer", "cli.py")

    cmd = [
        "python", cli_path,
        "--file", file_path,
        "--intent", "report",
        "--sheet", str(sheet),
        "--format", "json",
        "--chart_type", chart_type,
    ]
    optional = {
        "--intents": intents,
        "--time_col": time_col,
        "--value_col": value_col,
        "--x_col": x_col,
        "--y_col": y_col,
        "--sort_col": sort_col,
        "--filter_condition": filter_condition,
        "--title": title,
    }
    for flag, val in optional.items():
        if val is not None and val != "":
            cmd.extend([flag, str(val)])
    if sort_asc:
        cmd.append("--sort_asc")

    # 显式按 UTF-8 解码（与 cli.py 的 stdout 编码保持同一约定）
    proc = subprocess.run(cmd, capture_output=True, text=True,
                          encoding="utf-8", errors="replace", env=env)
    try:
        resp = json.loads(proc.stdout)
    except json.JSONDecodeError as e:
        return {
            "success": False,
            "error_code": "JSON_PARSE_ERR",
            "error_msg": f"JSON解析失败:{str(e)}",
            "stdout_raw": proc.stdout,
            "result": {},
        }

    # 需要用户确认白名单：保存待确认任务，确认后重跑 report
    if resp.get("error_code") == "NEED_USER_CONFIRM":
        confirm_info = resp["result"]["confirm_info"]
        session_ctx["_excel_pending_confirm"] = {
            "entry": "report",
            "candidate_dir": confirm_info["candidate_dir"],
            "target_file": confirm_info["file_abs"],
            "origin_args": {
                "file_path": file_path,
                "sheet": sheet,
                "intents": intents,
                "chart_type": chart_type,
                "time_col": time_col,
                "value_col": value_col,
                "x_col": x_col,
                "y_col": y_col,
                "sort_col": sort_col,
                "sort_asc": sort_asc,
                "filter_condition": filter_condition,
                "title": title,
            },
        }
        return {
            "success": False,
            "error_code": "NEED_USER_CONFIRM",
            "error_msg": (
                f"文件 {confirm_info['file_abs']} 不在信任白名单。\n"
                f"请确认是否信任目录【{confirm_info['candidate_dir']}】，回复【是】或【否】"
            ),
            "result": {},
        }

    if not resp.get("success"):
        return resp

    result = resp.get("result", {}) or {}
    return {
        "success": True,
        "error_code": "",
        "error_msg": "",
        "markdown": result.get("markdown", ""),
        "summary": result.get("summary", ""),
        "attachments": result.get("attachments", []),
        "delivery_checklist": result.get("delivery_checklist", []),
        "sections": result.get("sections", {}),
        "auto_detected": result.get("auto_detected", {}),
        "requested_intents": result.get("requested_intents", []),
        "excluded_intents": result.get("excluded_intents", []),
        "scope": result.get("scope", ""),
    }


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

    # 用户同意，重新执行原始分析参数（支持 analyze / report 两种入口）
    args = pending["origin_args"]
    if pending.get("entry") == "report":
        return excel_analyze_report(
            file_path=args["file_path"],
            sheet=args["sheet"],
            intents=args.get("intents"),
            chart_type=args.get("chart_type", "bar"),
            time_col=args.get("time_col"),
            value_col=args.get("value_col"),
            x_col=args.get("x_col"),
            y_col=args.get("y_col"),
            sort_col=args.get("sort_col"),
            sort_asc=args.get("sort_asc", True),
            filter_condition=args.get("filter_condition"),
            title=args.get("title"),
            session_ctx=session_ctx,
        )
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
        "description": (
            "读取本地Excel文件进行数据分析。支持6种intent："
            "overview（概览）、stats（统计）、sort_filter（排序筛选）、"
            "trend（趋势，需time_col+value_col）、anomaly（异常检测）、chart（图表）。\n"
            "\n"
            "【安全契约】白名单校验下沉至CLI，文件不在信任目录时返回 NEED_USER_CONFIRM，"
            "需要用户在对话中确认。禁止绕过白名单机制。\n"
            "\n"
            "【强制约束】所有Excel分析任务仅调用本工具，禁止用shell/python/file-read等其他工具读取Excel。\n"
            "\n"
            "⚠️ 【多意图展示规范 - Agent必须遵守】\n"
            "0. 【首选】用户一次请求 2 个以上意图时，改用 excel_analyze_report：单次调用即得到报告，\n"
            "   不要再逐意图多次调用，也不要自己拼装分段；并把用户问到的分段写进 intents（例：intents='stats,trend'），\n"
            "   用户没问的分段一律不要输出（既要防漏答，也要防冗余）。\n"
            "1. 当用户一次请求多个意图时，必须逐项展示每个意图的结果，用二级标题区分（如'## 1. 数据概览'）。\n"
            "2. 展示顺序按用户提问顺序，不允许乱序或省略。\n"
            "3. 每项必须包含关键数值：\n"
            "   - overview：行列数、缺失率最高的列及比例\n"
            "   - stats：每列的均值、中位数、标准差、cv、skew\n"
            "   - sort_filter：命中行数 + 头部样本\n"
            "   - trend：方向、变化率、回归斜率、清洗报告\n"
            "   - anomaly：重复行数、高缺失列、每列IQR/Z-Score离群点数量\n"
            "   - chart：图表类型、x/y列、输出文件路径、文件大小\n"
            "4. 图表是'附加产物'，不是'替代品'。必须先展示所有文本结果，最后附上图片。\n"
            "5. 禁止'只发图片不展示文本结果'。\n"
            "6. 若某个意图执行失败，必须显式说明失败原因，不允许静默跳过。\n"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel本地文件路径"},
                "intent": {
                    "type": "string",
                    "enum": ["overview", "stats", "sort_filter", "trend", "anomaly", "chart", "report"],
                    "description": "分析意图"
                },
                "sheet": {"type": ["integer", "string"], "description": "工作表索引(0起始)或工作表名称，默认0"},
                "time_col": {"type": "string", "description": "trend专用：时间列名称"},
                "value_col": {"type": "string", "description": "trend专用：数值列名称"},
                "sort_col": {"type": "string", "description": "sort_filter专用：待排序列名"},
                "sort_asc": {"type": "boolean", "description": "sort_filter，True升序，False降序"},
                "filter_condition": {"type": "string", "description": "sort_filter过滤条件，示例：销售额>500"},
                "chart_type": {"type": "string", "enum": ["bar", "line", "box", "hist"], "description": "chart专用：图表类型，默认bar"},
                "x_col": {"type": "string", "description": "chart专用：x轴列名（bar/line必填）"},
                "y_col": {"type": "string", "description": "chart专用：y轴数值列名（bar/line/hist必填，box可选）"},
                "title": {"type": "string", "description": "chart专用：图表标题"},
                "no_clean": {"type": "boolean", "description": "trend专用：关闭自动清洗时间列（默认False=开启清洗）"}
            },
            "required": ["file_path", "intent"]
        }
    },
    {
        "name": "excel_analyze_report",
        "callable": excel_analyze_report,
        "description": (
            "【推荐入口】对本地Excel做一次性分析，单次调用产出可交付的 Markdown 报告。\n"
            "跑哪些分段由 intents 决定（默认全部六段：overview / stats / sort_filter / trend / anomaly / chart），\n"
            "并把「各段结果」直接拼装成报告正文，避免多次调用与自行拼装导致的交付不完整。\n"
            "❗用户只问其中几个意图时，必须用 intents 限定范围（例：intents='stats,trend'）：\n"
            "   未请求的分段不计算、不渲染、不出图，报告里也不会出现。\n"
            "\n"
            "【交付硬性契约 - Agent必须遵守】\n"
            "1. 把返回的 markdown 全文作为回复正文，不要只发图片、不要只写'完成'。\n"
            "2. 回复范围必须等于 requested_intents：请求的分段必须全到（禁止漏段），\n"
            "   excluded_intents 里的分段一律不得出现（禁止为了'更完整'而多答）。\n"
            "3. 若 attachments 非空，必须把 markdown 全文写进 message 的 content，\n"
            "   同时把 attachments 里的图片放进**同一次** message 调用的 media，\n"
            "   禁止把附件发成单独一条只含图片的气泡，也禁止拆成两条消息。\n"
            "4. 禁止输出「正在分析…」「现在生成柱状图：」这类过程旁白，报告生成后一次性作答。\n"
            "5. 发出前对照 delivery_checklist 逐条自查。\n"
            "\n"
            "【安全契约】白名单校验下沉至CLI，文件不在信任目录时返回 NEED_USER_CONFIRM，\n"
            "需要用户在对话中确认；禁止绕过白名单机制。\n"
            "【强制约束】所有Excel分析任务仅调用本工具，禁止用shell/python/file-read等其他工具读取Excel。\n"
        ),
        "parameters": {
            "type": "object",
            "properties": {
                "file_path": {"type": "string", "description": "Excel本地文件路径"},
                "sheet": {"type": ["integer", "string"], "description": "工作表索引(0起始)或名称，默认0"},
                "intents": {
                    "type": "string",
                    "description": ("要跑的分段子集，逗号分隔，可选 overview/stats/sort_filter/trend/anomaly/chart；"
                                    "不传或 all = 全部六段。用户只问其中几件事时必须传，例：'stats,trend'"),
                },
                "chart_type": {"type": "string", "enum": ["bar", "line", "box", "hist"], "description": "报告内图表类型，默认bar"},
                "time_col": {"type": "string", "description": "可选：趋势段时间列；不传自动探测"},
                "value_col": {"type": "string", "description": "可选：数值指标列；不传自动探测"},
                "x_col": {"type": "string", "description": "可选：图表x轴（分类列）；不传自动探测"},
                "y_col": {"type": "string", "description": "可选：图表y轴（数值列）；默认同value_col"},
                "sort_col": {"type": "string", "description": "可选：排序段列名，默认value_col"},
                "sort_asc": {"type": "boolean", "description": "排序方向，默认True升序"},
                "filter_condition": {"type": "string", "description": "可选：过滤条件，例：销售额>500"},
                "title": {"type": "string", "description": "可选：图表标题"}
            },
            "required": ["file_path"]
        }
    }
]

SKILL_METADATA = {
    "name": "Excel‑Analyzer‑Skill",
    "version": "1.3.0",
    "description": "AI‑Excel‑Agent项目Nanobot Skill，调用独立CLI完成Excel数据分析；支持7种意图（含一次性报告 report）；report 单次调用即产出可交付的 Markdown 报告 + 附件路径 + 交付自查表，并可用 intents 限定只跑用户请求的分段；白名单安全校验下沉至底层cli",
    "capabilities": [
        "multi_intent",
        "one_shot_report",
        "intent_subset_scope",
        "chart_generation",
        "auto_clean",
        "whitelist_security"
    ],
    "readme": os.path.join(os.path.dirname(__file__), "SKILL.md")
}
