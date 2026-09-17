"""
schemas.py
常量、枚举、输出结构定义
"""
from typing import Any, Dict, List, Tuple

# ---------------------- 意图常量定义 ----------------------
VALID_INTENTS = {
    "overview",
    "stats",
    "sort_filter",
    "trend",
    "anomaly",
    "chart",          # ✨ 新增：图表生成
    "report",         # ✨ 新增：一次性完整报告（六意图编排）
}

INTENT_OVERVIEW = "overview"
INTENT_STATS = "stats"
INTENT_SORT_FILTER = "sort_filter"
INTENT_TREND = "trend"
INTENT_ANOMALY = "anomaly"
INTENT_CHART = "chart"  # ✨ 新增
INTENT_REPORT = "report"  # ✨ 新增：一次性完整报告

# ---------------------- 错误码常量 ----------------------
ERR_FILE_NOT_FOUND = "FILE_NOT_FOUND"
ERR_INVALID_SUFFIX = "INVALID_SUFFIX"
ERR_SHEET_NOT_EXIST = "SHEET_NOT_EXIST"
ERR_INVALID_INTENT = "INVALID_INTENT"
ERR_MISSING_ARG = "MISSING_ARG"
ERR_PARSE_EXCEL_FAILED = "PARSE_EXCEL_FAILED"
ERR_COLUMN_NOT_FOUND = "COLUMN_NOT_FOUND"
ERR_DATA_TYPE_ERR = "DATA_TYPE_ERR"
ERR_INTERNAL_ERROR = "INTERNAL_ERROR"
ERR_NEED_USER_CONFIRM = "NEED_USER_CONFIRM"
ERR_CHART_GEN_FAILED = "CHART_GEN_FAILED"      # ✨ 新增
ERR_NO_NUMERIC_COL = "NO_NUMERIC_COL"          # ✨ 新增
ERR_INVALID_ARG = "INVALID_ARG"                # ✨ 新增：参数取值非法（如 --intents 含未知分段）

# ---------------------- 文件允许后缀 ----------------------
ALLOWED_EXCEL_SUFFIX = {".xlsx", ".xls"}

# ---------------------- 图表类型 ----------------------
VALID_CHART_TYPES = {"bar", "line", "box", "hist"}

# ---------------------- report 意图输出格式 ----------------------
VALID_REPORT_FORMATS = {"json", "md"}

# ---------------------- report 分段（子集）定义 ----------------------
# 顺序即报告正文中段落的出现顺序；--intents 只接受这里的名字
REPORT_SECTION_ORDER = ("overview", "stats", "sort_filter", "trend", "anomaly", "chart")
REPORT_INTENTS_ALL = "all"


def parse_report_sections(raw) -> Tuple[List[str], List[str]]:
    """把「用户/调用方给出的 report 分段子集」解析成规范形式。

    参数 raw 可为 None / "" / "all" / "stats,trend" / ["stats", "trend"]（中英文逗号均可）。
    返回 (sections, unknown)：
      · sections：按 REPORT_SECTION_ORDER 规范排序并去重的分段列表；
        raw 为空或含 all 时返回全部六段（表示「未指定子集」）；
      · unknown：无法识别的分段名——调用方据此报错，
        避免「用户问过的分段被静默丢掉」这类隐蔽漏答。
    """
    if raw is None:
        tokens: List[str] = []
    elif isinstance(raw, str):
        tokens = raw.replace("，", ",").split(",")
    else:
        tokens = [str(t) for t in raw]
    tokens = [t.strip() for t in tokens]
    tokens = [t for t in tokens if t]
    if not tokens or any(t.lower() == REPORT_INTENTS_ALL for t in tokens):
        return list(REPORT_SECTION_ORDER), []
    unknown = [t for t in tokens if t not in REPORT_SECTION_ORDER]
    keep = set(tokens)
    return [s for s in REPORT_SECTION_ORDER if s in keep], unknown

# ---------------------- 构建标准返回JSON模板 ----------------------
def build_response(
    success: bool,
    error_code: str = "",
    error_msg: str = "",
    meta: Dict[str, Any] = None,
    result: Dict[str, Any] = None
) -> Dict[str, Any]:
    """统一构造返回对象"""
    resp = {
        "success": success,
        "error_code": error_code,
        "error_msg": error_msg,
        "meta": meta if meta is not None else {},
        "result": result if result is not None else {}
    }
    return resp
