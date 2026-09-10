"""
schemas.py
常量、枚举、输出结构定义
"""
from typing import Dict, Any
# ---------------------- 意图常量定义 ----------------------
VALID_INTENTS = {
    "overview",
    "stats",
    "sort_filter",
    "trend",
    "anomaly"
}
INTENT_OVERVIEW = "overview"
INTENT_STATS = "stats"
INTENT_SORT_FILTER = "sort_filter"
INTENT_TREND = "trend"
INTENT_ANOMALY = "anomaly"
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
ERR_NEED_USER_CONFIRM = "NEED_USER_CONFIRM"  # ✨新增：需要用户确认添加白名单

# ---------------------- 文件允许后缀 ----------------------
ALLOWED_EXCEL_SUFFIX = {".xlsx", ".xls"}

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