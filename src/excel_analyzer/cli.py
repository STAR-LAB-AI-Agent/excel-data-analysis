"""
cli.py
命令行入口层
负责参数解析、白名单校验、调用core、异常捕获、输出标准JSON结果
"""
import argparse
import json
import sys
import math
from datetime import datetime
from src.excel_analyzer.schemas import (
    VALID_INTENTS,
    build_response,
    ERR_INTERNAL_ERROR,
    ERR_NEED_USER_CONFIRM,
    INTENT_SORT_FILTER,
    INTENT_TREND
)
from src.excel_analyzer.exceptions import (
    BaseBusinessError,
    InvalidIntentError,
    MissingArgumentError
)
from src.excel_analyzer.logger_cfg import logger
from src.excel_analyzer.analyzer_core import (
    load_excel,
    do_overview,
    do_stats,
    do_sort_filter,
    do_trend,
    do_anomaly
)
from src.excel_analyzer.whitelist_manager import (
    interactive_check,
    list_whitelist,
    remove_whitelist_by_index,
    agent_precheck
)

class DateTimeEncoder(json.JSONEncoder):
    """JSON序列化：datetime对象转为iso格式字符串，处理NaN"""
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
        return super().default(obj)

def main():
    parser = argparse.ArgumentParser(description="AI Excel数据分析CLI工具")
    parser.add_argument("--file", required=False, help="Excel文件路径(.xlsx/.xls)")
    parser.add_argument("--intent", required=False, help=f"分析意图，可选:{','.join(VALID_INTENTS)}")
    parser.add_argument("--sheet", type=str, default="0", help="sheet名称或者索引数字，默认0(第一个sheet)")
    parser.add_argument("--sort_col", type=str, default=None, help="sort_filter：排序列名")
    parser.add_argument("--sort_asc", action="store_true", help="sort_filter：升序排序，默认关闭(降序)")
    parser.add_argument("--filter_condition", type=str, default=None, help="sort_filter：过滤条件，例：销售额>1000")
    parser.add_argument("--time_col", type=str, default=None, help="trend：时间列名称")
    parser.add_argument("--value_col", type=str, default=None, help="trend：分析指标数值列")
    parser.add_argument("--output_json", type=str, default=None, help="输出JSON文件路径；不填输出到stdout")
    parser.add_argument("--list-whitelist", action="store_true", help="列出全部信任白名单目录")
    parser.add_argument("--remove-whitelist-idx", type=int, help="删除指定序号的白名单目录")

    args = parser.parse_args()

    # 白名单管理子命令，不需要文件参数
    if args.list_whitelist:
        list_whitelist()
        sys.exit(0)
    if args.remove_whitelist_idx is not None:
        try:
            remove_whitelist_by_index(args.remove_whitelist_idx)
        except IndexError as e:
            print(json.dumps(build_response(False, error_code=ERR_INTERNAL_ERROR, error_msg=str(e)), ensure_ascii=False))
            sys.exit(1)
        sys.exit(0)

    # 业务分析模式必须提供 file / intent
    if args.file is None or args.intent is None:
        print(json.dumps(build_response(False, error_code="MISSING_ARG", error_msg="--file 和 --intent 为必选参数"), ensure_ascii=False))
        sys.exit(1)

    meta = {
        "file_path": args.file,
        "sheet_name": args.sheet,
        "intent": args.intent
    }

    # ==========白名单校验阶段==========
    is_tty = sys.stdin.isatty()
    if is_tty:
        # 本地终端交互模式：控制台input询问用户
        access_ok, _ = interactive_check(args.file)
        if not access_ok:
            resp = build_response(
                success=False,
                error_code="SECURITY_REJECT",
                error_msg="用户拒绝访问该文件（白名单校验不通过）",
                meta=meta
            )
            print(json.dumps(resp, ensure_ascii=False, cls=DateTimeEncoder))
            sys.exit(1)
    else:
        # 非TTY：被subprocess调用，不能input，返回NEED_USER_CONFIRM结构给上层Agent
        pre = agent_precheck(args.file)
        if not pre["allowed"]:
            resp = build_response(
                success=False,
                error_code=ERR_NEED_USER_CONFIRM,
                error_msg="文件不在白名单，需要用户确认是否添加目录",
                meta=meta,
                result={"confirm_info": pre}
            )
            print(json.dumps(resp, ensure_ascii=False, cls=DateTimeEncoder))
            sys.exit(1)

    try:
        intent = args.intent.strip()
        if intent not in VALID_INTENTS:
            raise InvalidIntentError(f"非法意图[{intent}]，合法意图仅支持：{','.join(VALID_INTENTS)}")
        sheet_arg: int | str
        if args.sheet.isdigit():
            sheet_arg = int(args.sheet)
        else:
            sheet_arg = args.sheet

        logger.info(f"开始执行任务 | file={args.file} | intent={intent} | sheet={sheet_arg}")
        if intent == INTENT_TREND:
            if args.time_col is None or args.value_col is None:
                raise MissingArgumentError(f"intent=trend 需要提供 --time_col 和 --value_col 参数")

        df = load_excel(args.file, sheet_name=sheet_arg)
        result = {}
        if intent == "overview":
            result = do_overview(df)
        elif intent == "stats":
            result = do_stats(df)
        elif intent == INTENT_SORT_FILTER:
            result = do_sort_filter(
                df,
                sort_col=args.sort_col,
                sort_asc=args.sort_asc,
                filter_condition=args.filter_condition
            )
        elif intent == INTENT_TREND:
            result = do_trend(df, time_col=args.time_col, value_col=args.value_col)
        elif intent == "anomaly":
            result = do_anomaly(df)

        resp = build_response(success=True, meta=meta, result=result)
        logger.info(f"任务执行成功 | intent={intent}")

    except BaseBusinessError as be:
        logger.warning(f"业务异常 | code={be.error_code} | msg={be.message}")
        resp = build_response(
            success=False,
            error_code=be.error_code,
            error_msg=be.message,
            meta=meta
        )
    except Exception as e:
        logger.exception("未知内部异常")
        resp = build_response(
            success=False,
            error_code=ERR_INTERNAL_ERROR,
            error_msg=f"内部错误:{str(e)}",
            meta=meta
        )

    json_text = json.dumps(resp, ensure_ascii=False, indent=2, cls=DateTimeEncoder)
    if args.output_json is not None:
        with open(args.output_json, "w", encoding="utf-8") as f:
            f.write(json_text)
    else:
        print(json_text)

if __name__ == "__main__":
    main()