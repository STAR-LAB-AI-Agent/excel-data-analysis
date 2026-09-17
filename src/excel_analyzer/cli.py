"""
cli.py
命令行入口层
负责参数解析、校验、调用core、异常捕获、输出标准JSON结果

校验顺序（从低成本到高成本）：
    1. 参数完整性（file/intent 必须提供）
    2. 文件存在性
    3. 文件后缀
    4. 意图合法性
    5. 意图专属参数
    6. 白名单校验
    7. 业务逻辑
"""
import argparse
import json
import sys
import os
import math
from datetime import datetime

from src.excel_analyzer.schemas import (
    VALID_INTENTS,
    VALID_CHART_TYPES,
    VALID_REPORT_FORMATS,
    ALLOWED_EXCEL_SUFFIX,
    build_response,
    ERR_INTERNAL_ERROR,
    ERR_NEED_USER_CONFIRM,
    ERR_FILE_NOT_FOUND,
    ERR_INVALID_SUFFIX,
    INTENT_SORT_FILTER,
    INTENT_TREND,
    INTENT_CHART,
    INTENT_REPORT,
)
from src.excel_analyzer.exceptions import (
    BaseBusinessError,
    InvalidIntentError,
    MissingArgumentError,
)
from src.excel_analyzer.logger_cfg import logger
from src.excel_analyzer.analyzer_core import (
    load_excel,
    do_overview,
    do_stats,
    do_sort_filter,
    do_trend,
    do_anomaly,
    do_chart,
)
from src.excel_analyzer.report_builder import build_report
from src.excel_analyzer.whitelist_manager import (
    interactive_check,
    list_whitelist,
    remove_whitelist_by_index,
    agent_precheck,
)


def _force_utf8_stdio() -> None:
    """把 stdout/stderr 固定为 UTF-8。

    Windows 下控制台与管道默认使用本地编码（简体中文为 cp936/GBK）：
    中文报告按 GBK 写出、上层按 UTF-8 解析时就会整段乱码。
    这里在进程启动时统一重配置，使 stdout 字节流**与平台代码页、
    PYTHONIOENCODING 无关地恒为 UTF-8**，调用方（Agent / subprocess /
    重定向到文件）按 UTF-8 解码即可，无需任何编码兜底技巧。
    """
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            # AttributeError：非文本流或旧解释器；ValueError：流已关闭
            pass


_force_utf8_stdio()


class DateTimeEncoder(json.JSONEncoder):
    """JSON序列化：datetime转 ISO；NaN/Inf 转 None"""
    def default(self, obj):
        if isinstance(obj, datetime):
            return obj.isoformat()
        if isinstance(obj, float):
            if math.isnan(obj) or math.isinf(obj):
                return None
        return super().default(obj)


def _output_and_exit(resp: dict, output_json: str = None):
    """统一输出并退出"""
    json_text = json.dumps(resp, ensure_ascii=False, indent=2, cls=DateTimeEncoder)
    if output_json:
        with open(output_json, "w", encoding="utf-8") as f:
            f.write(json_text)
    else:
        print(json_text)
    sys.exit(0 if resp.get("success") else 1)


def main():
    parser = argparse.ArgumentParser(description="AI Excel数据分析CLI工具")
    parser.add_argument("--file", required=False, help="Excel文件路径(.xlsx/.xls)")
    parser.add_argument("--intent", required=False, help=f"分析意图，可选:{','.join(sorted(VALID_INTENTS))}")
    parser.add_argument("--sheet", type=str, default="0", help="sheet名称或者索引数字，默认0")

    # sort_filter 参数
    parser.add_argument("--sort_col", type=str, default=None, help="sort_filter：排序列名")
    parser.add_argument("--sort_asc", action="store_true", default=None,
                        help="sort_filter：升序，默认降序（report：默认升序）")
    parser.add_argument("--filter_condition", type=str, default=None, help="sort_filter：过滤条件，例：销售额>1000")

    # trend 参数
    parser.add_argument("--time_col", type=str, default=None, help="trend：时间列名称")
    parser.add_argument("--value_col", type=str, default=None, help="trend：分析指标数值列")
    parser.add_argument("--no_clean", action="store_true", help="trend：关闭自动清洗时间列")

    # chart 参数
    parser.add_argument("--chart_type", type=str, default="bar",
                        choices=sorted(VALID_CHART_TYPES),
                        help="chart：图表类型 bar/line/box/hist")
    parser.add_argument("--x_col", type=str, default=None, help="chart：x轴列名")
    parser.add_argument("--y_col", type=str, default=None, help="chart：y轴列名（数值列）")
    parser.add_argument("--title", type=str, default=None, help="chart：图表标题")

    # 输出
    parser.add_argument("--output_json", type=str, default=None, help="输出JSON文件路径；不填输出到stdout")
    parser.add_argument("--format", type=str, default="json", choices=sorted(VALID_REPORT_FORMATS),
                        help="intent=report 时的stdout格式：json（默认）或 md（直接输出可交付的Markdown正文）")

    # 白名单管理
    parser.add_argument("--list-whitelist", action="store_true", help="列出全部信任白名单目录")
    parser.add_argument("--remove-whitelist-idx", type=int, help="删除指定序号的白名单目录")

    args = parser.parse_args()

    # ========== 白名单管理子命令（不需要 file/intent）==========
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

    meta = {
        "file_path": args.file,
        "sheet_name": args.sheet,
        "intent": args.intent,
    }

    # ========== 第 1 层：参数完整性 ==========
    if args.file is None or args.intent is None:
        _output_and_exit(
            build_response(False, error_code="MISSING_ARG",
                           error_msg="--file 和 --intent 为必选参数", meta=meta),
            args.output_json
        )

    # ========== 第 2 层：文件存在性（白名单之前，成本最低）==========
    if not os.path.exists(args.file):
        _output_and_exit(
            build_response(False, error_code=ERR_FILE_NOT_FOUND,
                           error_msg=f"文件不存在：{args.file}", meta=meta),
            args.output_json
        )

    # ========== 第 3 层：文件后缀 ==========
    _, ext = os.path.splitext(args.file)
    if ext.lower() not in ALLOWED_EXCEL_SUFFIX:
        _output_and_exit(
            build_response(False, error_code=ERR_INVALID_SUFFIX,
                           error_msg=f"不支持文件后缀 {ext}，仅支持 {ALLOWED_EXCEL_SUFFIX}", meta=meta),
            args.output_json
        )

    # ========== 第 4 层：意图合法性 ==========
    intent = args.intent.strip()
    if intent not in VALID_INTENTS:
        _output_and_exit(
            build_response(False, error_code="INVALID_INTENT",
                           error_msg=f"非法意图[{intent}]，合法意图仅支持：{','.join(sorted(VALID_INTENTS))}",
                           meta=meta),
            args.output_json
        )

    # ========== 第 5 层：意图专属参数 ==========
    if intent == INTENT_TREND:
        if args.time_col is None or args.value_col is None:
            _output_and_exit(
                build_response(False, error_code="MISSING_ARG",
                               error_msg="intent=trend 需要提供 --time_col 和 --value_col 参数",
                               meta=meta),
                args.output_json
            )
    if intent == INTENT_CHART:
        if args.chart_type in ("bar", "line"):
            if args.x_col is None or args.y_col is None:
                _output_and_exit(
                    build_response(False, error_code="MISSING_ARG",
                                   error_msg="chart_type=bar/line 需要提供 --x_col 和 --y_col 参数",
                                   meta=meta),
                    args.output_json
                )
        if args.chart_type == "hist":
            if args.y_col is None:
                _output_and_exit(
                    build_response(False, error_code="MISSING_ARG",
                                   error_msg="chart_type=hist 需要提供 --y_col 参数",
                                   meta=meta),
                    args.output_json
                )

    # ========== 第 6 层：白名单校验（需要读配置，成本较高）==========
    is_tty = sys.stdin.isatty()
    if is_tty:
        access_ok, _ = interactive_check(args.file)
        if not access_ok:
            _output_and_exit(
                build_response(False, error_code="SECURITY_REJECT",
                               error_msg="用户拒绝访问该文件（白名单校验不通过）",
                               meta=meta),
                args.output_json
            )
    else:
        pre = agent_precheck(args.file)
        if not pre["allowed"]:
            _output_and_exit(
                build_response(False, error_code=ERR_NEED_USER_CONFIRM,
                               error_msg="文件不在白名单，需要用户确认是否添加目录",
                               meta=meta,
                               result={"confirm_info": pre}),
                args.output_json
            )

    # ========== 第 7 层：业务逻辑 ==========
    try:
        sheet_arg: int | str
        if args.sheet.isdigit():
            sheet_arg = int(args.sheet)
        else:
            sheet_arg = args.sheet

        logger.info(f"开始执行任务 | file={args.file} | intent={intent} | sheet={sheet_arg}")

        df = load_excel(args.file, sheet_name=sheet_arg)

        if intent == "overview":
            result = do_overview(df)
        elif intent == "stats":
            result = do_stats(df)
        elif intent == INTENT_SORT_FILTER:
            result = do_sort_filter(
                df,
                sort_col=args.sort_col,
                sort_asc=bool(args.sort_asc),
                filter_condition=args.filter_condition,
            )
        elif intent == INTENT_TREND:
            result = do_trend(
                df,
                time_col=args.time_col,
                value_col=args.value_col,
                auto_clean=not args.no_clean,
            )
        elif intent == "anomaly":
            result = do_anomaly(df)
        elif intent == INTENT_CHART:
            result = do_chart(
                df,
                chart_type=args.chart_type,
                x_col=args.x_col,
                y_col=args.y_col,
                title=args.title,
            )
        elif intent == INTENT_REPORT:
            # ✨ 一次性完整报告：单次调用产出六个分段的完整交付物
            result = build_report(
                df,
                file_path=args.file,
                sheet_name=str(sheet_arg),
                chart_type=args.chart_type,
                time_col=args.time_col,
                value_col=args.value_col,
                x_col=args.x_col,
                y_col=args.y_col,
                sort_col=args.sort_col,
                sort_asc=(True if args.sort_asc is None else bool(args.sort_asc)),
                filter_condition=args.filter_condition,
                title=args.title,
            )
            resp = build_response(success=True, meta=meta, result=result)
            logger.info(
                f"任务执行成功 | intent={intent} | sections={list(result['sections'].keys())} "
                f"| attachments={result['attachments']}"
            )
            if args.format == "md":
                # stdout 仅输出可直接作为回复正文的 Markdown；日志仍只走 stderr
                sys.stdout.write(result["markdown"])
                sys.stdout.flush()
                if args.output_json:
                    with open(args.output_json, "w", encoding="utf-8") as f:
                        f.write(json.dumps(resp, ensure_ascii=False, indent=2, cls=DateTimeEncoder))
                sys.exit(0)
            _output_and_exit(resp, args.output_json)
        else:
            result = {}

        resp = build_response(success=True, meta=meta, result=result)
        logger.info(f"任务执行成功 | intent={intent}")
        _output_and_exit(resp, args.output_json)

    except BaseBusinessError as be:
        logger.warning(f"业务异常 | code={be.error_code} | msg={be.message}")
        _output_and_exit(
            build_response(success=False, error_code=be.error_code,
                           error_msg=be.message, meta=meta),
            args.output_json
        )
    except Exception as e:
        logger.exception("未知内部异常")
        _output_and_exit(
            build_response(success=False, error_code=ERR_INTERNAL_ERROR,
                           error_msg=f"内部错误:{str(e)}", meta=meta),
            args.output_json
        )


if __name__ == "__main__":
    main()
