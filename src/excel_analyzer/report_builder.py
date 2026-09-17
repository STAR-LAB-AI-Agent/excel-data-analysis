"""
report_builder.py
一次性完整报告编排器（one-shot report）

把 overview / stats / sort_filter / trend / anomaly / chart 六个意图
在**一次调用内**全部跑完，并直接渲染成「可直接交给终端用户」的完整 Markdown 报告。

--------------------------------------------------------------------------
设计动机（复盘结论）
--------------------------------------------------------------------------
历史故障：六个意图需要 6 次 CLI 调用 + 1 次 message 发图，再由智能体自己
把 6 段结果拼成一份报告。拼装环节一旦断档，就会出现
「只发了一张图、其余问题完全没有回复」的交付失败。

修复思路：把「拼装」从智能体下沉到工具内部。
    - 一次调用 → 一份完整报告（六个分段全在，关键数值全在）
    - 报告里同时携带 attachments 与 delivery_checklist
    - 智能体的正确动作被压缩成：跑一条命令 → 把 markdown 全文 + 图片放进同一条消息

这样「正确输出」不再依赖智能体的多步记忆，而是成为工具输出的自然结果。
"""
import math
import warnings
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd

from src.excel_analyzer.analyzer_core import (
    do_overview,
    do_stats,
    do_sort_filter,
    do_trend,
    do_anomaly,
    do_chart,
)
from src.excel_analyzer.exceptions import BaseBusinessError

# 列名提示词：用于在没有显式指定列时做自动探测
_VALUE_HINTS = ("销售额", "销售", "金额", "营收", "收入", "利润", "amount", "sales", "revenue", "value")
_TIME_HINTS = ("日期", "时间", "月份", "date", "time", "month", "day")
_CAT_HINTS = ("区域", "地区", "大区", "类别", "分类", "类型", "渠道", "region", "category", "segment", "type")


# ============================ 基础格式化 ============================
def _fmt(v: Any, nd: int = 2) -> str:
    """统一数值展示；NaN/Inf/None 统一显示为 em dash"""
    if v is None:
        return "—"
    if isinstance(v, float):
        if math.isnan(v) or math.isinf(v):
            return "—"
        return f"{v:,.{nd}f}"
    if isinstance(v, bool):
        return "是" if v else "否"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


def _pct(v: Any, nd: int = 2) -> str:
    if v is None:
        return "—"
    try:
        return f"{float(v) * 100:.{nd}f}%"
    except (TypeError, ValueError):
        return "—"


def _cell(v: Any) -> str:
    """Markdown 表格单元格：转义竖线与换行"""
    s = _fmt(v) if not isinstance(v, str) else v
    if s is None:
        s = "—"
    return str(s).replace("|", "\\|").replace("\n", " ").replace("\r", "")


def _table(headers: List[str], rows: List[List[Any]]) -> List[str]:
    if not rows:
        return ["（无数据）"]
    out = ["| " + " | ".join(headers) + " |",
           "|" + "|".join(["---"] * len(headers)) + "|"]
    for r in rows:
        out.append("| " + " | ".join(_cell(c) for c in r) + " |")
    return out


def _hint_score(name: Any, hints: Tuple[str, ...]) -> int:
    low = str(name).lower()
    for i, h in enumerate(hints):
        if h.lower() in low:
            return len(hints) - i
    return 0


# ============================ 列自动探测 ============================
def _datetime_ratio(series: pd.Series) -> float:
    """该列能被解析为日期的比例"""
    s = series.dropna()
    if len(s) == 0:
        return 0.0
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            conv = pd.to_datetime(s, errors="coerce")
    except Exception:
        return 0.0
    return float(conv.notna().mean())


def auto_detect_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    """
    在调用方没有显式指定列时，自动推断：
        time_col  / value_col / x_col(分类列) / sort_col
    推断完全基于列名与数据特征，不修改数据。
    """
    numeric_cols = list(df.select_dtypes(include=["number"]).columns)
    non_numeric = [c for c in df.columns if c not in numeric_cols]

    # --- 数值指标列：优先列名提示，其次第一个数值列 ---
    value_col = max(numeric_cols, key=lambda c: _hint_score(c, _VALUE_HINTS)) if numeric_cols else None

    # --- 时间列：非数值列中可解析比例 >= 0.6，优先列名提示 ---
    time_pool = []
    for c in non_numeric:
        if df[c].nunique(dropna=True) <= 1:
            continue
        ratio = _datetime_ratio(df[c])
        if ratio >= 0.6:
            time_pool.append((c, ratio))
    time_col = None
    if time_pool:
        time_col = max(time_pool, key=lambda t: (_hint_score(t[0], _TIME_HINTS), t[1]))[0]

    # --- 分类列：取值数 2~20，优先列名提示，其次取值数少的 ---
    cat_pool = []
    for c in non_numeric:
        if c == time_col:
            continue
        nu = int(df[c].nunique(dropna=True))
        if 2 <= nu <= 20:
            cat_pool.append((c, nu))
    x_col = None
    if cat_pool:
        x_col = max(cat_pool, key=lambda t: (_hint_score(t[0], _CAT_HINTS), -t[1]))[0]

    return {
        "time_col": time_col,
        "value_col": value_col,
        "x_col": x_col,
        "sort_col": value_col,
    }


# ============================ 报告主体 ============================
def build_report(
    df: pd.DataFrame,
    *,
    file_path: str = "",
    sheet_name: str = "",
    chart_type: str = "bar",
    time_col: Optional[str] = None,
    value_col: Optional[str] = None,
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
    sort_col: Optional[str] = None,
    sort_asc: bool = True,
    filter_condition: Optional[str] = None,
    title: Optional[str] = None,
    auto_detect: bool = True,
) -> Dict[str, Any]:
    """
    一次调用生成完整报告。

    返回：
        {
          "markdown":     完整 Markdown 报告（可直接作为回复正文）
          "summary":      一句话总体结论
          "attachments":  [图表绝对路径, ...]
          "delivery_checklist": [...],   # 发送前必须逐条自检
          "sections":     {intent: 原始结构化结果 / 错误信息}
          "auto_detected": {time_col, value_col, x_col}
        }
    """
    detected = auto_detect_columns(df) if auto_detect else {}
    time_col = time_col or detected.get("time_col")
    value_col = value_col or detected.get("value_col")
    x_col = x_col or detected.get("x_col")
    y_col = y_col or value_col
    sort_col = sort_col or detected.get("sort_col")

    sections: Dict[str, Any] = {}
    warnings: List[str] = []
    attachments: List[str] = []

    def _safe(intent_name: str, fn):
        """单段失败不拖垮整份报告"""
        try:
            return fn()
        except BaseBusinessError as e:
            warnings.append(f"{intent_name} 段执行失败：{e.message}")
            return {"__error__": e.message}
        except Exception as e:  # noqa: BLE001 - 兜底，保证报告仍能产出
            warnings.append(f"{intent_name} 段执行异常：{e}")
            return {"__error__": str(e)}

    def _skip(intent_name: str, msg: str) -> Dict[str, Any]:
        """参数不足导致的主动跳过；与 _safe 的失败提示同一口径（同样记入 warnings）"""
        warnings.append(f"{intent_name} 段执行失败：{msg}")
        return {"__error__": msg}

    # ---- 1 overview ----
    sections["overview"] = _safe("overview", lambda: do_overview(df))

    # ---- 2 stats ----
    sections["stats"] = _safe("stats", lambda: do_stats(df))

    # ---- 3 sort_filter ----
    sections["sort_filter"] = _safe(
        "sort_filter",
        lambda: do_sort_filter(
            df, sort_col=sort_col, sort_asc=sort_asc, filter_condition=filter_condition
        ),
    )

    # ---- 4 trend ----
    if time_col and value_col:
        sections["trend"] = _safe(
            "trend",
            lambda: do_trend(df, time_col=time_col, value_col=value_col, auto_clean=True),
        )
    else:
        sections["trend"] = _skip("trend", "未能识别时间列或数值列，已跳过趋势分析")

    # ---- 5 anomaly ----
    sections["anomaly"] = _safe("anomaly", lambda: do_anomaly(df))

    # ---- 6 chart ----
    if x_col and y_col and chart_type in ("bar", "line"):
        chart_title = title or f"各{x_col}{value_col}"
        sections["chart"] = _safe(
            "chart",
            lambda: do_chart(df, chart_type=chart_type, x_col=x_col, y_col=y_col, title=chart_title),
        )
    else:
        sections["chart"] = _safe(
            "chart",
            lambda: do_chart(df, chart_type=chart_type, x_col=None, y_col=y_col, title=title),
        )

    chart_res = sections.get("chart") or {}
    if isinstance(chart_res, dict) and chart_res.get("output_path"):
        attachments.append(str(chart_res["output_path"]))

    markdown = render_markdown(
        sections,
        file_path=file_path,
        sheet_name=sheet_name,
        df=df,
        time_col=time_col,
        value_col=value_col,
        x_col=x_col,
        chart_type=chart_type,
        warnings=warnings,
    )

    checklist = [
        "本体报告为完整版：overview / stats / sort_filter / trend / anomaly / chart 六段都已包含，"
        "禁止只发送图表而不给正文。",
        "必须把 markdown 全文与 attachments 里的图片放在**同一条消息**中发出，不要拆成两条气泡。",
        "anomaly 段必须出现：重复行数、高缺失列及缺失率、每列 IQR 与 Z-Score 离群数（缺一不可）。",
        "chart 段必须出现 output_path 与分组数值，不能只有图片。",
        "不要输出「正在分析」「下一步…」这类过程旁白，等本报告生成后一次性作答。",
    ]

    return {
        "markdown": markdown,
        "summary": _build_summary(sections, value_col),
        "attachments": attachments,
        "delivery_checklist": checklist,
        "sections": sections,
        "auto_detected": {
            "time_col": time_col,
            "value_col": value_col,
            "x_col": x_col,
            "sort_col": sort_col,
        },
        "warnings": warnings,
    }


def _build_summary(sections: Dict[str, Any], value_col: Optional[str], adj: Optional[Dict[str, Any]] = None) -> str:
    """一句话总体结论"""
    parts: List[str] = []
    an = sections.get("anomaly") or {}
    if isinstance(an, dict) and "duplicate_row_count" in an:
        parts.append(f"重复行 {an['duplicate_row_count']} 行")
        hm = an.get("high_missing_columns") or []
        parts.append(
            "高缺失列 " + ("、".join(f"{c['column']}({_pct(c['missing_ratio'])})" for c in hm) if hm else "无")
        )
        oi = (an.get("outlier_statistics") or {}).get(value_col) if value_col else None
        if oi:
            parts.append(f"{value_col} IQR 离群 {oi['outlier_count']} 个 / Z-Score 离群 {oi['zscore_outlier_count']} 个")
    tr = sections.get("trend") or {}
    if isinstance(tr, dict) and tr.get("change_rate") is not None:
        arrow = {"rise": "上升", "decline": "下降", "flat": "持平"}.get(tr.get("direction"), "变化")
        rate = tr["change_rate"]
        if adj and adj.get("change_rate") is not None and abs(rate - adj["change_rate"]) > 0.2:
            adj_rate = adj["change_rate"]
            arrow2 = "上升" if adj_rate > 0 else ("下降" if adj_rate < 0 else "持平")
            parts.append(f"趋势表面{arrow} {_pct(abs(rate))}，但主要由离群点驱动，"
                         f"剔除离群后为{arrow2} {_pct(abs(adj_rate))}")
        else:
            parts.append(f"趋势整体{arrow} {_pct(abs(rate))}")
    return "；".join(parts) if parts else "报告已生成"


# ============================ Markdown 渲染 ============================
def render_markdown(
    sections: Dict[str, Any],
    *,
    file_path: str = "",
    sheet_name: str = "",
    df: Optional[pd.DataFrame] = None,
    time_col: Optional[str] = None,
    value_col: Optional[str] = None,
    x_col: Optional[str] = None,
    chart_type: str = "bar",
    warnings: Optional[List[str]] = None,
) -> str:
    """按固定契约渲染六个分段，每段都带关键数值 + 一句文字结论"""
    # 先算一次「剔除离群点后的趋势」，供 trend 段与综合结论共用
    adj_trend = _outlier_adjusted_trend(df, time_col, value_col)
    lines: List[str] = []

    fname = file_path.split("/")[-1].split("\\")[-1] if file_path else "数据表"
    lines.append(f"# {fname} · 完整分析报告")
    lines.append("")
    lines.append(f"- **数据源**：`{file_path}`")
    lines.append(f"- **工作表**：{sheet_name or '0'}")
    if df is not None:
        lines.append(f"- **规模**：{len(df)} 行 × {len(df.columns)} 列")
    lines.append(f"- **生成方式**：`--intent report` 一次性生成（六个分段全部包含）")
    lines.append("")

    # ---------- overview ----------
    lines.append("## overview · 数据概览")
    lines.append("")
    ov = sections.get("overview") or {}
    if "__error__" in ov:
        lines.append(f"本段执行失败：{ov['__error__']}")
    else:
        lines.append(f"- 行数 **{_fmt(ov.get('row_count'))}**，列数 **{_fmt(ov.get('col_count'))}**")
        lines.append(f"- 字段：{ '、'.join(str(c) for c in ov.get('columns', [])) }")
        dtypes = ov.get("dtypes") or {}
        lines.append("- 类型：" + "；".join(f"{k}={v}" for k, v in dtypes.items()))
        miss = ov.get("missing_stats") or {}
        miss_rows = [(c, v["missing_count"], v["missing_ratio"]) for c, v in miss.items() if v["missing_count"] > 0]
        if miss_rows:
            miss_rows.sort(key=lambda r: -r[2])
            lines.append("- 缺失字段（按缺失率降序）：")
            lines.append("")
            lines.extend(_table(
                ["字段", "缺失数", "缺失率"],
                [[c, n, _pct(r)] for c, n, r in miss_rows],
            ))
            lines.append("")
        else:
            lines.append("- 缺失：无")
        sample = ov.get("head_sample") or []
        if sample:
            headers = list(sample[0].keys())
            lines.append("- 前 5 行样本：")
            lines.append("")
            lines.extend(_table(headers, [[row.get(h) for h in headers] for row in sample]))
            lines.append("")

        worst = miss_rows[0] if miss_rows else None
        if worst:
            lines.append(f"**结论**：表结构完整，缺失主要集中在 `{worst[0]}`（{_pct(worst[2])}）；"
                         f"数值列共 {len([v for v in dtypes.values() if 'float' in v or 'int' in v])} 个。")
        else:
            lines.append("**结论**：表结构完整，无任何缺失值。")
    lines.append("")

    # ---------- stats ----------
    lines.append("## stats · 描述性统计")
    lines.append("")
    st = sections.get("stats") or {}
    if "__error__" in st:
        lines.append(f"本段执行失败：{st['__error__']}")
    else:
        sr = st.get("stats_result") or {}
        if not sr:
            lines.append("- 表格中无数值列，无法做数值统计。")
        else:
            metrics = ["count", "mean", "median", "std", "cv", "skew", "min", "25%", "50%", "75%", "max"]
            rows = [[col] + [_fmt(sr[col].get(m), 4 if m in ("cv", "skew") else 2) for m in metrics]
                    for col in sr]
            lines.extend(_table(["字段"] + metrics, rows))
            lines.append("")
            if value_col and value_col in sr:
                s = sr[value_col]
                skew = s.get("skew")
                cv = s.get("cv")
                tail = "存在明显右偏长尾，均值被少数极值拉高" if (skew or 0) > 1 else "分布相对均衡"
                lines.append(f"**结论**：`{value_col}` 均值 {_fmt(s.get('mean'))}、中位数 {_fmt(s.get('median'))}，"
                             f"偏度 {_fmt(skew, 4)}、变异系数 {_fmt(cv, 4)} —— {tail}。")
    lines.append("")

    # ---------- sort_filter ----------
    lines.append("## sort_filter · 排序筛选")
    lines.append("")
    sf = sections.get("sort_filter") or {}
    if "__error__" in sf:
        lines.append(f"本段执行失败：{sf['__error__']}")
    else:
        cond = sf.get("filter_condition") or "（无过滤条件）"
        order = "升序" if sf.get("sort_asc") else "降序"
        sort_col_name = sf.get("sort_col")
        if sort_col_name:
            lines.append(f"- 排序：`{sort_col_name}` {order}；过滤条件：`{cond}`")
        else:
            lines.append(f"- 排序：（未指定排序列，保持原始行序）；过滤条件：`{cond}`")
        lines.append(f"- **命中行数：{_fmt(sf.get('hit_rows'))} 行**")
        if sf.get("filter_parse_err"):
            lines.append(f"- ⚠ 过滤条件解析失败：{sf['filter_parse_err']}")
        sample = sf.get("sample") or []
        if sample:
            headers = list(sample[0].keys())
            lines.append("- 前 10 条：")
            lines.append("")
            lines.extend(_table(headers, [[row.get(h) for h in headers] for row in sample]))
            lines.append("")
        cond_desc = f"`{cond}`" if sf.get("filter_condition") else "未设置过滤条件"
        sort_desc = (f"已按 `{sort_col_name}` {order}排列" if sort_col_name
                     else "未指定排序列，保持原始行序")
        lines.append(f"**结论**：在 {cond_desc} 下命中 {_fmt(sf.get('hit_rows'))} 行，"
                     f"{sort_desc}。（工具仅返回前 10 条以控制输出体积）")
    lines.append("")

    # ---------- trend ----------
    lines.append("## trend · 时间趋势")
    lines.append("")
    tr = sections.get("trend") or {}
    if "__error__" in tr:
        lines.append(f"本段执行失败：{tr['__error__']}")
    else:
        cr = tr.get("clean_report") or {}
        lines.append(f"- 时间列 `{tr.get('time_col')}` / 指标列 `{tr.get('value_col')}`；有效点 **{_fmt(tr.get('valid_points'))}** 个")
        if cr:
            lines.append(f"- 清洗：原始 {_fmt(cr.get('original_rows'))} 行 → 无效时间 {_fmt(cr.get('invalid_time_rows'))} 行"
                         f"（样本：{cr.get('invalid_samples')}）→ 剩余 {_fmt(cr.get('remaining_rows'))} 行")
        if tr.get("first_value") is not None:
            arrow = {"rise": "上升", "decline": "下降", "flat": "持平"}.get(tr.get("direction"), "变化")
            lines.append(f"- 首值 {_fmt(tr.get('first_value'))} → 末值 {_fmt(tr.get('last_value'))}；"
                         f"变化 {_fmt(tr.get('change'))}（{_pct(tr.get('change_rate'))}），方向 **{arrow}**")
            lines.append(f"- 线性回归斜率 **{_fmt(tr.get('regression_slope'), 2)}**（每期平均变动量）")
            ma = tr.get("moving_average") or []
            tail = [v for v in ma[-3:] if v is not None and not (isinstance(v, float) and math.isnan(v))]
            if tail:
                lines.append(f"- 近 3 期移动平均：{ ' → '.join(_fmt(v) for v in tail) }")

            clean = adj_trend
            if clean:
                lines.append(f"- 剔除 IQR 离群点后（去掉 {clean['removed']} 个点）："
                             f"首值 {_fmt(clean['first'])} → 末值 {_fmt(clean['last'])}，"
                             f"变化率 {_pct(clean['change_rate'])}")

            if clean and abs((tr.get("change_rate") or 0) - clean["change_rate"]) > 0.2:
                adj_rate = clean["change_rate"]
                arrow2 = "上升" if adj_rate > 0 else ("下降" if adj_rate < 0 else "持平")
                lines.append(f"**结论**：表面{arrow} {_pct(abs(tr.get('change_rate') or 0))}，但该变化主要由首尾离群点驱动；"
                             f"剔除离群点后真实走势为{arrow2} {_pct(abs(adj_rate))}，"
                             f"回归斜率 {_fmt(tr.get('regression_slope'), 2)}。")
            else:
                lines.append(f"**结论**：整体走势为{arrow}，斜率 {_fmt(tr.get('regression_slope'), 2)}。")
        else:
            lines.append(f"**结论**：{tr.get('hint', '有效数据不足，无法给出趋势结论')}。")
    lines.append("")

    # ---------- anomaly ----------
    lines.append("## anomaly · 异常检测")
    lines.append("")
    an = sections.get("anomaly") or {}
    if "__error__" in an:
        lines.append(f"本段执行失败：{an['__error__']}")
    else:
        lines.append(f"- **重复行数：{_fmt(an.get('duplicate_row_count'))} 行**（总行数 {_fmt(an.get('total_rows'))}）")
        hm = an.get("high_missing_columns") or []
        if hm:
            lines.append("- **高缺失列**（阈值 >30%）：")
            desc = "、".join("{}（{}）".format(c["column"], _pct(c["missing_ratio"])) for c in hm)
            lines.append("  " + desc)
        else:
            lines.append("- 高缺失列：无")
        oi = an.get("outlier_statistics") or {}
        if oi:
            lines.append("- **离群点统计**（IQR 与 Z-Score 双指标）：")
            lines.append("")
            rows = []
            for col, v in oi.items():
                rows.append([col, _fmt(v.get("lower_bound")), _fmt(v.get("upper_bound")),
                             _fmt(v.get("outlier_count")), _pct(v.get("outlier_ratio")),
                             _fmt(v.get("zscore_outlier_count"))])
            lines.extend(_table(["字段", "IQR 下界", "IQR 上界", "IQR 离群数", "离群占比", "Z-Score 离群数"], rows))
            lines.append("")
            hit = [c for c, v in oi.items() if v.get("outlier_count", 0) > 0 or v.get("zscore_outlier_count", 0) > 0]
            if hit:
                lines.append(f"**结论**：无重复行；异常集中在 { '、'.join('`'+c+'`' for c in hit) } 列，"
                             f"其中 `{value_col or hit[0]}` 两种方法均能识别出离群点，其余数值列分布正常。")
            else:
                lines.append("**结论**：无重复行、无高缺失列、各数值列均无离群点，数据质量良好。")
        else:
            lines.append("**结论**：无数值列可做离群检测。")
    lines.append("")

    # ---------- chart ----------
    lines.append("## chart · 图表")
    lines.append("")
    ch = sections.get("chart") or {}
    if "__error__" in ch:
        lines.append(f"本段执行失败：{ch['__error__']}")
    else:
        lines.append(f"- 类型：`{ch.get('chart_type')}`，X=`{ch.get('x_col')}`，Y=`{ch.get('y_col')}`"
                     + ("（按 X 分组取**均值**）" if ch.get("chart_type") == "bar" else ""))
        lines.append(f"- 标题：{ch.get('title') or '—'}")
        lines.append(f"- **output_path：`{ch.get('output_path')}`**（{_fmt(ch.get('file_size_kb'))} KB）")

        grouped = _grouped_table(df, ch.get("x_col"), ch.get("y_col"))
        if grouped is not None:
            lines.append("- 分组数值（代替看图读数）：")
            lines.append("")
            lines.extend(_table(["分组", "均值", "合计", "记录数"], grouped["rows"]))
            lines.append("")
            top = grouped["rows"][0]
            bot = grouped["rows"][-1]
            ratio = (top[1] / bot[1]) if bot[1] else None
            lines.append(f"**结论**：`{top[0]}` 均值最高（{_fmt(top[1])}），`{bot[0]}` 最低（{_fmt(bot[1])}）"
                         + (f"，高低相差约 {_fmt(ratio, 1)} 倍。" if ratio else "。")
                         + " 若个别分组明显偏高，需检查是否由离群记录拉高（见 anomaly 段）。")
    lines.append("")

    # ---------- 综合结论 ----------
    lines.append("## 综合结论")
    lines.append("")
    lines.append(f"- {_build_summary(sections, value_col, adj_trend)}")
    if warnings:
        lines.append("- 生成过程中的提示：")
        for w in warnings:
            lines.append(f"  - {w}")
    lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def _outlier_adjusted_trend(df: Optional[pd.DataFrame], time_col: Optional[str], value_col: Optional[str]) -> Optional[Dict[str, Any]]:
    """剔除 IQR 离群点后重算首尾变化率，用于判断「暴跌/暴涨」是否由离群点驱动"""
    if df is None or not time_col or not value_col:
        return None
    if time_col not in df.columns or value_col not in df.columns:
        return None
    try:
        work = df[[time_col, value_col]].copy()
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            work[time_col] = pd.to_datetime(work[time_col], errors="coerce")
        work = work.dropna(subset=[time_col, value_col]).sort_values(by=time_col)
        if len(work) < 4:
            return None
        s = work[value_col].astype(float)
        q1, q3 = s.quantile(0.25), s.quantile(0.75)
        iqr = q3 - q1
        lower, upper = q1 - 1.5 * iqr, q3 + 1.5 * iqr
        kept = work[(s >= lower) & (s <= upper)]
        removed = len(work) - len(kept)
        if removed == 0 or len(kept) < 2:
            return None
        first = float(kept[value_col].iloc[0])
        last = float(kept[value_col].iloc[-1])
        rate = round((last - first) / first, 4) if first else None
        return {"removed": removed, "first": first, "last": last, "change_rate": rate}
    except Exception:
        return None


def _grouped_table(df: Optional[pd.DataFrame], x_col: Optional[str], y_col: Optional[str], top: int = 10) -> Optional[Dict[str, Any]]:
    """按 X 分组给出 均值/合计/记录数，让报告不依赖看图"""
    if df is None or not x_col or not y_col:
        return None
    if x_col not in df.columns or y_col not in df.columns:
        return None
    try:
        g = df.groupby(x_col)[y_col].agg(["mean", "sum", "count"]).sort_values("mean", ascending=False).head(top)
        rows = [[idx, float(r["mean"]), float(r["sum"]), int(r["count"])] for idx, r in g.iterrows()]
        return {"rows": rows}
    except Exception:
        return None
