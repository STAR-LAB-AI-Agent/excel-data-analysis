"""
analyzer_core.py
Excel分析核心业务逻辑，5类意图实现
只抛出自定义业务异常，不处理输出、日志打印（日志由调用方处理）
"""
import os
from typing import Dict, Any, Optional, List
import pandas as pd
from src.excel_analyzer.schemas import ALLOWED_EXCEL_SUFFIX
from src.excel_analyzer.exceptions import (
    FileNotFoundError,
    InvalidSuffixError,
    SheetNotExistError,
    ColumnNotFoundError,
    DataTypeError,
    ParseExcelFailedError
)


def load_excel(file_path: str, sheet_name: int | str = 0) -> pd.DataFrame:
    """
    加载excel，前置校验，返回DataFrame
    :param file_path: 文件路径
    :param sheet_name: sheet索引/名称
    :return: pd.DataFrame
    """
    # 文件存在校验
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")

    # 后缀校验
    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    if ext not in ALLOWED_EXCEL_SUFFIX:
        raise InvalidSuffixError(f"不支持文件后缀 {ext}, 仅支持 {ALLOWED_EXCEL_SUFFIX}")

    # sheet有效性先校验
    try:
        excel_file = pd.ExcelFile(file_path)
    except Exception as e:
        raise ParseExcelFailedError(f"Excel文件解析失败: {str(e)}") from e

    if sheet_name not in excel_file.sheet_names:
        if isinstance(sheet_name, int):
            if sheet_name >= len(excel_file.sheet_names):
                raise SheetNotExistError(f"索引{sheet_name}对应的Sheet不存在，当前sheet列表：{excel_file.sheet_names}")
        else:
            raise SheetNotExistError(f"Sheet名称[{sheet_name}]不存在，当前sheet列表：{excel_file.sheet_names}")

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        raise ParseExcelFailedError(f"读取sheet失败: {str(e)}") from e

    return df


def do_overview(df: pd.DataFrame) -> Dict[str, Any]:
    """I1 数据概览：行列、字段、类型、缺失统计、前5行样本"""
    row_count, col_count = df.shape
    col_names = list(df.columns)
    col_dtypes = {str(col): str(dtype) for col, dtype in df.dtypes.items()}

    missing_stats = {}
    for col in df.columns:
        missing_cnt = int(df[col].isna().sum())
        missing_ratio = round(missing_cnt / len(df) if len(df) > 0 else 0, 4)
        missing_stats[str(col)] = {
            "missing_count": missing_cnt,
            "missing_ratio": missing_ratio
        }

    head_sample = df.head(5).to_dict(orient="records")
    return {
        "row_count": row_count,
        "col_count": col_count,
        "columns": col_names,
        "dtypes": col_dtypes,
        "missing_stats": missing_stats,
        "head_sample": head_sample
    }


def do_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """I2 描述性统计：只对数值列计算；非数值列跳过并提示"""
    numeric_df = df.select_dtypes(include=["number"])
    numeric_cols = list(numeric_df.columns)
    non_numeric_cols = list(set(df.columns) - set(numeric_cols))

    if len(numeric_cols) == 0:
        return {
            "hint": "表格中无数值类型字段，无法做数值统计",
            "numeric_columns": [],
            "non_numeric_columns": non_numeric_cols,
            "stats_result": {}
        }

    stats_df = numeric_df.describe()
    stats_result = {}
    for col in numeric_cols:
        s = stats_df[col]
        stats_result[str(col)] = {
            "count": float(s["count"]),
            "mean": float(s["mean"]),
            "median": float(numeric_df[col].median()),
            "std": float(s["std"]),
            "min": float(s["min"]),
            "25%": float(s["25%"]),
            "50%": float(s["50%"]),
            "75%": float(s["75%"]),
            "max": float(s["max"])
        }

    return {
        "numeric_columns": numeric_cols,
        "non_numeric_columns": non_numeric_cols,
        "stats_result": stats_result
    }


def do_sort_filter(
    df: pd.DataFrame,
    sort_col: Optional[str] = None,
    sort_asc: bool = False,
    filter_condition: Optional[str] = None
) -> Dict[str, Any]:
    """
    I3 排序筛选
    filter_condition: 简单表达式字符串例如 "销售额>1000"，使用pd.eval安全限制
    """
    work_df = df.copy()

    # 排序
    if sort_col is not None:
        if sort_col not in work_df.columns:
            raise ColumnNotFoundError(f"排序字段[{sort_col}]不存在")
        work_df = work_df.sort_values(by=sort_col, ascending=sort_asc).reset_index(drop=True)

    # 条件过滤
    filter_hit_rows = -1
    filter_parse_err = None
    if filter_condition is not None and filter_condition.strip() != "":
        try:
            mask = work_df.eval(filter_condition)
            work_df = work_df[mask].reset_index(drop=True)
        except Exception as e:
            filter_parse_err = f"过滤条件解析失败: {str(e)}"

    filter_hit_rows = len(work_df)

    sample_out = work_df.head(10).to_dict(orient="records")
    return {
        "sort_col": sort_col,
        "sort_asc": sort_asc,
        "filter_condition": filter_condition,
        "filter_parse_err": filter_parse_err,
        "hit_rows": filter_hit_rows,
        "sample": sample_out
    }


def do_trend(df: pd.DataFrame, time_col: str, value_col: str) -> Dict[str, Any]:
    """
    I4 趋势分析
    :param df: 原始df
    :param time_col: 时间字段
    :param value_col: 需要分析的数值指标字段
    """
    if time_col not in df.columns:
        raise ColumnNotFoundError(f"时间列[{time_col}]不存在")
    if value_col not in df.columns:
        raise ColumnNotFoundError(f"指标数值列[{value_col}]不存在")

    work_df = df.copy()
    # 尝试转为时间格式
    try:
        work_df[time_col] = pd.to_datetime(work_df[time_col])
    except Exception as e:
        raise DataTypeError(f"时间列[{time_col}]无法解析为时间格式：{str(e)}") from e

    if not pd.api.types.is_numeric_dtype(work_df[value_col]):
        raise DataTypeError(f"指标列[{value_col}]不是数值类型，不能做趋势计算")

    work_df = work_df.sort_values(by=time_col).reset_index(drop=True)
    valid_data = work_df.dropna(subset=[time_col, value_col])

    if len(valid_data) <= 1:
        return {
            "hint": "有效时间-指标数据不足2条，无法分析趋势",
            "time_col": time_col,
            "value_col": value_col,
            "valid_points": len(valid_data)
        }

    # 基础趋势指标
    first_val = float(valid_data[value_col].iloc[0])
    last_val = float(valid_data[value_col].iloc[-1])
    change = last_val - first_val
    change_rate = round(change / first_val, 4) if first_val != 0 else None
    direction = "rise" if change > 0 else ("decline" if change < 0 else "flat")

    sample_points = valid_data.head(10).to_dict(orient="records")
    return {
        "time_col": time_col,
        "value_col": value_col,
        "valid_points": len(valid_data),
        "first_value": first_val,
        "last_value": last_val,
        "change": change,
        "change_rate": change_rate,
        "direction": direction,
        "sample_points": sample_points
    }


def do_anomaly(df: pd.DataFrame) -> Dict[str, Any]:
    """
    I5 异常检测：缺失严重列、重复行、IQR离群点检测（仅数值列）
    """
    row_total = len(df)
    duplicate_count = int(df.duplicated().sum())

    # 缺失严重列：缺失率>0.3
    high_missing_cols = []
    for col in df.columns:
        miss_ratio = round(df[col].isna().sum() / row_total if row_total > 0 else 0, 4)
        if miss_ratio > 0.3:
            high_missing_cols.append({"column": str(col), "missing_ratio": miss_ratio})

    # IQR离群点
    numeric_df = df.select_dtypes(include=["number"])
    outlier_info = {}
    for col in numeric_df.columns:
        series = df[col].dropna()
        q1 = series.quantile(0.25)
        q3 = series.quantile(0.75)
        iqr = q3 - q1
        lower_bound = q1 - 1.5 * iqr
        upper_bound = q3 + 1.5 * iqr
        outliers = series[(series < lower_bound) | (series > upper_bound)]
        outlier_count = int(len(outliers))
        outlier_ratio = round(outlier_count / len(series) if len(series) > 0 else 0, 4)
        outlier_info[str(col)] = {
            "lower_bound": float(lower_bound),
            "upper_bound": float(upper_bound),
            "outlier_count": outlier_count,
            "outlier_ratio": outlier_ratio
        }

    return {
        "total_rows": row_total,
        "duplicate_row_count": duplicate_count,
        "high_missing_columns": high_missing_cols,
        "outlier_statistics": outlier_info
    }