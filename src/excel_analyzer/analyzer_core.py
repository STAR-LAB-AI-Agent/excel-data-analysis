"""
analyzer_core.py
Excel分析核心业务逻辑，6类意图实现（含 chart）
只抛出自定义业务异常，不处理输出、日志打印
"""
import os
from typing import Dict, Any, Optional, List, Tuple
import pandas as pd
import numpy as np

from src.excel_analyzer.schemas import ALLOWED_EXCEL_SUFFIX
from src.excel_analyzer.exceptions import (
    FileNotFoundError,
    InvalidSuffixError,
    SheetNotExistError,
    ColumnNotFoundError,
    DataTypeError,
    ParseExcelFailedError,
    NoNumericColumnError,
)
from src.excel_analyzer.data_cleaner import clean_time_column


def load_excel(file_path: str, sheet_name: int | str = 0) -> pd.DataFrame:
    """加载excel，前置校验，返回DataFrame"""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"文件不存在：{file_path}")

    _, ext = os.path.splitext(file_path)
    ext = ext.lower()
    if ext not in ALLOWED_EXCEL_SUFFIX:
        raise InvalidSuffixError(f"不支持文件后缀 {ext}, 仅支持 {ALLOWED_EXCEL_SUFFIX}")

    try:
        excel_file = pd.ExcelFile(file_path)
    except Exception as e:
        raise ParseExcelFailedError(f"Excel文件解析失败: {str(e)}") from e

    if sheet_name not in excel_file.sheet_names:
        if isinstance(sheet_name, int):
            if sheet_name >= len(excel_file.sheet_names):
                raise SheetNotExistError(
                    f"索引{sheet_name}对应的Sheet不存在，当前sheet列表：{excel_file.sheet_names}"
                )
        else:
            raise SheetNotExistError(
                f"Sheet名称[{sheet_name}]不存在，当前sheet列表：{excel_file.sheet_names}"
            )

    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
    except Exception as e:
        raise ParseExcelFailedError(f"读取sheet失败: {str(e)}") from e

    return df


# ============================ I1 概览 ============================
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
            "missing_ratio": missing_ratio,
        }

    head_sample = df.head(5).to_dict(orient="records")
    return {
        "row_count": row_count,
        "col_count": col_count,
        "columns": col_names,
        "dtypes": col_dtypes,
        "missing_stats": missing_stats,
        "head_sample": head_sample,
    }


# ============================ I2 统计 ============================
def do_stats(df: pd.DataFrame) -> Dict[str, Any]:
    """I2 描述性统计：数值列；增加偏度与变异系数"""
    numeric_df = df.select_dtypes(include=["number"])
    numeric_cols = list(numeric_df.columns)
    non_numeric_cols = list(set(df.columns) - set(numeric_cols))

    if len(numeric_cols) == 0:
        return {
            "hint": "表格中无数值类型字段，无法做数值统计",
            "numeric_columns": [],
            "non_numeric_columns": non_numeric_cols,
            "stats_result": {},
        }

    stats_result = {}
    for col in numeric_cols:
        s = numeric_df[col].dropna()
        mean_val = float(s.mean()) if len(s) > 0 else None
        std_val = float(s.std()) if len(s) > 0 else None
        # 变异系数：标准差/均值，衡量相对离散度
        cv = round(std_val / mean_val, 4) if (mean_val not in (None, 0) and std_val is not None) else None
        # 偏度：>0 右偏，<0 左偏
        skew = round(float(s.skew()), 4) if len(s) > 2 else None

        stats_result[str(col)] = {
            "count": int(s.count()),
            "mean": mean_val,
            "median": float(s.median()) if len(s) > 0 else None,
            "std": std_val,
            "cv": cv,
            "skew": skew,
            "min": float(s.min()) if len(s) > 0 else None,
            "25%": float(s.quantile(0.25)) if len(s) > 0 else None,
            "50%": float(s.quantile(0.50)) if len(s) > 0 else None,
            "75%": float(s.quantile(0.75)) if len(s) > 0 else None,
            "max": float(s.max()) if len(s) > 0 else None,
        }

    return {
        "numeric_columns": numeric_cols,
        "non_numeric_columns": non_numeric_cols,
        "stats_result": stats_result,
    }


# ============================ I3 排序筛选 ============================
def do_sort_filter(
    df: pd.DataFrame,
    sort_col: Optional[str] = None,
    sort_asc: bool = False,
    filter_condition: Optional[str] = None,
) -> Dict[str, Any]:
    """I3 排序筛选；过滤条件使用 pd.eval 安全限制"""
    work_df = df.copy()

    if sort_col is not None:
        if sort_col not in work_df.columns:
            raise ColumnNotFoundError(f"排序字段[{sort_col}]不存在")
        work_df = work_df.sort_values(by=sort_col, ascending=sort_asc).reset_index(drop=True)

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
        "sample": sample_out,
    }


# ============================ I4 趋势 ============================
def _linear_regression_slope(y: np.ndarray) -> Optional[float]:
    """计算线性回归斜率，用于判断整体趋势方向与强度"""
    if len(y) < 2:
        return None
    x = np.arange(len(y), dtype=float)
    try:
        slope = float(np.polyfit(x, y, 1)[0])
        return round(slope, 4)
    except Exception:
        return None


def do_trend(
    df: pd.DataFrame,
    time_col: str,
    value_col: str,
    auto_clean: bool = True,
) -> Dict[str, Any]:
    """
    I4 趋势分析（增强版）
    - auto_clean=True 时自动剔除无法解析的时间行
    - 增加移动平均、线性回归斜率
    """
    if time_col not in df.columns:
        raise ColumnNotFoundError(f"时间列[{time_col}]不存在")
    if value_col not in df.columns:
        raise ColumnNotFoundError(f"指标数值列[{value_col}]不存在")

    work_df = df.copy()
    clean_report = {}
    if auto_clean:
        work_df, clean_report = clean_time_column(work_df, time_col, drop_invalid=True)
    else:
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
            "valid_points": len(valid_data),
            "clean_report": clean_report,
        }

    values = valid_data[value_col].astype(float).values
    first_val = float(values[0])
    last_val = float(values[-1])
    change = last_val - first_val
    change_rate = round(change / first_val, 4) if first_val != 0 else None
    direction = "rise" if change > 0 else ("decline" if change < 0 else "flat")

    # 移动平均（窗口=3，不足则用全部）
    window = min(3, len(values))
    ma = pd.Series(values).rolling(window=window).mean().round(4).tolist()

    # 线性回归斜率
    slope = _linear_regression_slope(values)

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
        "moving_average": ma,
        "regression_slope": slope,
        "clean_report": clean_report,
        "sample_points": sample_points,
    }


# ============================ I5 异常检测 ============================
def _zscore_outliers(series: pd.Series, threshold: float = 3.0) -> int:
    """Z-Score 离群点计数"""
    s = series.dropna()
    if len(s) < 3:
        return 0
    std = s.std()
    if std == 0:
        return 0
    z = (s - s.mean()) / std
    return int((z.abs() > threshold).sum())


def do_anomaly(df: pd.DataFrame) -> Dict[str, Any]:
    """
    I5 异常检测（增强版）
    - 重复行
    - 高缺失列（>0.3）
    - IQR 离群点 + Z-Score 离群点双指标
    """
    row_total = len(df)
    duplicate_count = int(df.duplicated().sum())

    high_missing_cols = []
    for col in df.columns:
        miss_ratio = round(df[col].isna().sum() / row_total if row_total > 0 else 0, 4)
        if miss_ratio > 0.3:
            high_missing_cols.append({"column": str(col), "missing_ratio": miss_ratio})

    numeric_df = df.select_dtypes(include=["number"])
    outlier_info = {}
    for col in numeric_df.columns:
        series = df[col].dropna()
        if len(series) == 0:
            continue
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
            "outlier_ratio": outlier_ratio,
            "zscore_outlier_count": _zscore_outliers(series),  # ✨ 新增
        }

    return {
        "total_rows": row_total,
        "duplicate_row_count": duplicate_count,
        "high_missing_columns": high_missing_cols,
        "outlier_statistics": outlier_info,
    }


# ============================ I6 图表 ============================
def do_chart(
    df: pd.DataFrame,
    chart_type: str,
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
    title: Optional[str] = None,
) -> Dict[str, Any]:
    """I6 图表生成（可选功能）"""
    from src.excel_analyzer.chart_generator import generate_chart
    return generate_chart(df, chart_type, x_col=x_col, y_col=y_col, title=title)