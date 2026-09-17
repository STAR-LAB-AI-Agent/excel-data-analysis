"""
data_cleaner.py
数据清洗工具：处理时间列脏数据、缺失值策略
只抛出自定义业务异常，不处理输出、日志打印
"""
from typing import Tuple, Optional
import pandas as pd
from src.excel_analyzer.exceptions import DataTypeError


def clean_time_column(
    df: pd.DataFrame,
    time_col: str,
    drop_invalid: bool = True
) -> Tuple[pd.DataFrame, dict]:
    """
    清洗时间列：
    - 尝试将时间列转为 datetime
    - drop_invalid=True 时剔除无法解析的行，而非整体报错
    - 返回 (清洗后df, 清洗报告)
    """
    work_df = df.copy()
    original_len = len(work_df)

    # 尝试转换；errors="coerce" 将无法解析的值置为 NaT
    work_df[time_col] = pd.to_datetime(work_df[time_col], errors="coerce")

    invalid_mask = work_df[time_col].isna()
    invalid_count = int(invalid_mask.sum())
    invalid_samples = (
        df.loc[invalid_mask, time_col].head(5).astype(str).tolist()
        if invalid_count > 0 else []
    )

    report = {
        "original_rows": original_len,
        "invalid_time_rows": invalid_count,
        "invalid_samples": invalid_samples,
        "dropped_invalid": False,
    }

    if invalid_count > 0 and drop_invalid:
        work_df = work_df[~invalid_mask].reset_index(drop=True)
        report["dropped_invalid"] = True
        report["remaining_rows"] = len(work_df)

    # 若全部无法解析，则抛出异常（无法继续分析）
    if len(work_df) == 0:
        raise DataTypeError(
            f"时间列[{time_col}]全部无法解析为时间格式，样本：{invalid_samples}"
        )

    return work_df, report


def fill_missing_numeric(
    df: pd.DataFrame,
    strategy: str = "median"
) -> Tuple[pd.DataFrame, dict]:
    """
    数值列缺失值填充
    strategy: mean / median / zero / drop
    """
    work_df = df.copy()
    numeric_cols = work_df.select_dtypes(include=["number"]).columns.tolist()
    report = {"strategy": strategy, "filled": {}, "dropped_rows": 0}

    if strategy == "drop":
        before = len(work_df)
        work_df = work_df.dropna(subset=numeric_cols).reset_index(drop=True)
        report["dropped_rows"] = before - len(work_df)
        return work_df, report

    for col in numeric_cols:
        missing = int(work_df[col].isna().sum())
        if missing == 0:
            continue
        if strategy == "mean":
            fill_val = work_df[col].mean()
        elif strategy == "median":
            fill_val = work_df[col].median()
        elif strategy == "zero":
            fill_val = 0
        else:
            raise ValueError(f"不支持的填充策略: {strategy}")

        work_df[col] = work_df[col].fillna(fill_val)
        report["filled"][str(col)] = {
            "missing_count": missing,
            "fill_value": float(fill_val) if pd.notna(fill_val) else None,
        }

    return work_df, report