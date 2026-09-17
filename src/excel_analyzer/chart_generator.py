"""
chart_generator.py
图表生成模块（可选功能）
支持 bar / line / box / hist 四种图表
输出到 outputs/ 目录，不修改原始 Excel
"""
import os
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, List
import pandas as pd
import matplotlib

# 使用非交互后端，避免在无 GUI 环境下报错
matplotlib.use("Agg")
import matplotlib.pyplot as plt

from src.excel_analyzer.exceptions import (
    ColumnNotFoundError,
    ChartGenFailedError,
    NoNumericColumnError,
)

# 中文字体配置（Windows / Linux 兼容）
plt.rcParams["font.sans-serif"] = ["SimHei", "Microsoft YaHei", "DejaVu Sans"]
plt.rcParams["axes.unicode_minus"] = False


def _get_output_dir() -> Path:
    """图表输出目录：项目根目录/outputs"""
    current = Path(__file__).resolve()
    for _ in range(5):
        if (current / "tests").exists():
            root = current
            break
        current = current.parent
    else:
        root = Path.cwd()
    out_dir = root / "outputs"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir


def _safe_filename(name: str) -> str:
    """将列名等转为安全文件名"""
    return "".join(c if c.isalnum() or c in "-_" else "_" for c in str(name))


def generate_chart(
    df: pd.DataFrame,
    chart_type: str,
    x_col: Optional[str] = None,
    y_col: Optional[str] = None,
    title: Optional[str] = None,
) -> Dict[str, Any]:
    """
    生成图表
    chart_type: bar / line / box / hist
    x_col: 分类轴或时间轴（bar/line 必填）
    y_col: 数值轴（bar/line/hist 必填；box 可多列）
    """
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if len(numeric_cols) == 0:
        raise NoNumericColumnError("表格中无数值列，无法生成图表")

    # 参数校验
    if chart_type in ("bar", "line"):
        if not x_col or not y_col:
            raise ColumnNotFoundError("bar/line 图需要同时提供 --x_col 和 --y_col")
        if x_col not in df.columns:
            raise ColumnNotFoundError(f"x 轴列[{x_col}]不存在")
        if y_col not in df.columns:
            raise ColumnNotFoundError(f"y 轴列[{y_col}]不存在")
        if y_col not in numeric_cols:
            raise NoNumericColumnError(f"y 轴列[{y_col}]不是数值类型")

    if chart_type == "hist" and (not y_col or y_col not in numeric_cols):
        raise ColumnNotFoundError(f"hist 图需要提供数值列 --y_col，当前：{y_col}")

    out_dir = _get_output_dir()
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    base_name = f"{chart_type}_{_safe_filename(y_col or 'multi')}_{ts}"
    out_path = out_dir / f"{base_name}.png"

    try:
        fig, ax = plt.subplots(figsize=(10, 6))

        if chart_type == "bar":
            grouped = df.groupby(x_col)[y_col].mean().sort_values(ascending=False).head(20)
            grouped.plot(kind="bar", ax=ax, color="#4C72B0")
            ax.set_xlabel(str(x_col))
            ax.set_ylabel(f"{y_col} (均值)")
            ax.set_title(title or f"{y_col} 按 {x_col} 分布（Top20）")
            plt.xticks(rotation=45, ha="right")

        elif chart_type == "line":
            # ✨ 修复：先处理 x 轴类型，避免混合类型排序报错
            plot_df = df[[x_col, y_col]].dropna().copy()

            # 尝试将 x 轴转为 datetime
            if not pd.api.types.is_datetime64_any_dtype(plot_df[x_col]):
                try:
                    converted = pd.to_datetime(plot_df[x_col], errors="coerce")
                    if converted.notna().all():
                        # 全部成功转 datetime
                        plot_df[x_col] = converted
                    else:
                        # 有脏数据 → 统一转字符串
                        plot_df[x_col] = plot_df[x_col].astype(str)
                except Exception:
                    plot_df[x_col] = plot_df[x_col].astype(str)
            else:
                # 已是 datetime → 剔除 NaT
                plot_df = plot_df.dropna(subset=[x_col])

            # 排序：根据 x 轴类型选择排序方式
            if pd.api.types.is_datetime64_any_dtype(plot_df[x_col]):
                plot_df = plot_df.sort_values(by=x_col)
            else:
                plot_df = plot_df.sort_values(by=x_col, key=lambda s: s.astype(str))

            ax.plot(plot_df[x_col], plot_df[y_col], marker="o", color="#DD8452")
            ax.set_xlabel(str(x_col))
            ax.set_ylabel(str(y_col))
            ax.set_title(title or f"{y_col} 随 {x_col} 变化趋势")
            plt.xticks(rotation=45, ha="right")

        elif chart_type == "box":
            cols = [y_col] if y_col else numeric_cols[:5]
            cols = [c for c in cols if c in numeric_cols]
            if not cols:
                raise ColumnNotFoundError("box 图无数值列可绘制")
            df[cols].boxplot(ax=ax)
            ax.set_title(title or "数值列箱线图")
            ax.set_ylabel("数值")

        elif chart_type == "hist":
            df[y_col].dropna().plot(kind="hist", bins=30, ax=ax, color="#55A868", edgecolor="white")
            ax.set_xlabel(str(y_col))
            ax.set_ylabel("频数")
            ax.set_title(title or f"{y_col} 分布直方图")

        else:
            raise ChartGenFailedError(f"不支持的图表类型: {chart_type}")

        plt.tight_layout()
        fig.savefig(out_path, dpi=120, bbox_inches="tight")
        plt.close(fig)

    except (ColumnNotFoundError, NoNumericColumnError, ChartGenFailedError):
        raise
    except Exception as e:
        raise ChartGenFailedError(f"图表生成失败: {str(e)}") from e

    return {
        "chart_type": chart_type,
        "x_col": x_col,
        "y_col": y_col,
        "output_path": str(out_path),
        "file_size_kb": round(os.path.getsize(out_path) / 1024, 2),
        "title": title or "",
    }