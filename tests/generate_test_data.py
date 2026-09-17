"""
generate_test_data.py
生成AI Excel分析测试样本数据（含脏数据场景）
运行：python tests/generate_test_data.py
输出：tests/test_data/sample_test.xlsx
"""
import os
import pandas as pd
import numpy as np

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "test_data")
OUTPUT_FILE = os.path.join(OUTPUT_DIR, "sample_test.xlsx")


def generate_sales_sheet() -> pd.DataFrame:
    """主Sheet：销售数据，覆盖 overview/stats/sort_filter/trend/anomaly/chart 全部场景"""
    dates = pd.date_range(start="2026-01-01", periods=30, freq="D")
    np.random.seed(42)

    region_list = ["华北", "华东", "华南", "西南", "华中"]
    data = {
        "日期": dates,
        "产品名称": [f"产品_{i % 5}" for i in range(30)],
        "销售额": np.random.randint(100, 5000, size=30).astype(float),
        "销量": np.random.randint(1, 200, size=30).astype(float),
        "单价": np.random.uniform(10, 500, size=30).round(2),
        "区域": [region_list[i % 5] for i in range(30)],
        "备注": ["正常" if i % 7 != 0 else "" for i in range(30)],
    }
    df = pd.DataFrame(data)

    # 注入缺失值：销售额缺失3条
    df.loc[df.sample(3, random_state=1).index, "销售额"] = np.nan
    # 注入高缺失列：备注缺失约40%
    df.loc[df.sample(12, random_state=2).index, "备注"] = np.nan
    # 注入IQR离群点：销售额加2个极端大值
    df.loc[0, "销售额"] = 99999.0
    df.loc[1, "销售额"] = 88888.0
    # 注入重复行
    df = pd.concat([df, df.iloc[[5]]], ignore_index=True)
    # 注入时间格式异常行
    df.loc[len(df) - 1, "日期"] = "not_a_date"

    return df


def generate_text_only_sheet() -> pd.DataFrame:
    """辅助Sheet：纯文本列，测试stats无数值列边界"""
    data = {
        "姓名": ["张三", "李四", "王五", "赵六"],
        "部门": ["研发", "市场", "销售", "人事"],
        "入职日期": ["2023-01-01", "2024-03-15", "2025-06-20", "2026-02-10"],
    }
    return pd.DataFrame(data)


def generate_multi_bad_time_sheet() -> pd.DataFrame:
    """辅助Sheet：多行时间脏数据，测试 trend auto_clean"""
    data = {
        "日期": ["2026-01-01", "2026-01-02", "bad1", "2026-01-04", "bad2", "2026-01-06"],
        "指标": [10.0, 20.0, 30.0, 40.0, 50.0, 60.0],
    }
    return pd.DataFrame(data)


def main():
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR, exist_ok=True)

    df_sales = generate_sales_sheet()
    df_text = generate_text_only_sheet()
    df_badtime = generate_multi_bad_time_sheet()

    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        df_sales.to_excel(writer, sheet_name="销售数据", index=False)
        df_text.to_excel(writer, sheet_name="员工信息", index=False)
        df_badtime.to_excel(writer, sheet_name="脏时间数据", index=False)

    print(f"✅测试Excel已生成: {OUTPUT_FILE}")
    print(f"  - Sheet1 [销售数据]: {df_sales.shape[0]}行 x {df_sales.shape[1]}列")
    print(f"  - Sheet2 [员工信息]: {df_text.shape[0]}行 x {df_text.shape[1]}列（纯文本）")
    print(f"  - Sheet3 [脏时间数据]: {df_badtime.shape[0]}行 x {df_badtime.shape[1]}列（多行脏时间）")


if __name__ == "__main__":
    main()