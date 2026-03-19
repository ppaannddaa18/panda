"""
发票处理工具示例

功能：批量读取Excel中的发票数据，进行汇总统计
"""

import pandas as pd
from pathlib import Path
from utils import format_amount, chinese_amount


def read_invoices(file_path):
    """读取发票数据"""
    df = pd.read_excel(file_path)
    return df


def summarize_invoices(df):
    """汇总发票金额"""
    if '金额' not in df.columns:
        print("警告：未找到'金额'列")
        return None

    total = df['金额'].sum()
    print(f"发票数量: {len(df)}")
    print(f"合计金额: {format_amount(total)}")
    print(f"大写金额: {chinese_amount(total)}")
    return total


if __name__ == "__main__":
    # 示例用法
    data_file = Path(__file__).parent.parent.parent / "data" / "invoices.xlsx"

    if data_file.exists():
        df = read_invoices(data_file)
        summarize_invoices(df)
    else:
        print(f"请将发票数据文件放到: {data_file}")