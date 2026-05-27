import pandas as pd
import numpy as np

# 1. 在路径字符串前面加上字母 r，防止 \T 被错误转义 (解决 SyntaxWarning)
file_path = r"D:\This computer\desktop\失信数据汇总_含财务数据3 - 副本.xlsx"

# 2. 使用 pd.read_excel() 来读取 .xlsx 文件 (解决 UnicodeDecodeError)
df = pd.read_excel(file_path)

# 统计有多少家公司（依据UID，相同UID记为一家）
num_companies = df['UID'].nunique()
print(f"公司总数（基于UID）: {num_companies}\n")

# 统计每个列中有多少数据（空值不计算，0.0001视为空值）
valid_counts = {}

for col in df.columns:
    s = df[col]
    # 判断是否为有效值
    is_valid = s.notna() & (s != 0.0001) & (s.astype(str).str.strip() != '0.0001')
    valid_counts[col] = is_valid.sum()

valid_counts_series = pd.Series(valid_counts)

# 打印结果
pd.set_option('display.max_rows', None)
print("各列有效数据统计：")
print(valid_counts_series)