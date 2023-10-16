import pandas as pd
import numpy as np

# 输入文件名
input_file = 'raw_fi.xlsx'
output_file = 'output_moveaverage.xlsx'

# 读取Excel文件中的数据
df = pd.read_excel(input_file, header=None)  # 假设数据在Excel文件的第一个工作表中

# 移动平均窗口的大小（可根据需要调整）
window_size = 5

# 对每一行数据进行移动平均滤波处理
smoothed_data = df.apply(lambda x: x.rolling(window=window_size, min_periods=1).mean(), axis=1)

# 创建新的DataFrame来保存平滑后的数据
smoothed_df = pd.DataFrame(smoothed_data)

# 将平滑后的数据保存到新的Excel文件
smoothed_df.to_excel(output_file, index=False, header=False)

print("平滑数据已保存到", output_file)
