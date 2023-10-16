import pandas as pd
from scipy.signal import savgol_filter
import xlsxwriter
# Savitzky-Golay滤波
# 输入文件名
input_file = 'raw_fi.xlsx'
output_file = 'output_SG.xlsx'

# 读取Excel文件中的数据
df = pd.read_excel(input_file, header=None)  # 假设数据在Excel文件的第一个工作表中

# 对每一行数据进行平滑处理
smoothed_data = df.apply(lambda row: savgol_filter(row, window_length=5, polyorder=2), axis=1)

# 创建新的Excel文件
workbook = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# 将平滑后的数据逐一写入Excel文件的不同单元格
for row_idx, row in enumerate(smoothed_data.values):
    for col_idx, value in enumerate(row):
        worksheet.write(row_idx, col_idx, value)

# 关闭Excel文件
workbook.close()

print("平滑数据已保存到", output_file)
