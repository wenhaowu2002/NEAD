import pandas as pd
import scipy.signal
import xlsxwriter

# 输入文件名
input_file = 'raw_fi.xlsx'
output_file = 'output_median.xlsx'

# 读取Excel文件中的数据
df = pd.read_excel(input_file, header=None)  # 假设数据在Excel文件的第一个工作表中

# 中值滤波窗口的大小（可根据需要调整）
window_size = 5

# 创建一个新的Excel文件，设置nan_inf_to_errors选项
workbook = xlsxwriter.Workbook(output_file, {'nan_inf_to_errors': True})
worksheet = workbook.add_worksheet()

# 对每一行数据进行中值滤波处理，并逐一写入Excel文件的不同单元格
for row_idx, row_data in enumerate(df.values):
    smoothed_row_data = scipy.signal.medfilt(row_data, kernel_size=window_size)
    for col_idx, cell_value in enumerate(smoothed_row_data):
        worksheet.write(row_idx, col_idx, cell_value)

# 关闭Excel文件
workbook.close()

print("平滑数据已保存到", output_file)
