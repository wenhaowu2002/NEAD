import pandas as pd


# 读取原始数据
input_file = 'raw_data_country.xlsx'
xls = pd.ExcelFile(input_file)

# 创建一个ExcelWriter以写入多个工作簿
output_file = 'processed_data_country.xlsx'
writer = pd.ExcelWriter(output_file, engine='xlsxwriter')

# 遍历每个工作簿并处理数据
for sheet_name in xls.sheet_names:
    df = pd.read_excel(xls, sheet_name, header=None)

    # 提取第一行和第一列
    first_row = df.iloc[0, 1:]
    first_col = df.iloc[1:, 0]

    # 剔除第一行和第一列
    df = df.iloc[1:, 1:]
    # 定义异常数对检测函数
    def detect_outliers(row):
        outliers = []  # 存储异常数对的索引
        for i in range(1, len(row)):
            if row[i] == 0 or row[i - 1] == 0:
                continue  # 跳过包含0的数对
            else:
                ratio = row[i] / row[i - 1]
                if ratio > 2 or ratio < 1/2:
                    outliers.append(i)
        return outliers

    # 处理每一行数据并保存结果
    output_data = []

    for row_index in range(df.shape[0]):
        row_data = df.iloc[row_index, :].values.tolist()  # 获取当前行的数据
        print(row_data)
        raw_row_data = row_data[:]
        classifications = [-1] * len(row_data)  # 初始化分类数组，-1表示未分类

        for i in range(0, len(classifications) - 1):
            if row_data[i] is None:
                row_data[i] = 0
            if row_data[i] == '':
                row_data[i] = 0
            if row_data[i + 1] == '':
                row_data[i] = 0
            if row_data[i - 1] == '':
                row_data[i] = 0
            if row_data[i] != 0 and row_data[i - 1] != 0 and row_data[i + 1] != 0:
                ratio_left = row_data[i] / row_data[i - 1]
                print(row_data[i+1])
                print(row_data[i])
                ratio_right = row_data[i + 1] / row_data[i]
                if ratio_left > 3 and ratio_right < 1/3:
                    row_data[i] = (row_data[i - 1] + row_data[i + 1]) / 2
                elif ratio_left < 1/3 and ratio_right > 3:
                    row_data[i] = (row_data[i - 1] + row_data[i + 1]) / 2
                elif ratio_left > 3 and ratio_right > 3:
                    row_data[i] = row_data[i + 1]
                elif ratio_left < 1/3 and ratio_right < 1/3:
                    row_data[i] = row_data[i - 1]
        # 检测异常数对并进行分类
        outlier_indices = detect_outliers(row_data)
        print(outlier_indices)
        if len(outlier_indices) > 0:
            for i in range(0, len(outlier_indices)):
                if row_data[outlier_indices[i]] == 0 or row_data[outlier_indices[i] - 1] == 0:
                    classifications[outlier_indices[i]] = -1
                elif row_data[outlier_indices[i]] > row_data[outlier_indices[i] - 1]:
                    classifications[outlier_indices[i]] = 1
                    classifications[outlier_indices[i] - 1] = 0
                else:
                    classifications[outlier_indices[i]] = 0
                    classifications[outlier_indices[i] - 1] = 1
        print(row_data)
        # 区分由0值导致的-1和由不在非异常数对所导致的-1
        for i in range(0, len(classifications)):
            if row_data[i] == 0:
                classifications[i] = 2
        print(classifications)
        print(len(classifications))
        # 补全非异常数对中的分类
        if any(classification != -1 for classification in classifications):
            while -1 in classifications:
                for i in range(0, len(classifications)):
                    if i != 0 and i != len(classifications) - 1:
                        if classifications[i] == -1:
                            left_classification = classifications[i - 1]
                            right_classification = classifications[i + 1]
                            if left_classification != -1 and left_classification != 2:
                                classifications[i] = left_classification
                            elif right_classification != -1 and right_classification != 2:
                                classifications[i] = right_classification
                            elif left_classification == -1 and right_classification == -1:
                                classifications[i] = classifications[i]
                            else:
                                classifications[i] = 2
                        print(classifications)
                    elif i != len(classifications) - 1:
                        if classifications[i] == -1:
                            right_classification = classifications[i + 1]
                            if right_classification != -1 and right_classification != 2:
                                classifications[i] = right_classification
                            elif right_classification == -1:
                                classifications[i] = classifications[i]
                            else:
                                classifications[i] = 2
                    else:
                        if classifications[i] == -1:
                            left_classification = classifications[i - 1]
                            if left_classification != -1 and left_classification != 2:
                                classifications[i] = left_classification
                            elif left_classification == -1:
                                classifications[i] = classifications[i]
                            else:
                                classifications[i] = 2
        print("original class:", classifications)
        print(raw_row_data)
        # 统计各个分类的数量
        class_counts = {0: 0, 1: 0}
        for classification in classifications:
            if classification != 2 and classification != -1:
                class_counts[classification] += 1

        print(class_counts)
        # 确定优势和劣势类别
        dominant_class = max(class_counts, key=class_counts.get)
        inferior_class = 1 - dominant_class  # 假设只有0和1两个分类
        print(raw_row_data)
        first_inferior = None
        last_inferior = None
        first_dominant = None
        last_dominant = None

        # 提前找到最后一个优势值
        for i in range(1, len(classifications)):
            if classifications[i] == dominant_class:
                if first_dominant is None:
                    first_dominant = i
                last_dominant = i
        print("inferior class: ", inferior_class)

        #监测是不是存在15年前后的断层
        credit = 0

        for i in range(15, len(classifications)):
            if classifications[i] == inferior_class:
                credit += 1
        sum_ratio = 0
        gap_ratio = 0
        if row_data[15] != 0 and row_data[16] != 0 and row_data[17] != 0 and row_data[18] != 0:
            ratio_2 = row_data[16] / row_data[15]
            ratio_3 = row_data[17] / row_data[16]
            ratio_4 = row_data[18] / row_data[17]
            ratio_5 = row_data[19] / row_data[18]
            sum_ratio = ratio_2 + ratio_3 + ratio_4 + ratio_5
        if row_data[14] != 0:
            gap_ratio = row_data[15] / row_data[14]

        last_five_elements = row_data[-5:]
        total_last_five = sum(last_five_elements)
        average_newdata = total_last_five / 5

        other_elements = row_data[:-5]
        total_fore = 0
        nonzero_fore = 0
        for i in range(0, 14):
            if other_elements[i] != 0:
                total_fore += other_elements[i]
                nonzero_fore += 1
        if nonzero_fore != 0:
            average_olddata = total_fore / nonzero_fore

        if credit == 5 and classifications[14] == dominant_class and classifications[13] == dominant_class:
            inferior_class = 1 - inferior_class
            dominant_class = 1 - dominant_class
        elif 2.8 < sum_ratio < 4.8 and (gap_ratio > 2 or gap_ratio < 1/2) and average_olddata != 0:
            conversion_ratio = average_newdata / average_olddata
            print('we have got it! it should be renewed')
            for i in range(0, 14):
                row_data[i] = row_data[i] * conversion_ratio

        first_inferior = None

        # 进行插值
        print(row_data)
        for i in range(0, len(classifications)):
            #if classifications[i] == 2:
            #    classifications[i] = inferior_class
            if classifications[i] == inferior_class:  # 需要插值的点是弱势类别
                j = i + 1
                if first_inferior is None:
                    first_inferior = i
                while j < len(classifications) and classifications[j] == inferior_class:
                    j += 1
                if j < len(classifications) and i - 1 > -1:
                    start_value = row_data[i - 1]
                    print(i)
                    end_value = row_data[j]
                    print(j)
                    count = j - i + 1
                    for k in range(i, j + 1):
                        # 对弱势组进行插值，即分类为 `inferior_class` 的数据点
                        row_data[k] = start_value + (end_value - start_value) * (k - i + 1) / count
                last_inferior = i
                print("process: ", row_data)
        print("right after process: ", row_data)
        print("final class: ", classifications)
        print("first inferior: ", first_inferior)
        print("first dominant: ", first_dominant)
        # 处理第一优势值之前的所有劣势值
        if first_inferior is not None and first_dominant is not None:
            common_difference = row_data[first_dominant + 1] - row_data[first_dominant]
            for i in range(first_inferior, first_dominant):
                row_data[i] = row_data[first_dominant] - (first_dominant - i) * common_difference

        # 处理最后优势值之后的所有劣势值
        if last_inferior is not None and last_dominant is not None:
            common_difference = row_data[last_dominant] - row_data[last_dominant - 1]
            for i in range(last_dominant + 1, len(classifications)):
                row_data[i] = row_data[i - 1] + common_difference

        print(inferior_class)

        for i in range(1, len(row_data)):
            print(row_data[i])
            if row_data[i] != '':
                if row_data[i] < 0:
                    row_data[i] = raw_row_data[i]
                    print("the ", i, "th data is below 0.")
                    print("changed ", row_data[i], " into ", raw_row_data[i])
                if type(row_data[0]) != type(""):
                    if row_data[0] < 0:
                        row_data[0] = raw_row_data[0]
                if classifications[i] == inferior_class:
                    if row_data[i] == 0 or row_data[i - 1] == 0:
                        row_data[i] = row_data[i]  # 跳过包含0的数对
                    else:
                        ratio = row_data[i] / row_data[i - 1]
                        if ratio > 2 or ratio < 1 / 2:
                            row_data[i] = raw_row_data[i]

        print("last inferior: ", last_inferior)
        print(row_data)
        output_data.append(row_data)
        print("output is: ", output_data)

    # 创建新的DataFrame以存储处理后的数据
    result_df = pd.DataFrame(output_data)
    print("result is: ", result_df)
    print(df.shape)
    # 写入到新的Excel文件
    result_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)  # 写入当前工作簿

# 保存Excel文件
writer.close()



