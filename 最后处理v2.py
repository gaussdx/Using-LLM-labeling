import pandas as pd
import os
from collections import Counter

# 定义文件夹路径
folder_path = r''
file_names = [
    
]

# 存储所有数据
data_frames = []

# 读取所有Excel表格
for file_name in file_names:
    df = pd.read_excel(os.path.join(folder_path, file_name))
    data_frames.append(df)

# 统计结果的DataFrame
result_df = pd.DataFrame()

# 保留前两列和第一行
result_df = pd.concat([data_frames[0].iloc[:, :2], 
                       pd.DataFrame(columns=data_frames[0].columns[2:])], axis=1)

empty_count = 0

# 遍历每个需要比较的单元格
for index in range(1, len(data_frames[0])):
    for col in range(2, 5):
        values = [df.iloc[index, col] for df in data_frames]
        value_counts = Counter(values)
        
        # 找到出现次数超过一半的值
        majority_value = None
        for value, count in value_counts.items():
            if count > len(values) // 2:
                majority_value = value
                break
        
        if values.count(values[0]) >= n:  # 判断是否有n个相同
            result_df.iloc[index, col] = str(values[0])
        elif majority_value is not None:
            result_df.iloc[index, col] = '?'+str(majority_value)
        else:
            result_df.iloc[index, col] = ''
            empty_count += 1

# 添加第一行
result_df.iloc[0] = data_frames[0].iloc[0]

# 保存结果
result_df.to_excel(os.path.join(folder_path, 'shuchu2.xlsx'), index=False)

# 输出空值统计
print(f"空了 {empty_count} 个单元格")