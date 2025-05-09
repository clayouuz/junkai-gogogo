import pandas as pd
import os


def find_correct_file(c_name, folder_path):
    """
    根据C列文件名的前6个字符在文件夹中查找匹配的文件名
    """
    prefix = c_name[:6]
    for file in os.listdir(folder_path):
        if file.startswith(prefix) and file.endswith('.pdf'):
            filename = file[:-4]  # 去掉末尾的.pdf
            return filename
    return c_name


# 假设Excel文件名为your_file.xlsx，文件夹路径为your_folder_path
excel_file_path = 'table.xlsx'
folder_path = 'pdfs'

# 读取Excel文件
df = pd.read_excel(excel_file_path)

# 遍历DataFrame的每一行
for index, row in df.iterrows():
    if row['对或错'] == False:  # 检查B列的值是否为False
        correct_file_name = find_correct_file(row['文件名称'], folder_path)
        df.at[index, '文件名称'] = correct_file_name

# 转换日期格式为'yyyy/mm/dd'
df['变动开始日期'] = pd.to_datetime(df['变动开始日期'], errors='coerce').dt.strftime('%Y/%m/%d')
df['变动结束日期'] = pd.to_datetime(df['变动结束日期'], errors='coerce').dt.strftime('%Y/%m/%d')

# 将修改后的数据写回到Excel文件中
df.to_excel(excel_file_path, index=False)