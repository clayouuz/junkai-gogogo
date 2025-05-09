import pandas as pd
import os
import re
from openpyxl import load_workbook
from openpyxl.styles import PatternFill


def find_correct_file(stock_code, folder_path):
    """
    根据股票代码在文件夹中查找匹配的文件名
    """
    matching_files = []
    for file in os.listdir(folder_path):
        if file.startswith(stock_code) and file.endswith('.pdf'):
            filename = file[:-4]  # 去掉末尾的.pdf
            matching_files.append(filename)
    return matching_files


# 假设Excel文件名为your_file.xlsx，文件夹路径为your_folder_path
excel_file_path = 'table.xlsx'
folder_path = 'pdfs'

# 读取Excel文件
df = pd.read_excel(excel_file_path)

# 新增一列 '备注'
df['备注'] = ''

# 遍历DataFrame的每一行
for index, row in df.iterrows():
    if row['对或错'] == False:  # 检查'对或错'列的值是否为False
        stock_code = str(row['文件名称'])[:6]  # 获取文件名称的前6位作为股票代码
        matching_files = find_correct_file(stock_code, folder_path)

        if len(matching_files) == 1:
            df.at[index, '文件名称'] = matching_files[0]  # 如果找到一个匹配的文件，则更新'文件名称'列
        elif len(matching_files) > 1:
            df.at[index, '文件名称'] = ', '.join(matching_files)  # 将多个匹配的文件名连接成字符串
            df.at[index, '备注'] = 'fuck'  # 如果找到多个匹配的文件，则在'备注'列写入 'fuck'


# 转换日期格式为'yyyy/mm/dd'
df['变动开始日期'] = pd.to_datetime(df['变动开始日期'], errors='coerce').dt.strftime('%Y/%m/%d')
df['变动结束日期'] = pd.to_datetime(df['变动结束日期'], errors='coerce').dt.strftime('%Y/%m/%d')

# 将修改后的数据写回到Excel文件中
df.to_excel(excel_file_path, index=False)