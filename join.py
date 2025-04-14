import pandas as pd
import os
import re
import glob

def integrate_excels(file_paths):
    """
    整合多个 Excel 文件的内容并去除重复项
    :param file_paths: 包含多个 Excel 文件路径的列表
    :return: 整合并去重后的 DataFrame
    """
    all_dfs = []
    for file_path in file_paths:
        try:
            df = pd.read_excel(file_path)
            all_dfs.append(df)
            print(f"已读取: {file_path} (行数: {len(df)})")
        except Exception as e:
            print(f"读取文件 {file_path} 时出现错误: {e}")

    if all_dfs:
        combined_df = pd.concat(all_dfs, ignore_index=True)
        # 记录合并前的行数
        total_rows_before = combined_df.shape[0]
        unique_df = combined_df.drop_duplicates()
        # 记录去重后的行数
        total_rows_after = unique_df.shape[0]
        print(f"合并前总行数: {total_rows_before}, 去重后行数: {total_rows_after}, 删除重复项: {total_rows_before - total_rows_after}")
        return unique_df
    return None

# 获取当前脚本所在目录,并在该目录定位"outputs"文件夹
script_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "outputs")

# 使用正则表达式匹配目录下所有"提取结果_*.xlsx"文件
pattern = os.path.join(script_dir, "提取结果_*.xlsx")
file_paths = glob.glob(pattern)

if file_paths:
    print(f"找到 {len(file_paths)} 个匹配的Excel文件:")
    for file in file_paths:
        print(f"- {os.path.basename(file)}")
    
    # 生成带时间戳的输出文件名
    import datetime
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_file = os.path.join(script_dir, f"合并结果_{timestamp}.xlsx")
    
    result = integrate_excels(file_paths)
    if result is not None:
        result.to_excel(output_file, index=False)
        print(f"文件整合并去重完成，结果已保存为 {output_file}")
    else:
        print("没有成功读取到任何文件内容。")
else:
    print("在目录中没有找到匹配的Excel文件。")