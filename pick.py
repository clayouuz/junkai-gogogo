import pandas as pd


def copy_rows_to_new_excel(input_file, output_file_1, output_file_2):
    try:
        # 读取 Excel 文件
        df = pd.read_excel(input_file)

        # 筛选出第一列包含指定字段的行
        condition = df.iloc[:, 0].astype(str).str.contains('2022-|2023-|2024-')
        filtered_df = df[condition]

        # 筛选出第一列不包含指定字段的行
        remaining_df = df[~condition]

        # 将筛选后的数据保存到新的 Excel 文件
        filtered_df.to_excel(output_file_1, index=False)
        remaining_df.to_excel(output_file_2, index=False)
        print(f"已成功将符合条件的行复制到 {output_file_1}")
        print(f"已成功将剩余的行复制到 {output_file_2}")
    except FileNotFoundError:
        print(f"错误: 未找到文件 {input_file}")
    except Exception as e:
        print(f"发生未知错误: {e}")


if __name__ == "__main__":
    input_file = 'input.xlsx'  # 替换为你的输入 Excel 文件路径
    output_file_1 = 'output_1.xlsx'  # 替换为你想要保存符合条件行的输出 Excel 文件路径
    output_file_2 = 'output_2.xlsx'  # 替换为你想要保存剩余行的输出 Excel 文件路径
    copy_rows_to_new_excel(input_file, output_file_1, output_file_2)
    