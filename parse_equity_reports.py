import os
import pymupdf as fitz
from openai import OpenAI
import pandas as pd
from tqdm import tqdm
import json
import shutil  # 添加用于复制文件的模块
import datetime
import pytesseract
from pdf2image import convert_from_path

# 设置你的OpenAI API Key
client = OpenAI(
    api_key="",
    base_url=""
    )
model_name = "gemini-2.0-flash"
limit_rpm = 15 # 限制每分钟请求数

def extract_text_from_pdf(pdf_path, max_pages=5):
    text = ""
    with fitz.open(pdf_path) as doc:
        for page in doc:    #doc[:max_pages]的写法放弃
            text += page.get_text()
    return text
def ocr_pdf(pdf_path, lang='chi_sim'):
    """
    使用 OCR 解析 PDF 文件
    :param pdf_path: PDF 文件的路径
    :param lang: 识别语言，默认为英文
    :return: 解析后的文本
    """
    try:
        # 将 PDF 转换为图像
        images = convert_from_path(pdf_path)
        text = ""
        for image in tqdm(images):
            # 对每个图像进行 OCR 处理
            page_text = pytesseract.image_to_string(image, lang=lang)
            text += page_text
        return text
    except Exception as e:
        print(f"发生错误: {e}")
        return None
def construct_prompt(file_name, pdf_text):
    return f"""
你是信息提取专家，请从以下PDF内容中准确地提取结构化信息。以JSON数组形式返回一个或多个数据，如果没有提取到可信的信息则返回空数据。每个JSON对象包含以下字段：
```json
[
    {{
        "文件名称": "",
        "报告类型": "简式"或"详式",
        "变动方向": "增持"或"减持"或"不变",
        "变动方式": "集中竞价"或"连续竞价"或"大宗交易"或"协议转让"或"取得上市公司发行的新股"或"国有股行政划转或变更"或"间接方式转让"或"执行法院裁定"或"继承"或"赠与"或"被动的股权稀释"或"其他",
        "变动开始日期": "YYYY/MM/DD",
        "变动结束日期": "YYYY/MM/DD"
    }}
]
```

提取信息时需遵循以下规则：
1. **文件名称**：直接采用PDF文件名，如“688519南亚新材2022 - 10 - 18简式权益变动报告书” 。
2. **报告类型**：首先查看文件名，若文件名中包含“详式”，则报告类型为“详式”；若文件名未体现，查看目录或节标题中是否有“资金来源”和“后续计划”这两节，若有则为“详式”；若仍无法判断，出现以下三种情况也判定为“详式”：单独或合计持股比例达到20% ；持股比例未达到20%，但信息披露义务人是公司第一大股东或实际控制人；涉及公司控制权变更，或有后续增持计划甚至收购意图。若以上均不满足，则为“简式” 。
3. **变动方向**：从"增持"和"减持"和"不变"中选择文档中所涉及的变动方向，你需要根据文档内容充分判断是增持还是减持，以下两种情况视为不变：第一，股票在同一实际控制人下属不同直接持股股东之间的划转；第二，
一致行动人关系的解除。如果无法判断请留空。
4. **变动方式**：从“集中竞价”“连续竞价”“大宗交易”“协议转让”“取得上市公司发行的新股”“国有股行政划转或变更”“间接方式转让”“执行法院裁定”“继承”“赠与”“被动的股权稀释”“其他”中选取文档提及的变动方式。
5. **变动开始日期和变动结束日期**：确保日期格式为“YYYY/MM/DD”。若文档仅公布日期，则起始日期和结束日期为同一天；若仅公布月份，则起始日期和结束日期为这个月的第一天和最后一天；其他模糊情况参照仅公布月份的处理方式。若文档中未提及某些字段对应信息，相关字段则留空。 


PDF文件名为：{file_name}

PDF正文如下，重点分析权益变动方式一节：
{pdf_text}


"""

def call_openai(prompt):
    response = client.chat.completions.create(
        model=model_name,
        messages=[
            {"role": "user", "content": prompt}
        ],
        temperature=0.2
    )
    return response.choices[0].message.content

def process_all_pdfs(folder_path, output_excel_path=None):
    # 确保文件路径是绝对路径
    folder_path = os.path.abspath(folder_path)
    
    # 如果未指定输出路径，则在脚本所在目录创建一个带时间戳的输出文件
    if output_excel_path is None:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        output_excel_path = os.path.join(
            os.path.dirname(folder_path), 
            f"提取结果_{timestamp}.xlsx"
        )
    else:
        output_excel_path = os.path.abspath(output_excel_path)
    
    # 创建错误文件夹(如果不存在)
    err_folder = os.path.join(os.path.dirname(folder_path), "err_pdfs")
    # 创建已处理文件夹(如果不存在)
    processed_folder = os.path.join(os.path.dirname(folder_path), "processed_pdfs")
    os.makedirs(err_folder, exist_ok=True)
    os.makedirs(processed_folder, exist_ok=True)
    
    records = []
    failed_files = []
    processed_count = 0
    
    # 获取PDF文件列表
    pdf_files = [f for f in os.listdir(folder_path) if f.endswith(".pdf")]
    total_files = len(pdf_files)
    
    # 初始化进度条
    with tqdm(total=total_files) as pbar:
        for filename in pdf_files:
            full_path = os.path.join(folder_path, filename)
            file_name_no_ext = filename.replace(".pdf", "")
            
            try:
                pdf_text = extract_text_from_pdf(full_path)
                
                # 检查提取的文本是否为空或过短
                if not pdf_text or len(pdf_text.strip()) < 10:  # 假设少于10个字符视为无效文本
                    print(f"⚠️ 文件 {filename} 常规提取未获取到有效文本，尝试使用OCR...")
                    # 尝试使用OCR提取文本
                    pdf_text = ocr_pdf(full_path, lang='chi_sim')
                    
                    # 如果OCR也无法提取有效文本
                    if not pdf_text or len(pdf_text.strip()) < 10:
                        print(f"⚠️ 文件 {filename} OCR提取也失败，将复制到错误文件夹")
                        # 复制文件到错误文件夹
                        shutil.copy2(full_path, os.path.join(err_folder, filename))
                        # 删除原文件
                        os.remove(full_path)
                        failed_files.append(filename)
                        pbar.update(1)
                        continue
                    else:
                        print(f"✅ 文件 {filename} 通过OCR成功提取文本")
                    
                prompt = construct_prompt(file_name_no_ext, pdf_text)

                reply = call_openai(prompt)
                # 使用json.loads替代eval以提高安全性
                try:
                    data = json.loads(reply)
                except json.JSONDecodeError:
                    # 尝试清理响应中可能包含的额外文本
                    import re
                    json_str = re.search(r'\[\s*{.*}\s*\]', reply, re.DOTALL)
                    if json_str:
                        data = json.loads(json_str.group(0))
                    else:
                        raise Exception("无法解析JSON响应")
                
                # 检查提取的数据是否为空
                if not data:
                    raise Exception("提取的数据为空")
                    
                # 添加新记录
                for record in data:
                    record["文件名称"] = file_name_no_ext
                    records.append(record)
                
                # 每处理一个文件都更新Excel
                df = pd.DataFrame(records)
                df.to_excel(output_excel_path, index=False)
                
                # 将文件移动到已处理文件夹
                shutil.copy2(full_path, os.path.join(processed_folder, filename))
                # 删除原文件
                os.remove(full_path)
                
                processed_count += 1
                print(f"✅ 已处理文件 {filename} 并更新Excel表格，共处理{processed_count}/{total_files}个文件")
                
            except Exception as e:
                print(f"⚠️ 无法处理文件 {filename}：{e}")
                # 复制文件到错误文件夹
                shutil.copy2(full_path, os.path.join(err_folder, filename))
                # 删除原文件
                os.remove(full_path)
                failed_files.append(filename)
            if limit_rpm > 0:
                import time
                # 控制请求速率，确保每分钟不超过 limit_rpm 次请求
                time.sleep(60 / limit_rpm)
            
            pbar.update(1)
    
    # 输出处理结果统计
    print(f"\n处理完成! 总文件数: {total_files}, 成功处理: {processed_count}, 失败: {len(failed_files)}")
    print(f"✅ 所有数据已保存至：{output_excel_path}")
    print(f"✅ 已处理文件已移动到：{processed_folder}")
    if failed_files:
        print(f"⚠️ 失败文件已复制到: {err_folder}")
        
    return output_excel_path

# 用法示例
if __name__ == "__main__":
    # 使用相对于脚本的路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 默认PDF文件夹位于脚本同级目录下的pdfs文件夹
    default_pdf_folder = os.path.join(script_dir, "pdfs")
    
    # 如果pdfs文件夹不存在，则创建
    if not os.path.exists(default_pdf_folder):
        os.makedirs(default_pdf_folder)
        print(f"已创建PDF文件夹：{default_pdf_folder}")
        print("请将PDF文件放入该文件夹后重新运行脚本")
    else:
        # 使用带时间戳的默认输出路径
        process_all_pdfs(default_pdf_folder)
