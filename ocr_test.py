import pytesseract
from pdf2image import convert_from_path


def ocr_pdf(pdf_path, lang='eng'):
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
        for image in images:
            # 对每个图像进行 OCR 处理
            page_text = pytesseract.image_to_string(image, lang=lang)
            text += page_text
        return text
    except Exception as e:
        print(f"发生错误: {e}")
        return None
if __name__ == "__main__":
    # 示例 PDF 文件路径
    pdf_path = ".\\processed_pdfs\\000007_全新好_2022-09-06_深圳市全新好股份有限公司简式权益变动报告书（更新前）.pdf"  # 替换为你的 PDF 文件路径
    extracted_text = ocr_pdf(pdf_path, lang='chi_sim+eng')
    if extracted_text:
        print("OCR 解析结果:")
        print(extracted_text)
    else:
        print("未能成功解析 PDF 文件。")
    