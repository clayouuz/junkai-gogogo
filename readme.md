## junkai gogogo

该项目仅供学习交流使用

### 使用方法：

使用ocr功能需要安装Tesseract ，Poppler 

安装一些python包：

```shell
pip install pytesseract pdf2image pymupdf shutil openai pandas dotenv
```

api可以到https://ai.google.dev/gemini-api/docs领一个免费的,然后修改配置

```python
limit_rpm=15#免费api有限制
api key="上面获取的api"
base_url="https://generativelanguage.googleapis.com/v1beta/openai/"
```

把要处理的pdf放在pdfs文件夹下，运行python脚本

处理结果会保存到一个xlsx文件内，处理失败的pdf会复制到err_pdfs文件夹里

部分数据无法提取,所以请处理完成后检查一下xlsx文件内有无异常值


join.py可以把多次运行的结果合并起来，更加懒人
