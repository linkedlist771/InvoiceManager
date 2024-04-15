from enum import Enum
import os
import fitz  # PyMuPDF
import re


class FileType(Enum):
    PDF = "pdf"
    PNG = "png"
    JPG = "jpg"
    JPEG = "jpeg"


def get_all_invoice_file(dir_path: str, file_type: FileType) -> dict[str, str]:
    invoice_files = {}
    for root, dirs, files in os.walk(dir_path):
        for file in files:
            if file.endswith(file_type.value):
                file_name = file.split(".")[0]
                absolute_path = os.path.join(root, file)
                invoice_files[file_name] = absolute_path
    return invoice_files


def extract_final_price(pdf_path):
    # 打开PDF文件
    doc = fitz.open(pdf_path)

    # 初始化一个空字符串来存储文本内容
    text = ''

    # 读取PDF的每一页
    for page in doc:
        text += page.get_text()

    # 使用正则表达式匹配价格
    # 正则表达式说明：匹配 ¥ 后面跟着数字和小数点的模式
    price_pattern = r'¥(\d+\.\d{2})'
    matches = re.findall(price_pattern, text)

    # 关闭文档
    doc.close()

    if matches:
        # 将匹配到的最后一个价格转换为浮点数并返回
        return float(matches[-1])
    else:
        return -1