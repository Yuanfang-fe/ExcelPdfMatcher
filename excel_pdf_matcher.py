# excel_pdf_matcher.py
# -*- coding: utf-8 -*-
import pandas as pd
import fitz  # PyMuPDF
import os


def clean_part_number(part):
    """
    清理和标准化 Part No：去除空格，连接破折号。
    """
    return part.replace(" ", "").replace("- ", "-").replace(" -", "-").strip()


def extract_part_numbers_from_excel(excel_path, field_name="Part No"):
    """
    从 Excel 中提取指定字段（第 12 行作为表头）对应的 Part No。
    """
    try:
        df = pd.read_excel(excel_path, header=11)  # header=11 表示第 12 行作为表头
        if field_name not in df.columns:
            raise ValueError(f"Excel 中找不到字段名：{field_name}")
        part_numbers = df[field_name].dropna().astype(str).apply(clean_part_number).tolist()
        return part_numbers
    except Exception as e:
        raise ValueError(f"读取 Excel 文件失败：{e}")


def extract_text_from_pdf(pdf_path):
    """
    提取 PDF 中的文本内容，并清洗破折号格式。
    """
    try:
        text = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text += page.get_text()
        return text.replace("- ", "-").replace(" -", "-")
    except Exception as e:
        raise ValueError(f"读取 PDF 文件失败：{e}")


def match_part_numbers(excel_parts, pdf_text):
    """
    匹配 Excel 中的 Part No 是否出现在 PDF 文本中。
    """
    matched = [p for p in excel_parts if p in pdf_text]
    return matched


def save_matched_results(matched_parts, output_path):
    """
    保存匹配结果到 Excel 文件。
    """
    df = pd.DataFrame({'Matched Part No': matched_parts})
    df.to_excel(output_path, index=False)


def compare_excel_pdf(excel_path, pdf_path, field_name="Part No", output_path=None):
    """
    主函数：读取 Excel 和 PDF，进行 Part No 匹配，并保存结果。
    如果未提供 output_path，将自动生成一个结果文件名。
    """
    if not output_path:
        excel_name = os.path.splitext(os.path.basename(excel_path))[0]
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = f"{excel_name}_VS_{pdf_name}_匹配结果.xlsx"

    print(f"📄 读取 Excel: {excel_path}")
    excel_parts = extract_part_numbers_from_excel(excel_path, field_name)

    print(f"📄 读取 PDF: {pdf_path}")
    pdf_text = extract_text_from_pdf(pdf_path)

    print("🔍 开始匹配 Part No...")
    matched_parts = match_part_numbers(excel_parts, pdf_text)

    save_matched_results(matched_parts, output_path)

    print(f"✅ 共匹配到 {len(matched_parts)} 个 Part No")
    print(f"📁 结果已保存为: {output_path}")
    return output_path
