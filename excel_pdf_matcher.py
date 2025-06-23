# excel_pdf_matcher.py
# -*- coding: utf-8 -*-
import pandas as pd
import fitz  # PyMuPDF
import os
import re

def clean_part(value, is_weight_field=False):
    """
    清理字段值，必要时拼接单位（如 KG）。
    - 去除空格、异常破折号
    - 对数值型处理：如 1.0 → 1，再加 KG
    """
    value = str(value).strip()

    # 尝试转换为数字
    if is_weight_field:
        try:
            number = float(value)
            if number.is_integer():
                value = f"{int(number)}KG"
            else:
                value = f"{number}KG"
        except ValueError:
            # 若非数字，保留原样后加KG
            value = f"{value}KG"

    return value.replace(" ", "").replace("- ", "-").replace(" -", "-").strip()


def extract_field_values(df, field):
    """提取某一列的非空清洗值，自动判断是否为重量字段"""
    is_weight_field = "KG" in field.upper()  # 如 NW(KG)、GW(KG)
    if field not in df.columns:
        raise ValueError(f"Excel 中找不到列：{field}")
    return list(set(df[field].dropna().apply(lambda v: clean_part(v, is_weight_field)).tolist()))


def extract_part_rows_from_excel(excel_path, field_names):
    """
    从 Excel 第二张表中提取多个字段的非空值集合
    返回 dict：字段名 -> 清洗后的值列表
    自动判断 .xls 使用 xlrd 引擎
    """
    try:
        xl = pd.ExcelFile(excel_path, engine='xlrd' if excel_path.endswith('.xls') else None)
        df = xl.parse(xl.sheet_names[1], header=12)  # 第二张表，表头第13行

        field_values = {}
        for field in field_names:
            values = extract_field_values(df, field)
            field_values[field] = values
        return field_values
    except Exception as e:
        raise ValueError(f"读取 Excel 时出错：{e}")


def extract_text_from_pdf(pdf_path):
    """提取 PDF 文本"""
    try:
        text = ""
        with fitz.open(pdf_path) as doc:
            for page in doc:
                text += page.get_text()
        return text.replace("- ", "-").replace(" -", "-")
    except Exception as e:
        raise ValueError(f"读取 PDF 文件失败：{e}")


def match_part_values(values, pdf_text):
    """返回匹配上的列表"""
    return [v for v in values if v in pdf_text]


def save_results_to_excel_sheets(matched_dict, output_file):
    """按字段保存为多个 Sheet"""
    with pd.ExcelWriter(output_file) as writer:
        for field, matched_values in matched_dict.items():
            df = pd.DataFrame({f'Matched from {field}': matched_values})
            df.to_excel(writer, sheet_name=field[:31], index=False)
    print(f"📁 匹配结果已保存为: {output_file}")


def compare_excel_pdf(excel_path, pdf_path, field_input="Part No,NW(KG)", output_path=None):
    """
    主流程入口：支持多列字段（逗号分隔），将结果输出为多个 Sheet
    """
    field_names = [f.strip() for f in re.split(r'[，,]', field_input.strip())]

    if not output_path:
        excel_name = os.path.splitext(os.path.basename(excel_path))[0]
        pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
        output_path = f"{excel_name}_VS_{pdf_name}_字段分组匹配结果.xlsx"

    print(f"📄 读取字段列：{field_names}")
    field_values_dict = extract_part_rows_from_excel(excel_path, field_names)

    print(f"📄 读取 PDF 内容: {pdf_path}")
    pdf_text = extract_text_from_pdf(pdf_path)

    matched_dict = {}
    for field, values in field_values_dict.items():
        matched = match_part_values(values, pdf_text)
        matched_dict[field] = matched
        print(f"🔍 {field} 匹配到 {len(matched)} 条")

    save_results_to_excel_sheets(matched_dict, output_path)

    return output_path
