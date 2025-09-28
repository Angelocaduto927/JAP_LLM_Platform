import re
import os
import glob
import time
import string
import warnings
import docx
import mysql.connector
import pandas as pd
from docx import Document
from typing import Any
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# 示例：根据生成文本解析出题目、选项和答案
def parse_questions(text):
    """
    解析生成的文本，返回一个包含元组 (question_index, content, options, answer) 的列表。
    
    适用于格式：
    
    Q1 かれが　手伝って　（  　　　　　 ）　宿題 (しゅくだい) が　終わらなっかった。
    1　もらったから
    2　くれなかったから
    3　ほしいから
    4　ほしかったから
    Answer: 2

    Q2 うちの　子どもは　勉強 (べんきょう) しないで　（  　　　　　 ）　ばかりいる。
    1　あそび
    2　あそぶ
    3　あそばない
    4　あそんで
    Answer: 4
    """
    question_pattern = re.compile(r'(Q\d+.*?)\n(?:Answer:|$)\s*(\d)?', re.DOTALL)
    question_matches = question_pattern.findall(text)

    qa_list = []
    for q_text, answer in question_matches:
        lines = q_text.strip().splitlines()
        if not lines:
            continue

        # 解析题号和题目内容
        m = re.match(r'(Q\d+)[\.\s]*(.*)', lines[0].strip())
        if m:
            q_index = m.group(1)
            content = m.group(2).strip()
        else:
            q_index = ""
            content = lines[0].strip()

        # 解析选项
        options = []
        for line in lines[1:]:
            line = line.strip()
            m_opt = re.match(r'^(\d+)[\.\、\s]+(.*)', line)
            if m_opt:
                options.append(f"{m_opt.group(1)}. {m_opt.group(2).strip()}")
            else:
                # 处理题目可能换行的情况
                content += " " + line

        # 存入解析结果
        qa_list.append((q_index, content, "\n".join(options), answer))

    return qa_list

def store_questions_to_excel(qa_list, output, filename):
    """
    将题目信息存储到 Excel 文件中。
    """
    # 确保保存目录存在
    if not os.path.exists(output):
        os.makedirs(output)

    # 清理文件名中的非法字符
    filename = re.sub(r'[\\/*?:"<>|]', "", filename)

    # 创建 Excel 工作簿和工作表
    wb = Workbook()
    ws = wb.active
    ws.title = "Questions"

    # 设置表头
    headers = ["Question Index", "Content", "Options", "Answer", "Suggestions", "Modifications (if any)"]
    ws.append(headers)

    # 将题目信息写入表格
    for qa in qa_list:
        q_index, content, options, answer = qa
        if not options:
            options = "No options"
        if not answer:
            answer = "No answer"
        ws.append([q_index, content, options, answer, "", ""])

    # 设置各列宽度
    column_widths = {
        "A": 15,
        "B": 50,
        "C": 30,
        "D": 20,
        "E": 20,
        "F": 40
    }
    for col_letter, width in column_widths.items():
        ws.column_dimensions[col_letter].width = width

    # 设置所有单元格自动换行
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)

    # 创建下拉选择的数据验证
    dv = DataValidation(type="list", formula1='"High_Q,Low_Q,Drop,Minor changes"', allow_blank=True)
    ws.add_data_validation(dv)
    dv_range = f"E2:E{ws.max_row}"
    dv.add(dv_range)

    # 保存 Excel 文件
    output_path = os.path.join(output, filename)
    print(f"Saving Excel file to: {output_path}")
    wb.save(output_path)
    print(f"Successfully stored in {output_path}")



def process_word_to_excel(doc_filepath, excel_output):
    """
    Process the content from each Word document in a folder and store them in separate Excel files.
    """
    # 检查文件夹是否存在
    if not os.path.exists(doc_filepath):
        print(f"Error: Directory not found - {doc_filepath}")
        return

    # 获取文件夹中所有的 .docx 文件
    doc_files = glob.glob(os.path.join(doc_filepath, "*.docx"))
    if not doc_files:
        print("No .docx files found in the directory.")
        return

    for filepath in doc_files:
        try:
            filename = os.path.splitext(os.path.basename(filepath))[0]
            doc = Document(filepath)
            all_text = "\n".join([para.text for para in doc.paragraphs])

            # 解析提取的文本
            qa_list = parse_questions(all_text)

            # 将解析的题目和答案保存到 Excel 文件，文件名与 Word 文件名一致
            excel_filename = f"{filename}.xlsx"
            store_questions_to_excel(qa_list, excel_output, excel_filename)

        except Exception as e:
            print(f"Error processing {filepath}: {e}")

def parse_excel_to_text(excel_path: str) -> list:
    """
    从Excel文件中读取题目数据并返回列表格式
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        list: 包含题目数据的列表，每个元素为 [题目, 选项1, 选项2, 选项3, 选项4, 答案]
    """
    try:
        # 读取Excel文件
        df = pd.read_excel(excel_path)
        
        qa_list = []
        
        for index, row in df.iterrows():
            # 跳过标题行
            if pd.isna(row.iloc[1]):
                continue
            # 获取题目内容
            question_index = str(row.iloc[0]) if not pd.isna(row.iloc[0]) else ""
            
            question_content = str(row.iloc[1]) if not pd.isna(row.iloc[1]) else ""
            match = re.search(r': もんだい\d+', question_content)
            if match:
                mondai_str = match.group(0)  # 匹配到的内容
                question_content = question_content.replace(mondai_str, '', 1).strip()  # 删除并去除首尾空格
            else:
                mondai_str = ''
                question_content = question_content.strip()
            
            question_index = question_index.strip()+mondai_str
            
            # 解析选项 (假设选项在第3列，格式为 "1. 选项1\n2. 选项2\n3. 选项3\n4. 选项4")
            options_text = str(row.iloc[2]) if not pd.isna(row.iloc[2]) else ""
            
            # 提取四个选项
            options = ["", "", "", ""]
            if options_text:
                # 使用正则表达式提取选项
                option_matches = re.findall(r'(\d+)\.?\s*([^\n\d]+)', options_text)
                for opt_num, opt_text in option_matches:
                    opt_index = int(opt_num) - 1
                    if 0 <= opt_index < 4:
                        options[opt_index] = opt_text.strip()
            
            # 获取答案
            answer = str(row.iloc[3]) if not pd.isna(row.iloc[3]) else ""
            
            # 只有当题目内容不为空时才添加
            if question_content:
                qa_list.append([
                    question_index,
                    question_content,
                    options[0], 
                    options[1], 
                    options[2], 
                    options[3], 
                    answer
                ])
        
        return qa_list
        
    except Exception as e:
        print(f"Error reading Excel file {excel_path}: {e}")
        return []