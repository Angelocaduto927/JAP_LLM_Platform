'''Question_Generator v3.1 - Revision Papers Comparison'''

# 在开始全新一轮数量统计前，记得删去旧的model_comparison_summary文件夹
import re
import os
import time
import string
import pandas as pd
from openpyxl import Workbook, load_workbook

from docx import Document
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

from jap_excel_processor import parse_questions, parse_excel_to_text
from jap_excel_processor import store_questions_to_excel
from jap_excel_processor import process_word_to_excel



def find_latest_iteration_file(origin_paper: str, input_dir_revised: str):
    """_summary_

    Args:
        origin_paper (str): 要比对的原始试卷
        input_dir_revised (str): 修正后的试卷目录
    """
    pattern = rf"{re.escape(origin_paper)}_iteration_(\d+)\.xlsx$"
    files = [f for f in os.listdir(input_dir_revised) if re.match(pattern, f)]
    if not files:
        return None
    return max(files, key=lambda x: int(re.search(r'iteration_(\d+)\.xlsx$', x).group(1)))

def paper_comparison(original_paper_path: str, revised_paper_path: str):
    """_summary_

    Args:
        original_paper_path (str): 原始试卷路径
        revised_paper_path (str): 修正后试卷路径
        output_dir (str): 输出目录
        model (str): 使用的模型
        is_thinking_model (bool): 是否为思维模型
    """
    try:
        difference = []
        num = 0
        original_list = parse_excel_to_text(original_paper_path)
        revised_list = parse_excel_to_text(revised_paper_path)

        list_record = ["0"]*len(original_list)
        
        if not original_list or not revised_list:
            print(f"Failed to parse questions from one of the files: {original_paper_path} or {revised_paper_path}")
            return [], 0
        if len(original_list) != len(revised_list):
            print(f"Question count mismatch between original and revised papers: {original_paper_path} ({len(original_list)}) vs {revised_paper_path} ({len(revised_list)})")
            return [], 0
        
        for i, (orig, rev) in enumerate(zip(original_list, revised_list)):
            if len(orig) != len(rev):
                print(f"Question {i+1} length mismatch: original ({len(orig)}) vs revised ({len(rev)})")
                return [], 0
            if len(orig) < 7:
                print(f"Question {i+1} in original paper has wrong structure: original ({len(orig)})")
                return [], 0
            if len(rev) < 7:
                print(f"Question {i+1} in revised paper has wrong structure: revised ({len(rev)})")
                return [], 0
            if orig[0] != rev[0]:
                print(f"Question {i+1} number mismatch: original ({orig[0]}) vs revised ({rev[0]})")
                continue
            if orig[1] != rev[1] or orig[2] != rev[2] or orig[3] != rev[3] or orig[4] != rev[4] or orig[5] != rev[5] or orig[6] != rev[6]:
                difference.append(f"{orig[0]}")
                match = re.match(r'Q(\d+): もんだい(\d+)', orig[0])
                if match:
                    question_num = match.group(1)
                    problem_num = match.group(2)
                    list_record[(int(problem_num)-1)*10+int(question_num)-1] = "1"
                else:
                    print(f"Unexpected question format: {orig[0]}")
                num += 1

        string_record = "".join(list_record)
        
        return difference, num, string_record
    except Exception as e:
        print(f"Error during paper comparison: {e}")
        return [], 0, "0"

def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)
        elif os.path.isdir(file_path):
            clear_folder(file_path)
            os.rmdir(file_path)


# ---------------------------
# 批处理函数
# ---------------------------
def batch_process_excel_files(input_dir_origin: str, input_dir_revised: str, output_dir: str, model: str, is_thinking_model: bool):
    """
    批量处理一个目录下的所有Excel文件
    
    Args:
        input_dir_origin: 旧试卷输入目录(如 docs/Generated_paper/shatin/)
        input_dir_revised: 模型修正试卷输入目录
        output_dir: 比较结果输出目录
        model: 使用的模型名称
        is_thinking_model: 是否为thinking模型
        max_iterations: 最大迭代次数
    """


    # 查找所有Excel文件
    excel_files = [f for f in os.listdir(input_dir_origin) if f.endswith('.xlsx') and not f.startswith('~$')]

    print(f"Found {len(excel_files)} Excel files to process")
    print(f"Checking model: {model} ({'Thinking mode' if is_thinking_model else 'Standard mode'})")
    
    for excel_file in excel_files:
        print(f"\n{'='*50}")
        print(f"Processing: {excel_file}")
        print(f"{'='*50}")
        
        saving_path = os.path.join(output_dir, f"{excel_file}")
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        original_paper_path = os.path.join(input_dir_origin, excel_file)
        orig_len = len(parse_excel_to_text(original_paper_path))
        
        latest_file = find_latest_iteration_file(os.path.splitext(excel_file)[0], input_dir_revised)
        if latest_file == None:
            print(f"No revised file found for {excel_file} in {input_dir_revised}, skipping...")
            
            if not os.path.exists(saving_path):
                wb = Workbook()
                ws = wb.active
                ws.append(["Model", "Difference Count", "Differences"])
            else:
                wb = load_workbook(saving_path)
                ws = wb.active
            ws.append([model, 0, "", "0"*orig_len])
            wb.save(saving_path)
            
            continue
        revised_paper_path = os.path.join(input_dir_revised, latest_file)
        
        difference , num, string_record = paper_comparison(original_paper_path, revised_paper_path)
        
        '''
        if os.path.exists(output_dir):
            clear_folder(output_dir)
        '''
        feedback_path = os.path.join("docs/paper_with_feedback/question_index", f"{os.path.splitext(excel_file)[0]} - with HO's Comments.xlsx")
        if not os.path.exists(saving_path):
            wb = Workbook()
            ws = wb.active
            ws.append(["Model", "Difference Number", "Differences", "String", "Correct Modification Number", "Correct Rate"])
        else:
            wb = load_workbook(saving_path)
            ws = wb.active
            
        if not os.path.exists(feedback_path):
            ws.append([model, num, ", ".join(difference), string_record, "NA", "NA"])
        else:
            df = pd.read_excel(feedback_path)
            feedback_string = df.iloc[0, 1]
            bitwise_xor = int(feedback_string, 2) ^ int(string_record, 2)
            bitwise_xor_str = bin(bitwise_xor)[2:].zfill(orig_len)
            ws.append([model, num, ", ".join(difference), string_record, bitwise_xor_str.count("1"), bitwise_xor_str.count("1")/orig_len])
        wb.save(saving_path)
        
            
            
def main():
    """主函数 - Excel比对模式"""
    
    # 定义模型配置
    model_config = {
        "qwen3-max": {"is_thinking": False},
        "qwen3-max-preview": {"is_thinking": False},
        "qwen-plus": {"is_thinking": True}, 
        "qwen3-vl-235b-a22b-instruct": {"is_thinking": False},
        "qwen-flash": {"is_thinking": True},
        "qwen3-30b-a3b-instruct-2507": {"is_thinking": False},
        "qwen-mt-plus": {"is_thinking": False},
        "qwen3-30b-a3b": {"is_thinking": True},
        "qwen3-32b": {"is_thinking": True},
        "qwen3-vl-235b-a22b-thinking": {"is_thinking": True},
        "qwen3-235b-a22b-thinking-2507":{"is_thinking": True},
        "qwen3-next-80b-a3b-thinking":{"is_thinking": True},
        "qwen3-next-80b-a3b-instruct":{"is_thinking": False},
        "qwen3-235b-a22b-instruct-2507":{"is_thinking": False},
        "qwen3-235b-a22b":{"is_thinking": False},
        "qwen3-30b-a3b-thinking-2507":{"is_thinking": True}
    }
    
    available_models = list(model_config.keys())
    
    print("Available models:")
    for i, model in enumerate(available_models, 1):
        model_type = "Thinking" if model_config[model]["is_thinking"] else "Standard"
        print(f"{i}. {model} [{model_type}]")
    
    # 选择模型
    while True:
        try:
            choice = input(f"\nPlease select a model (1-{len(available_models)}): ").strip()
            model_index = int(choice) - 1
            if 0 <= model_index < len(available_models):
                selected_model = available_models[model_index]
                is_thinking_model = model_config[selected_model]["is_thinking"]
                break
            else:
                print("Invalid choice. Please try again.")
        except ValueError:
            print("Please enter a valid number.")
    
    print(f"\nSelected model: {selected_model}")
    print(f"Model type: {'Thinking mode' if is_thinking_model else 'Standard mode'}")
    
    # 设置路径
    input_dir_origin = "docs/Generated_paper/shatin"  # 输入目录input
    input_dir_revised = f"docs/revised_shatin/{selected_model}"  # 按模型名称分类的输出目录
    output_dir = f"docs/revised_shatin/model_comparison_summary"
    
    # 创建输出目录
    if not os.path.exists(input_dir_revised):
        print(f"can't find directory: {input_dir_revised}, run the model for revision first!")
    
    # 批处理所有Excel文件
    batch_process_excel_files(input_dir_origin, input_dir_revised, output_dir, selected_model, is_thinking_model)

    print(f"\nAll Excel files have been processed using model: {selected_model}")
    print(f"Results saved to: {output_dir}")


if __name__ == "__main__":
    main()