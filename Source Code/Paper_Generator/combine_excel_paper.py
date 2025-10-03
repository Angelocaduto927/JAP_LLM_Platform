import re
import os
import shutil
from openpyxl import Workbook, load_workbook
from jap_excel_processor import parse_excel_to_text

def combine_excel_files(selected_model, is_thinking_model, dir, split_number):
    excel = os.listdir(dir)
    record_dict = {}
    name_dict = {}
    for excel_file in excel:
        if excel_file.endswith("_revised_revised.xlsx"):
            pattern = re.match(r"^(.*)_(\d+)_(\d+)_(.*)_revised_revised\.xlsx$", excel_file)
            if pattern:
                prefix = pattern.group(1)
                suffix = pattern.group(4)
                group_key = int(pattern.group(2))
                sort_key = int(pattern.group(3))
                if group_key not in record_dict:
                    record_dict[group_key] = []
                record_dict[group_key].append((sort_key, excel_file))
                if group_key not in name_dict:
                    name_dict[group_key] = []
                    name_dict[group_key].append((prefix, suffix))
    for group_key in record_dict:
        record_dict[group_key].sort(key = lambda x: x[0])
        full_qa = []
        for sort_key, excel_file in record_dict[group_key]:
            new_dir = os.path.join(dir, excel_file)
            
            if full_qa == []:
                workbook = load_workbook(new_dir)
                worksheet = workbook.active
                header = [[cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]]
                full_qa.extend(header)
            qa = parse_excel_to_text(new_dir)
            for question in qa:
                match = re.match(r'(Q\d+)', question[0])
                if match:
                    rest = question[0].replace(match.group(1), '').strip()
                    question[1] = rest + " " + question[1]
                    question[0] = match.group(1)
                    question[2] = "1. "+question[2]+" 2. "+question[3]+" 3. "+question[4]+" 4. "+question[5]
                    question[3] = question[6]
            full_qa.extend(qa)
        tempt_path = os.path.join(dir, "tempt")
        if not os.path.exists(tempt_path):
            os.mkdir(tempt_path)
        workbook = Workbook()
        sheet = workbook.active
        for row in full_qa:
            sheet.append(row)
        prefix, suffix = name_dict[group_key][0]
        output_path = os.path.join(tempt_path,f"{prefix}_{group_key}_{suffix}_revised_revised.xlsx")
        workbook.save(output_path)

    files = os.listdir(dir)

    for excel_file in files:
        file_path = os.path.join(dir, excel_file)
        if os.path.isfile(file_path) and excel_file.endswith(".xlsx"):
            os.remove(file_path)
    
    for excel_file in os.listdir(tempt_path):
        src = os.path.join(tempt_path, excel_file)
        dst = os.path.join(dir, excel_file)
        shutil.move(src, dst)
    
    os.rmdir(tempt_path)

def main(split_number=10, model_index_input = None):
    print("====================================")
    print("如果是单独运行此脚本，试卷默认是按照10道题一份进行拆分的，如果需要更改分题数量，请用pipeline方式运行split_excel_paper.py，再运行此脚本// 或者可以在此脚本中修改split_number的默认值")
    print("====================================")
    
    #定义模型配置
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
    
    if model_index_input is not None:
        selected_model = model_index_input
        is_thinking_model = model_config[selected_model]["is_thinking"]
    else:
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
    dir = f"docs/revised_shatin/{selected_model}"  # 按模型名称分类的输出目录
    
    # 创建输出目录
    if not os.path.exists(dir):
        print(f"can't find directory: {dir}, run the model for revision first!")
    
    # 批处理所有Excel文件
    combine_excel_files(selected_model, is_thinking_model, dir, split_number)
    print(f"\nAll Excel files have been processed using model: {selected_model}")

if __name__ == "__main__":
    main()