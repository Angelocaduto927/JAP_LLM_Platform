'''Question_Generator v3.1 - Excel Revision Mode'''
'''
只修改了文件读入路径，这个文件针对的是question_num_significance_test/paper文档内的以切分好的试卷
'''

import re
import os
import time
import string

from docx import Document
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

from jap_excel_processor import parse_questions, parse_excel_to_text
from jap_excel_processor import store_questions_to_excel
from jap_excel_processor import process_word_to_excel
from jap_voting_machine_v3_4 import get_voting_result, get_experiment_config, create_llm, get_model_revision_for_question

# ---------------------------
# Excel处理相关函数
# ---------------------------
def load_excel_as_text(excel_path: str) -> str:
    """
    从Excel文件中读取题目并转换为文本格式
    
    Args:
        excel_path: Excel文件路径
        
    Returns:
        str: 格式化的题目文本
    """
    try:
        # 使用现有的parse_excel_to_text函数或实现新的解析逻辑
        qa_list = parse_excel_to_text(excel_path)
        
        # 转换为标准格式的文本
        formatted_text = ""
        for i, qa in enumerate(qa_list, 0):
            if len(qa) >= 7:  # 确保有足够的字段
                question_number = qa[0] #题号
                question_stem = qa[1]  # 题干
                option1 = qa[2]        # 选项1
                option2 = qa[3]        # 选项2  
                option3 = qa[4]        # 选项3
                option4 = qa[5]        # 选项4
                answer = qa[6]         # 答案

                formatted_text += f"{question_number}"
                formatted_text += f"{question_stem}\n"
                formatted_text += f"1. {option1} 2. {option2} 3. {option3} 4. {option4}\n"
                formatted_text += f"Answer: {answer}\n\n"
        
        return formatted_text.strip()
    except Exception as e:
        print(f"Error loading Excel file {excel_path}: {e}")
        return ""

def save_text_to_excel(text: str, output_path: str, filename: str):
    """
    将修订后的文本保存为Excel格式
    
    Args:
        text: 修订后的题目文本
        output_path: 输出目录
        filename: 文件名
    """
    try:
        qa_list = parse_questions(text)
        if qa_list:
            store_questions_to_excel(qa_list, output_path, filename)
            print(f"Successfully saved revised questions to {os.path.join(output_path, filename)}")
        else:
            print("No questions parsed from revised text")
    except Exception as e:
        print(f"Error saving to Excel: {e}")

# ---------------------------
# 错误检测函数（与v3相同）
# ---------------------------
def api_call_for_paper_revise(chain, input_data, max_retries=3, delay=1):
    retries = 0
    while retries < max_retries:
        try:
            return chain.invoke({"input_data": input_data["input_data"], "errors": input_data["errors"]})
        except Exception as e:
            print(f"API call failed: {e}. Retrying in {delay} seconds...")
            retries += 1
            time.sleep(delay)
    print("Max retries reached. Giving up.")
    return None

def api_call_for_error_detection(chain, input_data, max_retries=3, delay=1):
    retries = 0
    while retries < max_retries:
        try:
            return chain.invoke({"input_data": input_data["input_data"]})
        except Exception as e:
            print(f"API call failed: {e}. Retrying in {delay} seconds...")
            retries += 1
            time.sleep(delay)
    print("Max retries reached. Giving up.")
    return None

def paper_split_into_list(text):
    """将整个试卷拆分为单独的题目列表"""
    pattern = r'(Q\d+:\s*もんだい\d+[\s\S]*?Answer:\s*\d+)'
    questions = re.findall(pattern, text)
    return questions

def extract_question_by_id(text, question_id):
    """根据题号从整个试卷中提取特定题目"""
    # 转义特殊字符，确保正则表达式匹配准确
    escaped_id = re.escape(question_id)
    pattern = f'({escaped_id}[\\s\\S]*?Answer:\\s*\\d+)'
    match = re.search(pattern, text)
    if match:
        return match.group(1)
    return None

def replace_question(full_text, old_question, new_question):
    """在整个试卷中替换特定题目"""
    # 使用正则表达式避免意外替换
    if old_question in full_text:
        return full_text.replace(old_question, new_question)
    return full_text

def normalize_spaces(text):
    return text.replace('\u3000', ' ')

def normalize_text(text):
    text = text.lower()
    text = text.translate(str.maketrans('', '', string.punctuation))
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def has_multiple_correct_answers(text, llm):
    prompt = ChatPromptTemplate.from_messages(
        [
            ("system", "You are an experienced Japanese N4/N5 examiner reviewing the following multiple-choice questions:\n\n"),
            ("human", 
            """{input_data}\n\n
            Check if any question has **more than one correct answer**. This means that multiple options are valid for the question given its context.\n
            If at least one question has multiple valid correct answers, respond with the question numbers that potentially have the problem only, the form requirement is number + question type (eg. 'Q8: もんだい1, Q7: もんだい2' **questions separated by commas**).\n
            If not, you must return "False" only.\n
            
            Make sure the number of questions remains exactly the same as the input and the question index (the Qx: もんだいy part) for each question is unchanged.
            """
            ),
        ]
    )
    input_data = {'input_data': text}
    chain = prompt | llm
    result = api_call_for_error_detection(chain, input_data)
    if result and result.content != "False" and result.content.strip():
        return result.content
    return False

def has_stem_errors(text, llm):
    prompt = ChatPromptTemplate.from_messages(
        [
            ("system", "You are an experienced Japanese N4/N5 examiner reviewing the following multiple-choice questions:\n\n"),
            ("human", "{input_data}\n\n"
            "Check if any **question stem** (the main question part before the options) has errors, such as:\n"
            "- Grammatical mistakes\n"
            "- Unnatural sentence structures\n"
            "- Ambiguous wording\n"
            "If there is at least one issue in the stems, you must respond with the question numbers that potentially have the problem only, the form requirement is number + question type (eg. 'Q8: もんだい1, Q7: もんだい2').\n"
            "Otherwise, respond with 'False' only.\n"
            "Make sure the number of questions remains exactly the same as the input and the question index (the Qx: もんだいy part) for each question is unchanged."
            )
        ]
    )
    chain = prompt | llm
    input_data = {'input_data': text}
    result = api_call_for_error_detection(chain, input_data)
    if result and result.content != "False" and result.content.strip():
        return result.content
    return False

def has_duplicate_options(text):
    questions_with_options = re.findall(
        r'(\d+)\.\s*(.*?)\n(1\.\s*(.*?)\n)(2\.\s*(.*?)\n)(3\.\s*(.*?)\n)(4\.\s*(.*?)\n)',
        text, re.DOTALL
    )
    for question, _, opt1, _, opt2, _, opt3, _, opt4, _ in questions_with_options:
        options = {opt1.strip(), opt2.strip(), opt3.strip(), opt4.strip()}
        if len(options) < 4:
            print(f"Duplicate options detected in question {question}")
            return True
    return False

def has_duplicate_questions(text):
    questions = re.findall(
        r'(\d+)\.\s*(.*?)\n(1\.\s*(.*?)\n)(2\.\s*(.*?)\n)(3\.\s*(.*?)\n)(4\.\s*(.*?)\n)',
        text, re.DOTALL
    )
    seen_questions = set()
    for question, _, opt1, _, opt2, _, opt3, _, opt4, _ in questions:
        question_text = normalize_text(question.strip())
        options = {normalize_text(opt1.strip()), normalize_text(opt2.strip()),
                   normalize_text(opt3.strip()), normalize_text(opt4.strip())}
        normalized_question = f"{question_text} - {', '.join(sorted(options))}"
        if normalized_question in seen_questions:
            print(f"Duplicate question detected: {question_text}")
            return True
        seen_questions.add(normalized_question)
    return False

def check_for_error(revised_text, llm):
    errors = []
    try:
        Multiple_correct_answers = has_multiple_correct_answers(revised_text, llm)
        if Multiple_correct_answers != False:
            errors.append(["Multiple correct answers", Multiple_correct_answers])
        if has_duplicate_questions(revised_text):
            errors.append("Duplicate questions")
        Stem_question = has_stem_errors(revised_text, llm)
        if Stem_question != False:
            errors.append(["Stem errors", Stem_question])
        if has_duplicate_options(revised_text):
            errors.append("Duplicate options")
        return errors
    except Exception as e:
        print(f"Error in check_for_error: {e}")
        return []

def extract_wrong_questions(question_list_input, error):
    errors = error
    question_list = question_list_input
    multicorrect_problems_number = []
    stem_problems_number = []
    
    if [q for q in errors if isinstance(q, list) and q[0] == "Multiple correct answers"]:
        multicorrect_problems_number = [q for q in errors if isinstance(q, list) and q[0] == "Multiple correct answers"][0][1].split(',')
        multicorrect_problems_number = [q.strip() for q in multicorrect_problems_number]
    if [q for q in errors if isinstance(q, list) and q[0] == "Stem errors"]:
        stem_problems_number = [q for q in errors if isinstance(q, list) and q[0] == "Stem errors"][0][1].split(',')
        stem_problems_number = [q.strip() for q in stem_problems_number]
    
    questions_have_both_issues_number = set(multicorrect_problems_number) & set(stem_problems_number)
    question_have_multicorrect_only_number = set(multicorrect_problems_number) - questions_have_both_issues_number
    question_have_stemerror_only_number = set(stem_problems_number) - questions_have_both_issues_number

    wrong_questions_list = [['Multiple correct answers problems'], ['Stem error problems'], ['Multiple correct answers and Stem error problems']]
    
    for question in question_list:
        q_match = re.match(r'Q\d+:\s*もんだい\d', question)
        if not q_match:
            continue
        if q_match.group(0) in questions_have_both_issues_number:
            wrong_questions_list[2].append(question)
        elif q_match.group(0) in question_have_multicorrect_only_number:
            wrong_questions_list[0].append(question)
        elif q_match.group(0) in question_have_stemerror_only_number:
            wrong_questions_list[1].append(question)
    
    return wrong_questions_list

def combine_revised_questions_into_paper(revised_questions_input, original_text_input):
    revised_question = revised_questions_input
    original_text = original_text_input
    revised_questions = paper_split_into_list(normalize_spaces(revised_question))
    
    for revised in revised_questions:
        q_match = re.match(r'Q\d+:\s*もんだい\d', revised)
        if not q_match:
            continue
        for i, original in enumerate(original_text):
            if original.startswith(q_match.group(0)):
                original_text[i] = revised
                break
    return original_text

def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)
        elif os.path.isdir(file_path):
            clear_folder(file_path)
            os.rmdir(file_path)

# ---------------------------
# 主要修订函数（Excel版本）
# ---------------------------
def excel_revise_simple(excel_path: str, output_dir: str, model: str, is_thinking_model: bool, max_iterations=5, temperature=0.6):
    """
    对Excel文件中的题目进行多轮修订和检查
    
    Args:
        excel_path: 输入的Excel文件路径
        output_dir: 输出目录
        model: 使用的模型
        is_thinking_model: 是否为thinking模型
        max_iterations: 最大迭代次数
        temperature: 模型温度参数
    """
    # 从Excel文件加载题目
    print(f"Loading questions from {excel_path}")
    revised_result = load_excel_as_text(excel_path)
    
    if not revised_result:
        print("Failed to load questions from Excel file")
        return
    
    # 获取文件名（不含扩展名）
    filename = os.path.splitext(os.path.basename(excel_path))[0]
    
    # 根据模型类型初始化LLM
    if is_thinking_model:
        llm_revise = ChatOpenAI(
            temperature=temperature, 
            model=model, 
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", 
            max_tokens=2048, 
            max_retries=5,
            streaming=True,
            extra_body={"thinking_budget": 1000}
        )
    else:
        llm_revise = ChatOpenAI(
            temperature=temperature, 
            model=model, 
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", 
            max_tokens=2048, 
            max_retries=5,
            extra_body={"enable_thinking": False}
        )
    
    prompt_revise = ChatPromptTemplate.from_messages([
        ("human", '''
        You are an experienced Japanese N4/N5 examiner. There are some Japanese multiple-choice questions at the end of this message. All of the questions have some issues. You can refer to {errors}. 
        Your task is to modify those multiple-choice test questions to meet the following criteria and fix all the issues:

        1. No duplicate questions: Ensure that all questions are unique.
        2. No duplicate options: All four options within a question should be unique.  
        3. No multiple reasonable answers: Ensure that only one answer is correct and reasonable.
        4. Grammatical correctness: The title and stem of each question must be grammatically correct.
        5. Relevance of options: Ensure that the stem clearly indicates what cannot be chosen.
        6. Pronunciation and Word Usage: Ensure proper Japanese formatting.
        7. General guidance: Eliminate any ambiguity and ensure appropriate difficulty level.

        8. Output Format: Each question must **keep the original format**:
        - Each question must start with `Qx: もんだいy\n`
        - Each question must have exactly 4 options (`1` to `4`).
        - Each question must have an `Answer: x` at the end.
        - Each question must contain empty parentheses ( ) for the blank.
        - Do not include any other comments.
        
        9. make sure the number of questions remains exactly the same as the input and the question index (the Qx: もんだいy part) for each question is unchanged.
        
        Here are the questions to review and modify:
        {input_data}
        ''')
    ])
    
    # 错误检查LLM也需要同样处理
    if is_thinking_model:
        llm_error_check = ChatOpenAI(
            temperature=0.3, 
            model=model, 
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", 
            max_tokens=2048, 
            max_retries=5,
            streaming=True,
            extra_body={"thinking_budget": 2000}
        )
    else:
        llm_error_check = ChatOpenAI(
            temperature=0.3, 
            model=model, 
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", 
            max_tokens=2048, 
            max_retries=5,
            extra_body={"enable_thinking": False}
        )
    
    chain = prompt_revise | llm_revise
    params = {'input_data': "", 'errors': []}
    question_list = paper_split_into_list(normalize_spaces(revised_result))
    
    # 创建日志文件
    log_path = os.path.join(output_dir, f"{filename}_revision_log.txt")
    
    # 迭代修订
    for iteration in range(max_iterations):
        print(f"\n--- Iteration {iteration + 1} ---")
        errors = check_for_error(revised_result, llm_error_check)
        
        if not errors:
            print(f"No issues found after {iteration + 1} iterations.")
            break
            
        print(f"Detected errors: {errors}")
        
        # 记录错误到日志
        with open(log_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"Iteration {iteration + 1} Errors: {errors}\n")
        
        # 提取错误题目
        wrong_questions = extract_wrong_questions(question_list, errors)
        print(f"Questions to be revised: {len([q for sublist in wrong_questions for q in sublist[1:]])}")
        
        formatted_wrong_questions = '\n'.join(['\n'.join(sublist) for sublist in wrong_questions])
        params['input_data'] = formatted_wrong_questions
        params['errors'] = errors
        
        # 调用修订
        revised_result_obj = api_call_for_paper_revise(chain, params)
        if revised_result_obj is None:
            print("Failed to get revised result. Stopping iteration.")
            break
            
        revised_result = revised_result_obj.content
        
        # 合并修订结果
        revised_result = combine_revised_questions_into_paper(revised_result, question_list)
        revised_result = '\n'.join(revised_result)
        question_list = paper_split_into_list(normalize_spaces(revised_result))
        
        # 保存中间结果
        intermediate_excel = f"{filename}_iteration_{iteration + 1}.xlsx"
        save_text_to_excel(revised_result, output_dir, intermediate_excel)
        
    else:
        print(f"Maximum iterations ({max_iterations}) reached. Errors may still exist.")
    
    # 保存最终结果
    final_excel = f"{filename}_revised.xlsx"
    save_text_to_excel(revised_result, output_dir, final_excel)
    print(f"Final revised questions saved to {os.path.join(output_dir, final_excel)}")

# ---------------------------
# 使用voting machine修订函数
# ---------------------------
def excel_revise_with_voting(excel_path: str, output_dir: str, experiment_group: int, max_iterations=5, temperature=0.6):
    """
    使用voting machine进行Excel文件题目修订（输出与3_3一致：revision_log.txt + iteration_xxx.xlsx + final *_revised.xlsx）
    """
    # 从Excel文件加载题目
    print(f"Loading questions from {excel_path}")
    revised_result = load_excel_as_text(excel_path)
    if not revised_result:
        print("Failed to load questions from Excel file")
        return

    # 文件名与实验组配置
    filename = os.path.splitext(os.path.basename(excel_path))[0]
    config = get_experiment_config(experiment_group)

    # 选择权重最高的修订模型，并创建 LLM
    try:
        revision_model_info = max(config["models"], key=lambda x: x.get("weight", 0))
        revision_model_name = revision_model_info["model_name"]
        is_thinking_model = revision_model_info.get("is_thinking", False)
    except Exception:
        # 兜底：若配置不含权重信息，取第一个模型
        revision_model_info = config["models"][0]
        revision_model_name = revision_model_info["model_name"]
        is_thinking_model = revision_model_info.get("is_thinking", False)

    llm_revise = create_llm(revision_model_name, is_thinking_model, temperature)

    # 仅保留 3_3 风格文本日志
    os.makedirs(output_dir, exist_ok=True)
    log_txt_path = os.path.join(output_dir, f"{filename}_revision_log.txt")
    with open(log_txt_path, 'w', encoding='utf-8') as f:
        f.write(f"Voting Group: {experiment_group}\n")
        f.write(f"Revision model: {revision_model_name} (thinking={is_thinking_model})\n")
        f.write(f"Start: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

    # 迭代修订
    for iteration in range(max_iterations):
        print(f"\n--- Iteration {iteration + 1} ---")

        # 1) 投票检测
        question_votes = get_voting_result(revised_result, experiment_group)
        questions_to_revise = {q: info for q, info in question_votes.items() if info.get("exceeds_threshold", False)}

        # 写入文本日志
        with open(log_txt_path, 'a', encoding='utf-8') as f:
            f.write(f"\nIteration {iteration + 1}\n")
            f.write(f"Threshold: {config.get('threshold')}\n")
            f.write(f"Total voted: {len(question_votes)}\n")
            f.write("To revise: " + (", ".join(sorted(questions_to_revise.keys())) if questions_to_revise else "none") + "\n")

        if not questions_to_revise:
            print(f"No questions need revision after {iteration + 1} iterations.")
            break

        # 2) 按需修订
        revised_count = 0
        for question_id, vote_info in questions_to_revise.items():
            print(f"  - Revising {question_id} (errors: {', '.join(vote_info.get('error_types', []))})")
            original_question = extract_question_by_id(revised_result, question_id)
            if not original_question:
                print(f"    Warning: Could not find question {question_id} in text!")
                continue

            revised_question = original_question
            for error_type in vote_info.get("error_types", []):
                candidate = get_model_revision_for_question(
                    original_question,
                    llm_revise,
                    error_type,
                    is_thinking_model
                )
                if candidate:
                    revised_question = candidate

            if revised_question and revised_question != original_question:
                revised_result = replace_question(revised_result, original_question, revised_question)
                revised_count += 1
                print("    Successfully revised")
            else:
                print("    No changes made")

        print(f"Revised {revised_count} questions in iteration {iteration + 1}")

        # 3) 保存中间产物（与 3_3 一致）
        intermediate_excel = f"{filename}_iteration_{iteration + 1}.xlsx"
        save_text_to_excel(revised_result, output_dir, intermediate_excel)

        if revised_count == 0:
            print("No questions were modified in this iteration. Stopping.")
            break
    else:
        print(f"Maximum iterations ({max_iterations}) reached.")

    # 最终产物命名与 3_3 对齐：filename_revised.xlsx（split 文件名已含 _revised => 产出 *_revised_revised.xlsx）
    final_excel = f"{filename}_revised.xlsx"
    save_text_to_excel(revised_result, output_dir, final_excel)

    with open(log_txt_path, 'a', encoding='utf-8') as f:
        f.write(f"\nFinal saved: {final_excel}\nEnd: {time.strftime('%Y-%m-%d %H:%M:%S')}\n")

    print(f"Final revised questions saved to {os.path.join(output_dir, final_excel)}")

# ---------------------------
# 批处理函数
# ---------------------------
def batch_process_excel_files(input_dir: str, output_dir: str, model: str, is_thinking_model: bool, max_iterations=5):
    """
    批量处理一个目录下的所有Excel文件
    
    Args:
        input_dir: 输入目录（如 docs/Generated_paper/shatin/）
        output_dir: 输出目录
        model: 使用的模型名称
        is_thinking_model: 是否为thinking模型
        max_iterations: 最大迭代次数
    """
    #if not os.path.exists(output_dir):
    #    os.makedirs(output_dir)
        
    if os.path.exists(output_dir):
        clear_folder(output_dir)
    
    # 查找所有Excel文件
    excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    print(f"Found {len(excel_files)} Excel files to process")
    print(f"Using model: {model} ({'Thinking mode' if is_thinking_model else 'Standard mode'})")
    
    for excel_file in excel_files:
        print(f"\n{'='*50}")
        print(f"Processing: {excel_file}")
        print(f"{'='*50}")
        
        excel_path = os.path.join(input_dir, excel_file)
        excel_revise_simple(excel_path, output_dir, model, is_thinking_model, max_iterations)

def batch_process_with_voting(input_dir: str, output_base_dir: str, experiment_group: int, max_iterations=5):
    """
    使用指定实验组的voting machine批量处理Excel文件
    """
    config = get_experiment_config(experiment_group)
    output_dir = os.path.join(output_base_dir, f"voting_group_{experiment_group}")
    
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    else:
        clear_folder(output_dir)
    
    # 查找所有Excel文件
    excel_files = [f for f in os.listdir(input_dir) if f.endswith('.xlsx') and not f.startswith('~$')]
    
    print(f"Found {len(excel_files)} Excel files to process")
    print(f"Using {config['name']} (Group {experiment_group})")
    
    for excel_file in excel_files:
        print(f"\n{'='*50}")
        print(f"Processing: {excel_file}")
        print(f"{'='*50}")
        
        excel_path = os.path.join(input_dir, excel_file)
        excel_revise_with_voting(excel_path, output_dir, experiment_group, max_iterations)

# 修改主函数以支持voting machine选项
def main():
    """主函数 - 支持单模型和voting machine模式"""
    
    print("Select revision mode:")
    print("1. Single model revision (original v3.3 logic)")
    print("2. Voting machine revision (4 experiment groups)")
    print("3. Run all voting machine groups for comparison")
    
    mode_choice = input("Enter your choice (1-3): ").strip()
    
    if mode_choice == "1":
        # 原有的单模型逻辑
        model_config = {
            "qwen-flash": {"is_thinking": True},
            "qwen3-30b-a3b-instruct-2507": {"is_thinking": False},
            "qwen3-30b-a3b": {"is_thinking": True},
            "qwen3-vl-235b-a22b-thinking": {"is_thinking": True},
            "qwen3-235b-a22b-thinking-2507":{"is_thinking": True},
            "qwen3-next-80b-a3b-instruct":{"is_thinking": False},
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
        input_dir = "docs/Generated_paper/question_num_significance_test/paper"
        output_dir = f"docs/revised_shatin/{selected_model}"
        
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        batch_process_excel_files(input_dir, output_dir, selected_model, is_thinking_model, max_iterations=5)
        
        return selected_model
        
    elif mode_choice == "2":
        # 选择特定voting machine组
        print("\nAvailable voting machine groups:")
        for i in range(1, 5):
            config = get_experiment_config(i)
            print(f"{i}. {config['name']} (Threshold: {config['threshold']})")
        
        while True:
            try:
                group_choice = input("Select experiment group (1-4): ").strip()
                experiment_group = int(group_choice)
                if 1 <= experiment_group <= 4:
                    break
                else:
                    print("Invalid choice. Please enter 1-4.")
            except ValueError:
                print("Please enter a valid number.")
        
        # 设置路径
        input_dir = "docs/Generated_paper/question_num_significance_test/paper"
        output_base_dir = "docs/revised_shatin"
        
        batch_process_with_voting(input_dir, output_base_dir, experiment_group, max_iterations=5)
        
        return f"voting_group_{experiment_group}"
        
    elif mode_choice == "3":
        # 运行所有4个voting machine组进行对比
        input_dir = "docs/Generated_paper/question_num_significance_test/paper"
        output_base_dir = "docs/revised_shatin"
        
        print("Running all 4 voting machine groups for comparison...")
        
        for group in range(1, 5):
            print(f"\n{'='*60}")
            print(f"Running Experiment Group {group}")
            print(f"{'='*60}")
            
            start_time = time.time()
            batch_process_with_voting(input_dir, output_base_dir, group, max_iterations=5)
            end_time = time.time()
            
            print(f"Group {group} completed in {end_time - start_time:.2f} seconds")
        
        print("\nAll voting machine groups completed!")
        return "all_voting_groups"

if __name__ == "__main__":
    main()
