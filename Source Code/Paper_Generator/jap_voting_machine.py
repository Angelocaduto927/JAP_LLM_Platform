'''
Voting Machine for Question Review
'''
import os
import time
import re
import string
from typing import List, Dict, Tuple, Any

from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI


def create_model_instance(model_config: Dict) -> ChatOpenAI:
    """Create a ChatOpenAI instance with given configuration"""
    # 定义支持 thinking 模式的模型列表
    thinking_models = ["qwen3-235b-a22b-thinking-2507", "qwen3-next-80b-a3b-thinking"]

    # 基础配置
    config = {
        "temperature": 0.3,
        "model": model_config["model_name"],
        "api_key": os.getenv("DASHSCOPE_API_KEY"),
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "max_tokens": 2048,
        "max_retries": 5
    }
    
    if model_config["model_name"] in thinking_models:
        config["extra_body"] = {
            "enable_thinking": True,
            "thinking_budget": 2000
        }
    else:
        config["extra_body"] = {
            "enable_thinking": False
        }
    
    return ChatOpenAI(**config)

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
            If at least one question has multiple valid correct answers, respond with the question numbers that potentially have the problem only, the form requirement is number + question type (eg. 'Q8: もんだい1, Q7: もんだい2').\n
            If not, you must return "False" only.\n
            """
            ),
        ]
    )
    # chain = LLMChain(llm=llm, prompt=prompt)
    input_data = {'input_data': text}
    chain = prompt | llm
    result = api_call_for_error_detection(chain, input_data)
    if result.content != "False" and result.content is not None:
        result = result.content
        return result
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
            )
        ]
    )
    chain = prompt | llm
    input_data = {'input_data': text}
    result = api_call_for_error_detection(chain, input_data)
    if result.content != "False" and result.content is not None:
        result = result.content
        return result
    return False

def has_duplicate_options(text):
    questions_with_options = re.findall(
        r'(\d+)\.\s*(.*?)\n(1\.\s*(.*?)\n)(2\.\s*(.*?)\n)(3\.\s*(.*?)\n)(4\.\s*(.*?)\n)',
        text, re.DOTALL
    )
    for question, _, opt1, _, opt2, _, opt3, _, opt4, _ in questions_with_options:
        options = {opt1.strip(), opt2.strip(), opt3.strip(), opt4.strip()}
        if len(options) < 4:
            print(f"Duplicate options detected in question {question}: {opt1.strip()}, {opt2.strip()}, {opt3.strip()}, {opt4.strip()}")
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
            print(f"Duplicate question detected: {question_text} with options {options}")
            return True
        seen_questions.add(normalized_question)
    return False

def log_message(message: str, log_file: str = None):
    """
    Print message to console and optionally save to log file
    
    Args:
        message: The message to log
        log_file: Optional path to log file
    """
    print(message)
    if log_file:
        with open(log_file, 'a', encoding='utf-8') as f:
            f.write(f"{message}\n")

def get_model_votes(text: str, models: List[Dict], threshold: float = 0.5, log_file: str = None) -> Tuple[bool, List[int], List[Dict]]:
    """
    使用多个模型对问题进行投票判断。
    
    Args:
        text: 要检查的题目文本
        models: 包含多个模型配置的列表
        threshold: 判定为错误所需的投票比例阈值(0-1之间)
    
    Returns:
        tuple: (是否有错误, 投票结果列表, 具体评论)
    """
    votes = []
    comments = []
    
    for model_config in models:
        llm = create_model_instance(model_config)
        result = check_for_error(text, llm)
        
        log_message(f"\nChecking with model: {model_config['model_name']}", log_file)
        log_message(f"Check result: {result}", log_file)
        
        if result:
            votes.append(1)  # 有错误投1票
            comments.append({
                "model": model_config["model_name"],
                "errors": result,
                "weight": model_config.get("weight", 1.0)
            })
        else:
            votes.append(0)  # 无错误投0票
            
    # 计算加权投票比例
    weighted_votes = sum(v * m.get("weight", 1.0) for v, m in zip(votes, models))
    total_weights = sum(m.get("weight", 1.0) for m in models)
    vote_ratio = weighted_votes / total_weights
    
    has_error = vote_ratio >= threshold
    log_message(f"\nVoting Results:", log_file)
    log_message(f"Vote ratio: {vote_ratio:.2f}", log_file)
    log_message(f"Threshold: {threshold}", log_file)
    log_message(f"Final decision: {'Has errors' if has_error else 'No errors'}", log_file)
    
    return has_error, votes, comments

def aggregate_errors(comments: List[Dict], log_file: str = None) -> List:
    """聚合多个模型的错误检测结果"""
    error_counts = {}
    total_models = len(comments)
    
    log_message("\nAggregating errors from all models:", log_file)
    
    # 首先整理所有问题
    for comment in comments:
        if "errors" not in comment:
            continue
            
        log_message(f"\nModel {comment['model']} found errors:", log_file)
        for error in comment["errors"]:
            if not isinstance(error, list) or len(error) != 2:
                continue
                
            error_type = error[0]  # "Multiple correct answers" or "Stem errors"
            error_content = error[1].strip()  # 去除空白字符
            
            # 跳过空字符串
            if not error_content:
                log_message(f"- {error_type}: (empty content, skipping)", log_file)
                continue
                
            questions = error_content.split(',')  # Split question numbers
            questions = [q.strip() for q in questions if q.strip()]  # Clean up whitespace and remove empty
            
            # 为每个问题创建单独的计数
            for question in questions:
                error_key = f"{error_type}_{question}"
                if error_key not in error_counts:
                    error_counts[error_key] = {
                        "type": error_type,
                        "question": question,
                        "count": 0,
                        "models": []
                    }
                error_counts[error_key]["count"] += 1
                error_counts[error_key]["models"].append(comment["model"])
                log_message(f"- {error_type}: {question}", log_file)
    
    # 计算阈值并汇总共识错误
    threshold = total_models / 2
    consensus_questions = {
        "Multiple correct answers": set(),
        "Stem errors": set()
    }
    
    # 收集达到阈值的错误
    for info in error_counts.values():
        if info["count"] > threshold:
            consensus_questions[info["type"]].add(info["question"])
    
    # 构建最终的错误列表
    consensus_errors = []
    for error_type, questions in consensus_questions.items():
        if questions:
            consensus_errors.append([error_type, ", ".join(sorted(questions))])
    
    log_message(f"\nConsensus errors (agreed by > {threshold:.1f} models):", log_file)
    for error in consensus_errors:
        log_message(f"- {error[0]}: {error[1]}", log_file)
    
    return consensus_errors

def check_for_error(revised_text, llm, log_file: str = None):
    errors = []
    try:
        Multiple_correct_answers = has_multiple_correct_answers(revised_text, llm)
        if Multiple_correct_answers != False:
            errors.append(["Multiple correct answers", Multiple_correct_answers])
            log_message(f"Found multiple correct answers: {Multiple_correct_answers}", log_file)
            
        if has_duplicate_questions(revised_text):
            errors.append("Duplicate questions")
            log_message("Found duplicate questions", log_file)
            
        Stem_question = has_stem_errors(revised_text, llm)
        if Stem_question != False:
            errors.append(["Stem errors", Stem_question])
            log_message(f"Found stem errors: {Stem_question}", log_file)
            
        if has_duplicate_options(revised_text):
            errors.append("Duplicate options")
            log_message("Found duplicate options", log_file)
            
        return errors
    except Exception as e:
        error_msg = f"Error in check_for_error: {e}"
        log_message(error_msg, log_file)
        return []

