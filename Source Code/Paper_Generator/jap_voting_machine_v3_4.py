"""
Voting Machine for Japanese Question Revision
实现4组不同配置的voting machine
"""

import os
import re
import time
import json
from typing import List, Dict, Tuple, Any, Set
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI

def get_experiment_config(experiment_group: int) -> Dict:
    """获取不同实验组的模型配置"""
    
    configs = {
        1: {  # 对照组
            "name": "Control Group",
            "threshold": 0.65,
            "models": [
                {"model_name": "qwen3-vl-235b-a22b-thinking", "weight": 29, "is_thinking": True},
                {"model_name": "qwen3-235b-a22b-thinking-2507", "weight": 20, "is_thinking": True},
                {"model_name": "qwen3-30b-a3b-thinking-2507", "weight": 13, "is_thinking": True},
                {"model_name": "qwen3-next-80b-a3b-instruct", "weight": 17, "is_thinking": False},
                {"model_name": "qwen3-30b-a3b-instruct-2507", "weight": 11, "is_thinking": False},
                {"model_name": "qwen-flash", "weight": 10, "is_thinking": True}
            ]
        },
        2: {  # 高质量组
            "name": "High Quality Group", 
            "threshold": 0.61,
            "models": [
                {"model_name": "qwen3-vl-235b-a22b-thinking", "weight": 39, "is_thinking": True},
                {"model_name": "qwen3-235b-a22b-thinking-2507", "weight": 26, "is_thinking": True},
                {"model_name": "qwen3-next-80b-a3b-instruct", "weight": 20, "is_thinking": False},
                {"model_name": "qwen3-30b-a3b-thinking-2507", "weight": 15, "is_thinking": True}
            ]
        },
        3: {  # 快速组
            "name": "Fast Group",
            "threshold": 0.57,
            "models": [
                {"model_name": "qwen3-30b-a3b-thinking-2507", "weight": 19, "is_thinking": True},
                {"model_name": "qwen3-next-80b-a3b-instruct", "weight": 30, "is_thinking": False},
                {"model_name": "qwen3-30b-a3b-instruct-2507", "weight": 19, "is_thinking": False},
                {"model_name": "qwen-flash", "weight": 19, "is_thinking": True},
                {"model_name": "qwen3-30b-a3b", "weight": 13, "is_thinking": False}
            ]
        },
        4: {  # 平衡组
            "name": "Balanced Group",
            "threshold": 0.48,
            "models": [
                {"model_name": "qwen3-vl-235b-a22b-thinking", "weight": 33, "is_thinking": True},
                {"model_name": "qwen3-235b-a22b-thinking-2507", "weight": 24, "is_thinking": True},
                {"model_name": "qwen3-next-80b-a3b-instruct", "weight": 19, "is_thinking": False},
                {"model_name": "qwen3-30b-a3b-instruct-2507", "weight": 12, "is_thinking": False},
                {"model_name": "qwen-flash", "weight": 12, "is_thinking": True}
            ]
        }
    }
    
    return configs.get(experiment_group, configs[1])

def create_llm(model_name: str, is_thinking: bool, temperature: float = 0.3) -> ChatOpenAI:
    """创建LLM实例"""
    if is_thinking:
        return ChatOpenAI(
            temperature=temperature,
            model=model_name,
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
            max_tokens=2048,
            max_retries=3,
            streaming=True,
            extra_body={"thinking_budget": 1000}
        )
    else:
        return ChatOpenAI(
            temperature=temperature,
            model=model_name,
            api_key=os.getenv("DASHSCOPE_API_KEY"),
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
            max_tokens=2048,
            max_retries=3,
            extra_body={"enable_thinking": False}
        )

def api_call_for_voting(chain, input_data, is_streaming=False, max_retries=3, delay=1):
    """处理投票时的API调用，支持流式传输"""
    retries = 0
    while retries < max_retries:
        try:
            if is_streaming:
                # 对于流式调用，需要收集所有块
                response_chunks = []
                for chunk in chain.stream({"input_data": input_data}):
                    if hasattr(chunk, 'content'):
                        response_chunks.append(chunk.content)
                
                # 创建一个类似于非流式响应的对象
                class StreamResponse:
                    def __init__(self, content):
                        self.content = content
                
                return StreamResponse(''.join(response_chunks))
            else:
                return chain.invoke({"input_data": input_data})
        except Exception as e:
            print(f"API call failed: {e}. Retrying in {delay} seconds...")
            retries += 1
            time.sleep(delay)
    print("Max retries reached. Giving up.")
    return None

def extract_question_numbers(text: str) -> List[str]:
    """从模型输出中提取题号列表"""
    if text == "False" or not text.strip():
        return []
    
    pattern = r'Q\d+:\s*もんだい\d+'
    matches = re.findall(pattern, text)
    
    cleaned = []
    seen = set()
    for m in matches:
        normalized = re.sub(r'\s+', ' ', m).strip()
        if normalized not in seen:
            cleaned.append(normalized)
            seen.add(normalized)
    return cleaned

def get_model_vote(text: str, llm: ChatOpenAI, error_type: str, is_thinking: bool = False) -> Tuple[bool, List[str]]:
    """获取单个模型对特定错误类型的投票，返回具体题号列表"""
    
    if error_type == "multiple_correct_answers":
        prompt = ChatPromptTemplate.from_messages([
            ("system", "You are an experienced Japanese N4/N5 examiner reviewing the following multiple-choice questions:\n\n"),
            ("human", 
            """{input_data}\n\n
            Check if any question has **more than one correct answer**. This means that multiple options are valid for the question given its context.\n
            If at least one question has multiple valid correct answers, respond with the question numbers that potentially have the problem only, the form requirement is number + question type (eg. 'Q8: もんだい1, Q7: もんだい2' **questions separated by commas**).\n
            If not, you must return "False" only.\n
            
            Make sure the number of questions remains exactly the same as the input and the question index (the Qx: もんだいy part) for each question is unchanged.
            """
            ),
        ])
    elif error_type == "stem_errors":
        prompt = ChatPromptTemplate.from_messages([
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
        ])
    
    try:
        chain = prompt | llm
        # 使用支持流式传输的API调用
        result = api_call_for_voting(chain, text, is_streaming=is_thinking)
        
        if result is None:
            return False, []
        
        response = result.content.strip()
        # 提取题号列表
        question_numbers = extract_question_numbers(response)
        
        has_error = response != "False" and response.strip() != ""
        
        return has_error, question_numbers
    except Exception as e:
        print(f"Error getting vote from model: {e}")
        return False, []

def get_model_revision_for_question(question_text: str, llm: ChatOpenAI, error_type: str, is_thinking: bool = False) -> str:
    """获取模型对特定题目的修改建议"""
    
    prompt = ChatPromptTemplate.from_messages([
        ("human", f'''
        You are an experienced Japanese N4/N5 examiner. There is a Japanese multiple-choice question at the end of this message that has issues with {error_type}. 
        Your task is to modify this question to fix the issues:

        1. No duplicate options: All four options within the question should be unique.  
        2. No multiple reasonable answers: Ensure that only one answer is correct and reasonable.
        3. Grammatical correctness: The title and stem must be grammatically correct.
        4. Clear instructions: Make sure the question is clear and unambiguous.

        Output Format: The question must **keep the original format**:
        - The question must start with the same `Qx: もんだいy\n` as the original
        - The question must have exactly 4 options (`1` to `4`).
        - The question must have an `Answer: x` at the end.
        - Do not include any other comments.
        
        Here is the question to revise:
        {question_text}
        ''')
    ])
    
    try:
        chain = prompt | llm
        result = api_call_for_voting(chain, question_text, is_streaming=is_thinking)
        
        if result is None:
            return question_text  # 如果失败，返回原始问题
        
        return result.content.strip()
    except Exception as e:
        print(f"Error getting revision: {e}")
        return question_text

def get_voting_result(text: str, experiment_group: int) -> Dict[str, Dict]:
    """
    获取每个题目的投票结果
    每个模型对每个题目只计算一次权重，而不是按错误类型累加
    """
    config = get_experiment_config(experiment_group)
    models = config["models"]
    threshold = config["threshold"]
    
    print(f"Running voting with {config['name']} (Group {experiment_group})")
    print(f"Models: {len(models)}, Threshold: {threshold}")
    
    # 按题目编号跟踪投票
    question_votes = {}
    
    # 收集每个模型的投票
    for model_info in models:
        model_name = model_info["model_name"]
        weight = model_info["weight"]
        is_thinking = model_info["is_thinking"]
        
        print(f"Getting vote from {model_name} (weight: {weight})...")
        
        try:
            llm = create_llm(model_name, is_thinking)
            
            # 检查多个正确答案
            has_multiple, questions_multiple = get_model_vote(text, llm, "multiple_correct_answers", is_thinking)
            
            # 检查题干错误
            has_stem, questions_stem = get_model_vote(text, llm, "stem_errors", is_thinking)
            
            # 记录每个题目的投票情况
            all_questions = set(questions_multiple + questions_stem)  # 使用集合去重
            
            print(f"  - Found issues in {len(all_questions)} questions")
            if all_questions:
                print(f"  - Flagged questions: {', '.join(all_questions)}")
            
            # 更新每个题目的权重 - 修改为每个模型只计算一次权重
            for q in all_questions:
                if q not in question_votes:
                    question_votes[q] = {"error_types": [], "total_weight": 0, "models": []}
                
                # 添加模型权重（每题每模型只加一次）
                if model_name not in question_votes[q]["models"]:
                    question_votes[q]["total_weight"] += weight
                    question_votes[q]["models"].append(model_name)
                
                # 只记录错误类型（不重复累加权重）
                if q in questions_multiple and "multiple_correct_answers" not in question_votes[q]["error_types"]:
                    question_votes[q]["error_types"].append("multiple_correct_answers")
                
                if q in questions_stem and "stem_errors" not in question_votes[q]["error_types"]:
                    question_votes[q]["error_types"].append("stem_errors")
            
        except Exception as e:
            print(f"Error with model {model_name}: {e}")
    
    # 标记超过阈值的题目
    total_weight = sum(m["weight"] for m in models)
    threshold_value = total_weight * threshold
    
    for q, vote_info in question_votes.items():
        vote_info["exceeds_threshold"] = vote_info["total_weight"] >= threshold_value
        vote_info["weight_ratio"] = vote_info["total_weight"] / total_weight
    
    print(f"\nVoting results:")
    print(f"Total weight: {total_weight}, Threshold: {threshold_value} ({threshold*100}%)")
    
    for q, info in sorted(question_votes.items()):
        status = "FLAGGED" if info["exceeds_threshold"] else "ignored"
        print(f"  - {q}: {info['weight_ratio']:.2%} ({info['total_weight']}/{total_weight}) {status}")
    
    return question_votes
