'''Question_Generator'''
import re
import os
import time
import string

from docx import Document
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI
from langchain.prompts import ChatPromptTemplate  

from jap_excel_processor import parse_questions
from jap_excel_processor import store_questions_to_excel
from jap_excel_processor import process_word_to_excel

def split_into_sentences(text):
    sentence_endings = re.compile(r'(?<=[。！？])\s*')
    sentences = sentence_endings.split(text)
    return sentences


def extract_numbered_content(file_path, start_number, end_number):
    """
    Extracts numbered content (e.g., 1-8) from a Word document, including multiple-choice options and lists with different formats.

    Args:
        file_path (str): The path to the Word document.
        start_number (int): The starting number of the range to extract.
        end_number (int): The ending number of the range to extract.

    Returns:
        tuple: A tuple containing two lists:
            - num: A list of question numbers (e.g., [1, 2])
            - content_list: A list of lists of options (e.g., [['a. くださる', 'b. いただく'], ['a. 経験', 'b. あつい']])
    """
    # Load the Word document
    doc = Document(file_path)
    
    # Compile a regex pattern to match the numbered entries with or without additional details like kanji readings in brackets
    pattern = re.compile(rf"^\s*(\d+)\.\s*([^【]*)(?:【([^】]*)】)?\s*(.*)")
    
    # Lists to store the extracted content
    num = []
    content_list = []

    # Define a function to sanitize strings for file names
    def sanitize_filename(content):
        # Remove invalid characters for filenames (e.g., \ / : * ? " < > |)
        return re.sub(r'[\\/*?:"<>|]', '_', content).strip()

    # Iterate through all paragraphs in the document
    for para in doc.paragraphs:
        match = pattern.match(para.text)
        if match:
            number = int(match.group(1))
            item = match.group(2).strip()
            kanji_reading = match.group(3) if match.group(3) else ""
            additional_text = match.group(4).strip()

            # Prepare the full item content (item + kanji reading if exists)
            full_item = item + ("【" + kanji_reading + "】" if kanji_reading else "")
            full_item = full_item.strip()

            # Sanitize the full item to make it safe for filenames
            sanitized_full_item = sanitize_filename(full_item)

            # Check if the number is within the specified range
            if start_number <= number <= end_number:
                # If the number is already in the list, add to the content list
                if number not in num:
                    num.append(number)
                    content_list.append([f"{sanitized_full_item} {additional_text}"])
                else:
                    index = num.index(number)
                    content_list[index].append(f"{sanitized_full_item} {additional_text}")

    return num, content_list




# ---------------------------
# 第一部分：题目生成相关函数
# ---------------------------
def generate_prompt(knowledge_point: str, question_format: int, num_questions: int) -> str:
    """生成题目生成的prompt文本"""
    if question_format not in (1, 2):
        raise ValueError("question_format must be 1 or 2")
    if not isinstance(num_questions, int) or num_questions <= 0:
        raise ValueError("num_questions must be a positive integer")
    
    interference_rules = {
        1: "The grammar point is what should be filled into the blank space (i.e., the blank must be completed with this grammar point or its correct conjugation).",
        2: "The grammar point already appears in the question stem, and you must use it to either complete another blank or choose the most natural continuation (i.e., testing understanding and application of the grammar point)."
    }

    
    difficulty_levels = {
        1: "Basic difficulty: simple sentence structure, clear context, and direct application of grammar points.",
        2: "Intermediate difficulty: slightly complex sentence structure, may contain multiple grammatical elements, and requires certain analytical skills.",
        3: "Advanced difficulty: complex sentence structure, the context has certain implicit information, and requires a deep understanding of grammar points and context."
    }
    
    validation_prompt = """
        Generated questions must undergo the following quality validations:
        1. Grammatical accuracy: The grammar of the stems and options must be correct, natural and meet the standards of Japanese N4/N5 level.
        2. Distractor validity: Distractors should be plausible enough to confuse learners, but must be grammatically incorrect according to standard Japanese grammar rules (excluding cultural nuances or non-standard colloquial usages). Distractors may be designed to look very similar to the correct form, but they must represent non-existent or ungrammatical constructions in standard Japanese. In other words, distractors should be misleading at first glance, yet clearly incorrect under standard grammar rules.
        3. Reasonable difficulty: Design questions according to different difficulty levels to ensure a reasonable distribution of difficulty.
        4. Context authenticity: Sentence scenarios must meet the following requirements:
            - Context authenticity: Sentence scenarios must be sufficiently complete and detailed so that learners can perform thorough logical analysis. Only with enough contextual information can distractors be eliminated reliably based on standard grammatical rules, rather than guesswork
            - Daily conversations (such as chatting with friends, shopping)
            - Common exam scenarios (such as email writing, schedule planning)
            - Avoid artificial contexts (such as science fiction or professional fields)
"""
    
    base_count = max(1, int(num_questions * 0.4))
    intermediate_count = max(1, int(num_questions * 0.4))
    advanced_count = num_questions - base_count - intermediate_count
    if advanced_count < 0:
        base_count -= 1
        intermediate_count -= 1
        advanced_count = num_questions - base_count - intermediate_count
    
    prompt = (
        f"You are an experienced Japanese examiner for JLPT N4/N5. Create exactly {num_questions} questions "
        f"for the grammar point: **{knowledge_point}**.\n\n"
        f"Question Format {question_format}: {interference_rules.get(question_format, '')}\n\n"
        f"Generate questions with the following difficulty breakdown: "
        f"Basic ({base_count} questions), Intermediate ({intermediate_count} questions), Advanced ({advanced_count} questions):\n\n"
        f"{difficulty_levels[1]}\n{difficulty_levels[2]}\n{difficulty_levels[3]}\n\n"
        f"{validation_prompt}\n\n"
        f"Instructions:\n"
        f"1. Each question must start with a header: 'Qx: もんだい{question_format}'.\n"
        f"2. Provide exactly 4 options (1-4) in one line.\n"
        f"3. End each question with 'Answer: x'.\n"
        f"4. The question stem must always contain a pair of parentheses ( ) to indicate the blank where the option should be filled in and it must be empty in that parentheses (don't fill in the answer).\n"
        f"5. No extra text or formatting!\n\n"
        f"Example Format:\n"
        f"Q1: もんだい{question_format}\n"
        f"[Japanese Question Stem]\n"
        f"1. Option1 2. Option2 3. Option3 4. Option4\n"
        f"Answer: x\n"
    )
    return prompt

def generate_grammar_questions(knowledge_point: str, question_format: int, num_questions: int) -> str:
    """使用 LLM 生成语法题目"""
    message = generate_prompt(knowledge_point, question_format, num_questions)
    prompt = ChatPromptTemplate.from_messages([("human", "{message}")])
    llm = ChatOpenAI(
        api_key = os.getenv("DASHSCOPE_API_KEY"),
        base_url="https://dashscope.aliyuncs.com/compatible-mode/v1",
        model = "qwen3-235b-a22b-thinking-2507",
        temperature=0.7,
        max_tokens = 2048,
        max_retries = 5,
        extra_body = {"thinking_budget": 1000},
    )
    chain = prompt | llm
    try:
        return chain.invoke({"message": message}).content
    except Exception as e:
        print(f"Error generating questions: {e}")
        return ""

def split_sentences(text, question_counter):
    text = re.sub(r'Q\d+:\s*', '', text)
    problems = re.split(r'(もんだい\d+)', text)
    problems = [p.strip() for p in problems if p.strip()]
    result = []
    for part in problems:
        if part.startswith('もんだい'):
            result.append(f"Q{question_counter}: " + part)
            question_counter += 1
        else:
            result.append(part)
    return result, question_counter

def process_and_revise_document(doc_path: str):
    """调整文档中题目的格式"""
    doc = Document(doc_path)
    question_counter = 1
    for paragraph in doc.paragraphs:
        sentences, question_counter = split_sentences(paragraph.text, question_counter)
        p = paragraph._element
        for child in list(p):
            p.remove(child)
        for sentence in sentences:
            run = paragraph.add_run(sentence)
            if "Answer:" in sentence:
                paragraph.add_run("\n")
    doc.save(doc_path)

# ---------------------------
# 第二部分：题目检查与修订相关函数
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
    # 将试卷题目分割成列表，方便后续抽取和替换题目进行修改
    questions = re.findall(
        r'(Q\d+:\s*もんだい\d+\s*.*?Answer:\s*\d+)',
        text, re.DOTALL
    )
    return questions

def normalize_spaces(text):
    # 将全角空格替换为半角空格
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
        return ["Unexpected error in check_for_error"]

def extract_wrong_questions(question_list_input, error):
    errors = error
    question_list = question_list_input
    multicorrect_problems_number = []
    stem_problems_number = []
    if [q for q in errors if q[0] == "Multiple correct answers"] != []:
        multicorrect_problems_number = [q for q in errors if q[0] == "Multiple correct answers"][0][1].split(',')
        multicorrect_problems_number = [q.strip() for q in multicorrect_problems_number]
    if [q for q in errors if q[0] == "Stem errors"] != []:
        stem_problems_number = [q for q in errors if q[0] == "Stem errors"][0][1].split(',')
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
        else:
            continue
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

def question_revise_simple(rows, filename, output_dir, revised_newpaper_folder, max_iterations=5, model='qwen3-235b-a22b-thinking-2507', temperature=0.6):
    """
    对新生成的日语练习题进行多轮修订和检查，确保题目质量。
    """
    llm_revise = ChatOpenAI(temperature=temperature, model=model, api_key=os.getenv("DASHSCOPE_API_KEY"),base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", max_tokens=2048, max_retries=5, extra_body={"thinking_budget": 1000})
    prompt_revise = ChatPromptTemplate.from_messages(
        [
            ("human", 
            '''
            You are an experienced Japanese N4/N5 examiner. There are some Japanese multiple-choice questions at the end of this message. All of the questions have some issues. You can refer to {errors}. 
            Your task is to modify those multiple-choice test questions to meet the following criteria and fix all the issues:

            1. No duplicate questions: Ensure that all questions are unique. If a question is repeated or too similar to another, please replace it with a new question with a distinct structure or context. Provide specific suggestions on how to modify repeated questions.

            2. No duplicate options: All four options within a question should be unique. Distractors should be plausible but incorrect. If necessary, suggest how to modify similar options to increase their clarity.

            3. No multiple reasonable answers: Ensure that only one answer is correct and reasonable under the stem scenario. If more than one option are acceptable, modify the question or options to clarify the correct choice. Provide specific suggestions on how to make the answer clear and unambiguous.

            4. Grammatical correctness: The title and stem of each question must be grammatically correct. Review for unnatural sentence structures and revise them to ensure fluency and correctness. If you detect any grammatical errors, please explain how to fix them.

            5. Relevance of options: Ensure that the stem clearly indicates what cannot be chosen. To achieve this, ensure the scenario in the stem is clear and detailed. Avoid culturally biased content. Suggest how to improve the incorrect options by reflecting common mistakes learners make.

            6. Pronunciation and Word Usage: If the question involves pronunciation, katakana, or hiragana forms, the Japanese word should be enclosed in brackets for clarity. For hiragana or katakana conversion questions, ensure that the word is written in the correct form, and the correct answer is not shown in the question stem. Also, check for spelling inconsistencies.

            7. General guidance: Eliminate any ambiguity, revise unclear options, and avoid subjective or culturally biased phrasing. Ensure all questions are at an appropriate difficulty level for the target JLPT level (N4/N5). Avoid complex words or structures outside the typical N4/N5 range.

            8. Output Format: Each question must **keep the original format**:
            - Each question must start with `Qx: もんだいy\n` (x,y are an integer number, where x represents the question number and y represents the question type). The question number x and format type y should remain exactly the same as in the original question being revised. (eg. the question you get is Q3: もんだい2, then the corresponding revised question must also be Q3: もんだい2)
            - Each question must have exactly 4 options (`1` to `4`).
            - Each question must have an `Answer: x` at the end of it (x is an integer number).
            - Each question must contain a pair of parentheses ( ) to indicate the blank where the option should be filled in and must be empty in that parentheses.
            - Do not include any other comments.
            
            
            Here are the questions to review and modify:
            {input_data}
            '''
            ),
        ]
    )
    llm_error_check = ChatOpenAI(temperature=0.3, model=model, api_key=os.getenv("DASHSCOPE_API_KEY"),base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", max_tokens=2048, max_retries=5, extra_body={"thinking_budget": 2000})
    chain = prompt_revise | llm_revise
    revised_result = rows
    params = {
        'input_data': "",
        'errors': []
    }
    question_list = paper_split_into_list(normalize_spaces(revised_result))
    for iteration in range(max_iterations):
        errors = check_for_error(revised_result, llm_error_check)
        if not errors:
            print(f"No issues found after {iteration + 1} iterations.")
            break
        print(f"Iteration {iteration + 1}: Detected errors - {errors}")
        
        # 提取判断为错误的题目，进行下一轮修订
        wrong_questions = extract_wrong_questions(question_list, errors)
        print(f"Questions to be revised: {wrong_questions}")
        formatted_wrong_questions = '\n'.join(['\n'.join(sublist) for sublist in wrong_questions])
        params['input_data'] = formatted_wrong_questions
        params['errors'] = errors
        
        revised_result = api_call_for_paper_revise(chain, params)
        revised_result = revised_result.content if revised_result else None
        if revised_result is None:
            print("Failed to get revised result. Stopping iteration.")
            break
        
        # 合并题目和文档
        revised_result = combine_revised_questions_into_paper(revised_result, question_list)
        revised_result = '\n'.join(revised_result)
        question_list = paper_split_into_list(normalize_spaces(revised_result))
        
        # 保存每次迭代的中间结果和日志
        intermediate_path = os.path.join(output_dir, f"{filename}_iteration_{iteration + 1}.docx")
        output_doc = Document()
        sentences = split_into_sentences(revised_result)
        for sentence in sentences:
            output_doc.add_paragraph(sentence)
        output_doc.save(intermediate_path)
        log_path = os.path.join(output_dir, f"{filename}_error_log.txt")
        with open(log_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"Iteration {iteration + 1} Errors: {errors}\n")
    else:
        print(f"Maximum iterations ({max_iterations}) reached. Errors may still exist.")
    
    output_path = os.path.join(revised_newpaper_folder, f"{filename}_revised.docx")
    output_doc = Document()
    sentences = split_into_sentences(revised_result)
    for sentence in sentences:
        output_doc.add_paragraph(sentence)
    output_doc.save(output_path)
    
    qa_list = parse_questions(revised_result)
    excel_filename = f"{filename}.xlsx"
    store_questions_to_excel(qa_list, output_dir, excel_filename)

# ---------------------------
# 整合工作流：生成 -> 检查 -> 存储
# ---------------------------
def generate_check_store_pipeline(grammar_list, num_list, output_dir, revised_newpaper_folder, base_filepath):
    """
    对每个语法知识点：
    1. 生成题目（按5种题型各生成一定数量的题目）
    2. 合并生成的题目文本
    3. 进行题目检查和多轮修订
    4. 存储最终修订结果（docx、excel）
    """
    filename = os.path.splitext(os.path.basename(base_filepath))[0]
    
    # 遍历每个知识点
    for knowledge_point, question_number in zip(grammar_list, num_list):
        print(f"Processing knowledge point: {knowledge_point}")
        all_questions = ""
        # 生成每种题型的题目并合并
        for question_format in range(1, 3):
            print(f"Generating questions for format {question_format}...")
            questions_text = generate_grammar_questions(knowledge_point, question_format, 10)
            all_questions += f"\n### Format {question_format}\n" + questions_text
        
        # 保存初步生成的题目到一个临时文件（可选）
        temp_path = os.path.join(output_dir, f"{filename}_{question_number}_{knowledge_point}_generated.docx")
        doc = Document()
        doc.add_paragraph(all_questions)
        doc.save(temp_path)
        
        # 调整格式（如需要）
        process_and_revise_document(temp_path)
        
        # 调用题目检查与修订流程
        question_revise_simple(all_questions, f"{filename}_{question_number}_{knowledge_point}", output_dir, revised_newpaper_folder)
        
        print(f"Finished processing {knowledge_point}. Results stored in {revised_newpaper_folder}")




def main():
    test_grammar_original = "JAP_LLM_Platform/docs/test_grammar.docx"
    revised_output_grammar = "JAP_LLM_Platform/docs/Generated_paper/revised_grammar_questions"
    grammar_output = "JAP_LLM_Platform/docs/Generated_paper/original_grammar_questions"

    grammar_num = extract_numbered_content(test_grammar_original, 1, 4)[0]
    grammar_test = extract_numbered_content(test_grammar_original, 1, 4)[1]

    print(grammar_test)

    generate_check_store_pipeline(grammar_test, grammar_num, grammar_output, revised_output_grammar, test_grammar_original)
    process_word_to_excel(revised_output_grammar, revised_output_grammar)


if __name__ == "__main__":
    main()
