'''Question_Generator'''
import re
import os
import glob
import time
import string
import warnings
import docx
import mysql.connector
from docx import Document
from langchain.text_splitter import RecursiveCharacterTextSplitter
from langchain_core.prompts import ChatPromptTemplate
from langchain_openai import ChatOpenAI
from langchain.prompts import ChatPromptTemplate  
from langchain.chains import LLMChain   
from langchain_community.document_loaders import UnstructuredWordDocumentLoader
from typing import Any
from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.worksheet.datavalidation import DataValidation

# from jap_paper_revise import produce_new_question_list
from jap_paper_revise import read_docx_to_string_with_format

from jap_vocabulary_processor import vocabulary_points_revise
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
    # if question_format not in (1, 2, 3, 4, 5, 6, 7, 8):
    #     raise ValueError("question_format must be between 1 and 8")
    if question_format not in (1,):
        raise ValueError("question_format must be 1")
    if not isinstance(num_questions, int) or num_questions <= 0:
        raise ValueError("num_questions must be a positive integer")
    
    format_descriptions = {
    1:"""[もんだい1: Connector and Conjunction Selection]
        These items require you to choose expressions that link clauses by showing cause, sequence, or contrast.
        Examples:
        1. かれが　手伝って (   )　宿題 (しゅくだい)が終わらなかった.
            1 もらったから		2 くれなかったから		3 ほしいから		4 ほしかったから  
            Answer: 2
        Explanation: Tests the reason (“〜から”) behind the outcome.
        
        2. 日よう日は　道が　こむので　(   )。
            1　月よう日も　こまなかった			2　車で　行くことにした		
            3　やくそくの　時間に　間に合った		4　月よう日に　行くことにした
            Answer: 4
        Explanation: Requires linking a cause (crowded roads) with the consequence (choice of travel mode or time).

        3. 雨が　少ない　(   )、　やさいが　大きくなりません。
            1　より			2　すぎて		3　ため		4　けど
            Answer: 3
        Explanation: Needs a connector that indicates cause or reason.
    """,

    2:"""[もんだい2: Particle and Case Marker Selection]
        These questions focus on choosing the correct particles or markers that indicate relationships such as location, means, or quantification.
        Examples:
        1. 3時間だけ　仕事を　したら　10,000円　(   )　もらえた。
            1　し			2　に			3　も			4　で
            Answer: 3
        Explanation: Tests which particle correctly links the amount/method.

        2. 子ども	「お母さん、来週　着る　服を　あらって　(   )。」
	        母	「自分で　あらいなさい。」
            1　おく			2　ある			3　おいて		4　あって
            Answer: 3
        Explanation: Involves the proper form after a compound verb or adverbial expression.

        3.　1か月　(   )　5本　映画を　見ます。
            1　か			2　を			3　に			4　と
            Answer: 3
        Explanation: Focuses on the marker that correctly expresses frequency or count.
    """,

    3:"""[もんだい3: Verb and Adjective Conjugation/Transformation]
        These items require the candidate to select the correct conjugated form, negative form, command form, or transformation needed in context.
        Examples:
        1.宿題 (しゅくだい) を　したのに、　先生が　(   )。
            1　来なかった				2　してしまった		
            3　会わなかった			4　するつもりだった
            Anawer: 1
        Explanation: Tests which verb form correctly contrasts action and result.

        2.うちの　子どもは　勉強 (べんきょう) しないで　(   )　ばかりいる。
            1　あそび		2　あそぶ		3　あそばない		4　あそんで
            Answer: 4
        Explanation: Focuses on the proper connective form after a negative (しないで …あそんで).

        3. 今日は　何も　(   )　出かけました。
            1　食べないで		2　食べて		3　食べなくて		4　食べても
            Answer: 1
        Explanation: Demands the correct negative construction with “何も …ない”.

        4.11時だ。　明日も　學校なんだから　子どもは　早く　(   )。
            1　ねるな		2　ねろ		3　ねすぎ		4　ねそう
            Answer: 2
        Explanation: Requires choosing the appropriate command form (e.g., “ねろ”).

        5.手紙 (てがみ) によると、　田中さんは　(   )　そうです。
            1　元気		2　元気な		3　元気だ		4　元気という
            Answer: 3
        Explanation: Involves selecting the form used for reported speech.
    """,

    4: """[もんだい4: Modal, Volition, and Decision Expressions]
        These questions test how well you express desires, intentions, possibilities, or final decisions.
        Examples:
        1. もし　1000万円　もらったら、　わたしは　いろいろな国を　(   ).
            1. 旅行したがる       2. 旅行したがっている       3. 旅行したい       4. 旅行したかった
            Answer: 3
        Explanation: Requires choosing the expression that shows desire or volition.

        2. わたしは　来年　国へ　帰る　(   ).
            1. そうだ       2. らしい       3. ようになった       4. ことにした
            Answer: 4
        Explanation: Focuses on expressing a decision or planned change.

        3. わたしは　明日　仕事で　遅れる　(   )　から…
            1. かもしれない       2. そうだ       3. らしい       4. ところだ
            Answer: 1
        Explanation: Needs a possibility or tentative marker.
    """,

    5: """[もんだい5: Comparative and Degree Expressions]
        These items ask you to select forms that set up comparisons or express degrees or extents.
        Examples:
        1. 山田さんも　背が高いが　田中さん　(   )　高くない.
            1. から       2. ほど       3. なら       4. しか
            Answer: 2
        Explanation: Tests the proper comparative structure to indicate difference or limitation.

        2. あの子は　10さいなのに、赤ちゃんの　(   )　です.
            1. ほう       2. よう       3. こと       4. もの
            Answer: 2
        Explanation: Involves choosing the correct comparative form.

        3. きょうの　テストは　先週の　テスト　(   )　むずかしくなかった.
            1. ほど       2. も       3. までに       4. ばかり
            Answer: 1
        Explanation: Requires a degree marker to indicate the extent of difference.
    """,

    6: """[もんだい6: Special Constructions (Dialogue/Politeness/Requests)]
        These questions involve selecting forms that fit within conversational contexts or polite requests.
        Examples:
        1. すみませんが　父に　何か　あったら　電話を　(   ).
            1. してくださいませんか       2. してくれてもいいですか       
            3. してもらいませんか       4. してもらうのがいいですか
            Answer: 1
        Explanation: Tests the proper form for making a polite request.

        2. A:「田中さんは　かのじょが　いますか。」
            B:「いいえ、田中さんは　前の　かのじょと　別れてから、人を好き　(   ).」
            1. ではありませんでした       2. にならなくなりました       
            3. でもよくなりました       4. にしなくなりました
            Answer: 2
        Explanation: Requires selecting the correct transformation in a conversational exchange.

        3. A:「しゅんくんの　電話番号 (でんわばんごう) を　知っている？」
            B:「わたしは　(   )　けど、はなさんなら　わかるかもしれない。」
            1. わからなかった       2. わかっていない       
            3. 知らない       4. 知っていない
            Answer: 3
        Explanation: Conveys partial or uncertain knowledge appropriately in dialogue.
    """,

    7: """[もんだい7: Potential, Resultative, and Speculative Expressions]
        These items focus on expressions that show ability (potential), completed actions leading to a result, or unexpected outcomes.
        Examples:
        1. にもつは　多くて　このかばんに　(   )　そうもない.
            1. 入り       2. 入る       3. 入ら       4. 入れない
            Answer: 1
        Explanation: Requires the selection of a form that correctly expresses potential or capacity.

        2. 毎日　(   )　ため、目が　わるくなってしまった.
            1. ゲーム       2. ゲームをしない       
            3. ゲームをした       4. ゲームがしたい
            Answer: 3
        Explanation: Tests expressing a completed habitual action leading to an undesired result.

        3. サッカーの　試合 (しあい) は　中止になると　思っていたら　(   ).
            1. 行かなかった       2. 行けそうだった       
            3. することになった       4. 中止になった
            Answer: 3
        Explanation: Involves selecting a structure that shows an unexpected change in situation.
    """,

    8: """[もんだい8: Idiomatic and Simile Expressions]
        These questions ask you to choose between similar expressions that describe resemblance or appearance.
        Examples:
        1. その　指輪 (ゆびわ) は　星 (ほし) の　(   )　ひかっていた.
            1. みたい       2. らしく       3. ほどに       4. ように
            Answer: 4
        Explanation: Focuses on distinguishing between idiomatic expressions like “みたい”, “らしく”, and “ように”.
    """
    }

    interference_rules = {
        1: f"For Connector and Conjunction Selection related to {knowledge_point}, the interference items should be based on the correct grammatical patterns with only minor distortions. For example, if the correct connector is '〜から' (as in 'もらったから'), an interference might slightly alter the form (e.g., 'もらったかー'). Ensure the distractors are very close in form but grammatically off.",
        2: f"For Particle and Case Marker Selection concerning {knowledge_point}, the interference items should include common errors in using particles. For instance, if the correct usage is 'に' (as in '学校に行きます'), an interference might mistakenly use 'を' ('学校を行きます'). The distractors should be plausible yet clearly incorrect upon careful grammatical review.",
        3: f"For Verb and Adjective Conjugation/Transformation involving {knowledge_point}, the interference items should mirror the meaning of the correct option but feature incorrect conjugations or transformations. For example, if the correct て-form is '食べて', an interference might be '食べいて'; or if the correct reported speech form is '元気だ', an interference might be '元気で'.",
        4: f"For Modal, Volition, and Decision Expressions related to {knowledge_point}, the interference items should present errors in expressing desire, intention, or decision. For instance, if the correct form is '旅行したい', an interference might be '旅行したた'; if the decision marker is 'ことにした', an interference might use a mismatched expression that changes the intended nuance.",
        5: f"For Comparative and Degree Expressions concerning {knowledge_point}, the interference items should include subtle errors in the use of comparatives or degree markers. For example, if the correct comparative structure uses 'ほど', an interference might use a similar but inappropriate marker (like 'から' or 'なら'). The distractors should be close enough to tempt an error but grammatically unsound.",
        6: f"For Special Constructions (Dialogue/Politeness/Requests) involving {knowledge_point}, the interference items should incorporate common mistakes in honorific language, incorrect politeness levels, or dialogue-specific transformations. For example, if the correct polite request is 'してくださいませんか', an interference might simplify it to 'してください' or overcomplicate it with an unneeded form.",
        7: f"For Potential, Resultative, and Speculative Expressions related to {knowledge_point}, the interference items should distort the intended meaning by mixing up expressions of ability, result, or speculation. For instance, if the correct option is 'することになった', an interference might be 'することにした'; similarly, if the correct potential form is '入れない', an interference might be '入らない'.",
        8: f"For Idiomatic and Simile Expressions regarding {knowledge_point}, the interference items should offer choices that are similar in sound or appearance but idiomatically incorrect. For example, if the correct expression is 'ように', an interference might be 'みたいに' when the context demands a nuance that 'ように' provides. The distractors should be semantically close yet grammatically mismatched."
    }

    
    difficulty_levels = {
        1: "Basic difficulty: simple sentence structure, clear context, and direct application of grammar points.",
        2: "Intermediate difficulty: slightly complex sentence structure, may contain multiple grammatical elements, and requires certain analytical skills.",
        3: "Advanced difficulty: complex sentence structure, the context has certain implicit information, and requires a deep understanding of grammar points and context."
    }
    
    validation_prompt = """
        Generated questions must undergo the following quality validations:
        1. Grammatical accuracy: The grammar of the questions and options must be correct and meet the standards of Japanese N4/N5 level.
        2. Distractor validity: Distractors should be confusing, but can be eliminated with correct grammatical knowledge.
        3. Reasonable difficulty: Design questions according to different difficulty levels to ensure a reasonable distribution of difficulty.
        4. Contextual naturalness: The context of the sentence should be consistent with the actual situation in daily communication or exams.
        5. Context authenticity: Sentence scenarios must meet the following requirements:
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
        f"Question Format {question_format}: {format_descriptions.get(question_format, '')}\n\n"
        f"Interference Rules: {interference_rules.get(question_format, '')}\n\n"
        f"Generate questions with the following difficulty breakdown: "
        f"Basic ({base_count} questions), Intermediate ({intermediate_count} questions), Advanced ({advanced_count} questions):\n\n"
        f"{difficulty_levels[1]}\n{difficulty_levels[2]}\n{difficulty_levels[3]}\n\n"
        f"{validation_prompt}\n\n"
        f"Instructions:\n"
        f"1. Each question must start with a header: 'Qx: もんだい{question_format}'.\n"
        f"2. Provide exactly 4 options (1-4) in one line.\n"
        f"3. End each question with 'Answer: x'.\n"
        f"4. No extra text or formatting!\n\n"
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

# def grammar_points_revise(num_list, grammar_list, output_dir, filepath):
#     """针对每个知识点生成题目并修订"""
#     filename = os.path.splitext(os.path.basename(filepath))[0]
#     for knowledge_point, question_number in zip(grammar_list, num_list):
#         print(f'Processing {knowledge_point} with number {question_number}...')
#         output_doc = Document()
#         for question_format in range(1, 9):
#             print(f"Generating questions for format {question_format}...")
#             questions_text = generate_grammar_questions(knowledge_point, question_format, 5)
#             sentences, _ = split_sentences(questions_text, 1)
#             output_doc.add_paragraph(f"### Format {question_format}")
#             for sentence in sentences:
#                 sentence = sentence.replace("**Answers:**", "**Answers**")
#                 output_doc.add_paragraph(sentence)
#         output_path = os.path.join(output_dir, f"{filename}_{question_number}_{knowledge_point}.docx")
#         output_doc.save(output_path)
#         process_and_revise_document(output_path)

# ---------------------------
# 第二部分：题目检查与修订相关函数
# ---------------------------
def api_call_with_retry(chain, input_data, max_retries=3, delay=1):
    retries = 0
    while retries < max_retries:
        try:
            return chain.invoke({"input_data": input_data})
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
            If at least one question has multiple valid correct answers, respond with 'True'. Otherwise, respond with 'False'.\n
            Your output must be exactly 'True' or 'False', nothing else."""
            ),
        ]
    )
    # chain = LLMChain(llm=llm, prompt=prompt)
    input_data = {'input_data': text}
    chain = prompt | llm
    result = api_call_with_retry(chain, input_data)
    if result is not None:
        result = result.content.lower()
        return result == "true"
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
            "If there is at least one issue in the stems, respond with 'True'. Otherwise, respond with 'False'.\n"
            "Your output must be exactly 'True' or 'False', nothing else."
            )
        ]
    )
    chain = prompt | llm
    input_data = {'input_data': text}
    result = api_call_with_retry(chain, input_data)
    if result is not None:
        result = result.content.lower()
        return result == "true"
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
        if has_multiple_correct_answers(revised_text, llm):
            errors.append("Multiple correct answers")
        if has_duplicate_questions(revised_text):
            errors.append("Duplicate questions")
        if has_stem_errors(revised_text, llm):
            errors.append("Stem errors")
        if has_duplicate_options(revised_text):
            errors.append("Duplicate options")
        return errors
    except Exception as e:
        print(f"Error in check_for_error: {e}")
        return ["Unexpected error in check_for_error"]

def question_revise_simple(rows, filename, output_dir, revised_newpaper_folder, max_iterations=5, model='qwen3-235b-a22b-thinking-2507', temperature=0.6):
    """
    对新生成的日语练习题进行多轮修订和检查，确保题目质量。
    """
    llm = ChatOpenAI(temperature=temperature, model=model, api_key=os.getenv("DASHSCOPE_API_KEY"),base_url="https://dashscope.aliyuncs.com/compatible-mode/v1", max_tokens=2048, max_retries=5, extra_body={"thinking_budget": 1000})
    prompt_revise = ChatPromptTemplate.from_messages(
        [
            ("human", 
            '''
            Here are the new generated Japanese practice questions: {input_data}
            You are an experienced Japanese N4/N5 examiner tasked with reviewing and ensuring that all multiple-choice test questions meet the following criteria:

            1. No duplicate questions: Ensure that all questions are unique. If a question is repeated or too similar to another, please revise it to create a new question with a distinct structure or context. Provide specific suggestions on how to modify repeated questions.

            2. No duplicate options: All options within a question should be unique, contextually meaningful, and grammatically correct. Avoid options that are too similar to each other. If necessary, suggest how to modify similar options to increase their clarity.

            3. No duplicate correct answers: Ensure that only one answer is correct. If two options could be correct, modify the question or options to clarify the correct choice. Provide specific suggestions on how to make the answer clear and unambiguous.

            4. Grammatical correctness: The title and stem of each question must be grammatically correct. Review for unnatural sentence structures and revise them to ensure fluency and correctness. If you detect any grammatical errors, please explain how to fix them.

            5. Relevance of options: Ensure that the stem clearly indicates what cannot be chosen. One option should be inappropriate or clearly wrong in context, while all other options are suitable. Avoid culturally biased content. Suggest how to improve the incorrect options by reflecting common mistakes learners make.

            6. Pronunciation and Word Usage: If the question involves pronunciation, katakana, or hiragana forms, the Japanese word should be enclosed in brackets for clarity. For hiragana or katakana conversion questions, ensure that the word is written in the correct form, and the correct answer is not shown in the question stem. Also, check for spelling inconsistencies.

            7. General guidance: Eliminate any ambiguity, revise unclear options, and avoid subjective or culturally biased phrasing. Ensure all questions are at an appropriate difficulty level for the target JLPT level (N4/N5). Avoid complex words or structures outside the typical N4/N5 range.

            8. Output Format: Each question must keep the original format:
            - Each question must start with `Qx` (e.g., `Q1`, `Q2`...).
            - Each question must have exactly 4 options (`1` to `4`).
            - Each question must have an `Answer: x` at the end of it.
            - Do not include any other comments.
            '''
            ),
        ]
    )
    chain = prompt_revise | llm
    input_data = {'input_data': rows}
    
    for iteration in range(max_iterations):
        revised_result = api_call_with_retry(chain, input_data)
        revised_result = revised_result.content if revised_result else None
        if revised_result is None:
            print("Failed to get revised result. Stopping iteration.")
            break
        errors = check_for_error(revised_result, llm)
        if not errors:
            print(f"No issues found after {iteration + 1} iterations.")
            break
        print(f"Iteration {iteration + 1}: Detected errors - {', '.join(errors)}")
        input_data['new_paper'] = revised_result
        intermediate_path = os.path.join(output_dir, f"{filename}_iteration_{iteration + 1}.docx")
        output_doc = Document()
        sentences = split_into_sentences(revised_result)
        for sentence in sentences:
            output_doc.add_paragraph(sentence)
        output_doc.save(intermediate_path)
        log_path = os.path.join(output_dir, f"{filename}_error_log.txt")
        with open(log_path, 'a', encoding='utf-8') as log_file:
            log_file.write(f"Iteration {iteration + 1} Errors: {', '.join(errors)}\n")
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
        # for question_format in range(1, 9):
        for question_format in range(1, 2):
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
    test_grammar_original = "E:\\Microsoft VS Code\\Work Place\\JAP_GPT\\JAP_GPT\\2025_new_db\\new_questions\\test_grammar.docx"
    revised_output_grammar = "E:\\Microsoft VS Code\\Work Place\\JAP_GPT\\JAP_GPT\\2025_new_db\\new_questions\\grammar_test\\revised_grammar"
    grammar_output = "E:\\Microsoft VS Code\\Work Place\\JAP_GPT\\JAP_GPT\\2025_new_db\\new_questions\\grammar_test\\grammar_original"

    grammar_num = extract_numbered_content(test_grammar_original, 1, 4)[0]
    grammar_test = extract_numbered_content(test_grammar_original, 1, 4)[1]

    print(grammar_test)

    generate_check_store_pipeline(grammar_test, grammar_num, grammar_output, revised_output_grammar, test_grammar_original)
    process_word_to_excel(revised_output_grammar, revised_output_grammar)


if __name__ == "__main__":
    main()
