from pathlib import Path
    
test_grammar_path = Path("docs/test_grammar.docx")
revised_output_path = Path("docs/Generated_paper/revised_grammar_questions")
grammar_output_path = Path("docs/Generated_paper/original_grammar_questions")

test_grammar_format = "^\s*(\d+)\.\s*(?:No.)\s*(?:\d+)\s*([^【\n]*)\s*(?:【([^】]*)】)?\s*(.*)?$" #example 1. No. 01　繰り返す
start_number = 1
end_number = 8

num_question_each_knowledge = 15
num_types = 5
num_questions_each_type = int(num_question_each_knowledge // num_types)  #make sure total question is a multiply of num_types

interference_rules = {
        1: """
            Please place the target compound verb knowledge point within the four options, requiring the examinee to select the correct option based on the information in the question content.
            And the official answer you provided must be the option that uses the compound verb correctly in grammar and scenario.
            The question stem must always contain a pair of parentheses ( ) to indicate the blank where the option should be filled in and it must be empty in that parentheses (don't fill in the answer).
            Example:
            Q1: もんだい1 
            火災現場を悲しそうに (      ) 住民たちの様子が忘れられない。
            1. 見詰める                      2. 見付ける                       3. 見掛ける                       4. 見合わせる
            Answer: 1
            """,
        2: """
            Please generate a multiple-choice question about the correct usage of a given Japanese compound verb, following these strict rules: \n
            1. The question stem must be written entirely in Traditional Chinese characters. \n
            2. Do NOT use any Japanese words, particles, or grammar—especially DO NOT use “の”. Instead of “問題裏の複合動詞の正確用法”, write “問題裏的複合動詞的正確用法”. Always use the Chinese possessive particle “的”, never the Japanese “の”. \n
            3. The question stem must be exactly: “就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。” \n
            4. On the next line, provide the target compound verb in square brackets, e.g., 【繰り返す (くりかえす)】. \n
            5. Options 1–3 must be complete sentences using that compound verb. \n
            6. Option 4 must be "All of the above are correct." \n
            7. Among option 1-3, exactly one option must contain a correct usage (grammatically and contextually inappropriate). The other two must be incorrect and unnatural. Thus the answer must not be option 4.\n
            8. The incorrect options should clearly violate standard usage patterns of the verb.\n
            Critical: If you use “の” anywhere in the question stem, it is a very serious error and the question is invalid. Use only Traditional Chinese, including “的” for possession. \n
            Example:\n
            就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。\n
            【繰り返す (くりかえす)】\n
            1. 失敗を繰り返してはならないんだ！\n
            2. 準備が予定より遅れているため１回目の会議を繰り返す。\n
            3. 最近、新商品の発売日を繰り返す会社が多いようだ。\n
            4. All of the above are correct\n
            Answer: 1
            """,
        3: """
            Please generate a multiple-choice question about the correct usage of a given Japanese compound verb, following these strict rules: \n
            1. The question stem must be written entirely in Traditional Chinese characters. \n
            2. Do NOT use any Japanese words, particles, or grammar—especially DO NOT use “の”. Instead of “問題裏の複合動詞の正確用法”, write “問題裏的複合動詞的正確用法”. Always use the Chinese possessive particle “的”, never the Japanese “の”. \n
            3. The question stem must be exactly: “就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。” \n
            4. On the next line, provide the target compound verb in square brackets, e.g., 【立ち上がる (たちあがる)】. \n
            5. Options 1–3 must be complete sentences using that compound verb. \n
            6. Option 4 must be "All of the above are correct." \n
            7. Options 1-3 must be designed as correct usages (grammatically and contextually inappropriate). Such that the answer must be option 4. \n
            Critical: If you use “の” anywhere in the question stem, it is a very serious error and the question is invalid. Use only Traditional Chinese, including “的” for possession. \n
            Example:\n
            就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。\n
            【立ち上がる (たちあがる)】\n
            1. 彼らは敗戦から立ち上がった。\n
            2. 彼女たちは市民運動に立ち上がった。\n
            3. 新しいプロジェクトが立ち上がった。\n
            4. All of the above are correct\n
            Answer: 4
            """,
        4: """
            Please generate a multiple-choice question about the correct usage of a given Japanese compound verb, following these strict rules: \n
            1. The question stem must be written entirely in Traditional Chinese characters. \n
            2. Do NOT use any Japanese words, particles, or grammar—especially DO NOT use “の”. 
                - “問題裏の複合動詞の正確用法” WRONG!
                - “問題裏的複合動詞的正確用法” CORRECT!. \n
            3. The question stem must match this expression: ^就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。$ \n
            4. On the next line, provide the target compound verb in square brackets, e.g., 【受け取る (うけとる)】. \n
            5. Options 1–4 must be complete sentences using that compound verb. \n
            6. Exactly one option must contain a correct usage (grammatically and contextually inappropriate). The other three must be incorrect and unnatural. \n
            7. The incorrect options should clearly violate standard usage patterns of the verb. \n
            Critical: If you use “の” anywhere in the question stem, it is a serious error. Use only Traditional Chinese, including “的” for possession. \n
            Example:\n
            就以下四個選項，那個是表示問題裏的複合動詞的正確用法？請選擇一個最適當的答案。\n
            【受け取る (うけとる)】\n
            1. あの会社は社長に向けられた批判を謙虚に受け取らず、結局社員全員に辞められてしまった。\n
            2. 専門家からの意見を全面的に受け取り、新商品の性能を改善する。\n
            3. 先日送っていただいた報告書を受け取りました。内容を確認し、次のミーティングで議論させていただきます。\n
            4. 彼はその技を師匠から受け取った。\n
            Answer: 3
            """,
        5: """
            Please generate a multiple-choice question about the correct usage of a given Japanese compound verb, following these strict rules: \n
            1. The question stem must be written entirely in Traditional Chinese characters. \n
            2. Do NOT use any Japanese words, particles, or grammar—especially DO NOT use “の”. Instead of “問題裏の複合動詞の正確用法”, write “問題裏的複合動詞的正確用法”. Always use the Chinese possessive particle “的”, never the Japanese “の”. \n
            3. The question stem must be exactly: “就以下四個選項，那個是表示問題裏的複合動詞的錯誤用法？請選擇一個最適當的答案。” \n
            4. On the next line, provide the target compound verb in square brackets, e.g., 【引き取る】. \n
            5. Options 1–4 must be complete sentences using that compound verb. \n
            6. Exactly one option must contain an incorrect usage (grammatically or contextually inappropriate). The other three must be correct and natural. \n
            7. The incorrect option should clearly violate standard usage patterns of the verb. \n
            Critical: If you use “の” anywhere in the question stem, it is a serious error. Use only Traditional Chinese, including “的” for possession. \n
            Example: \n
            就以下四個選項，那個是表示問題裏的複合動詞的錯誤用法？請選擇一個最適當的答案。 \n
            【引き取る (ひきとる)】\n
            1. 店員：「お客様のご要望にはお応えできません。どうぞお引き取り下さい。」\n
            2. わたしの苦手な仕事を同僚の田中さんが引き取ってくれました。\n
            3. 不妊が発覚したのを機に、孤児院から孤児を引き取って育てる事を決めました。\n
            4. 昨日は先輩から電子メールで重要なメッセージを引き取った。\n
            Answer: 4
            """
    }  # need to match the number of types

    
difficulty_levels = {
    1: "Basic difficulty: simple sentence structure, clear context, and direct application of grammar points.",
    2: "Intermediate difficulty: slightly complex sentence structure, may contain multiple grammatical elements, and requires certain analytical skills.",
    3: "Advanced difficulty: complex sentence structure, the context has certain implicit information, and requires a deep understanding of grammar points and context."
}
    
validation_prompt = """
    Generated questions must undergo the following quality validations:
    1. Grammatical accuracy: The grammar of the stems and options must be correct, natural and meet the standards of Japanese N4/N5 level.
    2. Distractor validity: Distractors should be plausible enough to confuse learners (excluding cultural nuances or non-standard colloquial usages).
    3. Uniqueness: There must be only one option that satisfies the question requirement (i.e., only one correct answer when question ask which of the following is correct, or only one incorrect answer when question ask which of the following is incorrect). And the answer provided must be that unique option.
    4. Context authenticity: Sentence scenarios must meet the following requirements:
        - Context authenticity: Sentence scenarios must be sufficiently complete and detailed so that learners can perform thorough logical analysis. Only with enough contextual information can distractors be eliminated reliably based on standard grammatical rules, rather than guesswork
        - Daily conversations (such as chatting with friends, shopping)
        - Common exam scenarios (such as email writing, schedule planning)
        - Avoid artificial contexts (such as science fiction or professional fields)
    5. Unarguable answers: The correct answer must be clearly supported by the context and grammar rules, avoiding ambiguity or multiple interpretations. For the not suitable options, they must be clearly invalid based on grammar rules or context.
"""
    
base_count = max(1, int(num_questions_each_type * 0.3))
intermediate_count = max(1, int(num_questions_each_type * 0.4))
advanced_count = num_questions_each_type - base_count - intermediate_count
if advanced_count < 0:
    base_count -= 1
    intermediate_count -= 1
    advanced_count = num_questions_each_type - base_count - intermediate_count
    
prompt_config = (
    "You are an experienced Japanese teacher. Create exactly {num_questions_each_type} questions " +
    "for the grammar point: **{knowledge_point}**.\n\n" +
    "Question Format {i} and its requirements: {interference_rules}\n\n" +
    "Generate questions with the following difficulty breakdown: " +
    "Basic ({base_count} questions), Intermediate ({intermediate_count} questions), Advanced ({advanced_count} questions):\n\n" +
    "{difficulty_levels[1]}\n{difficulty_levels[2]}\n{difficulty_levels[3]}\n\n" +
    "{validation_prompt}\n\n" +
    "Instructions:\n" +
    "1. Each question must start with a header: 'Qx: もんだい{i}'.\n" +
    "2. Provide exactly 4 options (1-4) in one line.\n" +
    "3. End each question with 'Answer: x'.\n" +
    #"4. The question stem must always contain a pair of parentheses ( ) to indicate the blank where the option should be filled in and it must be empty in that parentheses (don't fill in the answer).\n" +
    "5. No extra text or formatting!\n\n" +
    "Example Format:\n" +
    "Q1: もんだい{i}\n" +
    "[Japanese Question Stem]\n" +
    "1. Option1 2. Option2 3. Option3 4. Option4\n" +
    "Answer: x\n"
)