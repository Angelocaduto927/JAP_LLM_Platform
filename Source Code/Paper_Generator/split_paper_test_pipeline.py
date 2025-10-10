import os
import re
import split_excel_paper
import combine_excel_paper
import jap_question_generator_v3_3_qwen_based
import combine_excel_paper
import jap_question_generator_v3_2_qwen_based
import jap_question_generator_v3_4_qwen_based


def main():
    #切分试卷
    split_number = split_excel_paper.main()

    # 3_4：在控制台选择 2（Voting machine revision）
    choice = jap_question_generator_v3_4_qwen_based.main()

    # 合并（支持 3_3 和 3_4 的 per-split 产物）
    combine_excel_paper.main(split_number=split_number, model_index_input=choice)

    # 对比（v3_2 保持不变，找 base_revised.xlsx）
    jap_question_generator_v3_2_qwen_based.main(model_index_input = choice)
    
if __name__ == '__main__':
    main()