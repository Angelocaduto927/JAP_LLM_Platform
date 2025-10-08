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
    
    #对切分后的试卷检查修改
    # choice = jap_question_generator_v3_3_qwen_based.main()
    choice = jap_question_generator_v3_4_qwen_based.main()
    
    #合并试卷
    combine_excel_paper.main(split_number=split_number, model_index_input=choice)
    
    #比较修正结果和何老师的修正结果
    jap_question_generator_v3_2_qwen_based.main(model_index_input = choice)
    
if __name__ == '__main__':
    main()