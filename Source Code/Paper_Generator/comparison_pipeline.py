import jap_question_generator_v3_1_qwen_based as revise
import extract_bad_questions
import jap_question_generator_v3_2_qwen_based as compare

def main():
    print("Step 1: 使用大模型对试卷进行检查和修改")
    print("如果列表中没有要使用的大模型，请在jap_question_generator.v3.1(qwen based).py和jap_question_generator.v3.2(qwen based).py的main函数的模型列表添加对应模型")
    revise.main()
    
    print("Step 2: 给已批改的试卷打标签")
    extract_bad_questions.main()
    
    print("Step 3: 比较大模型修改结果和人工修改结果")
    print("选择模型前请确认该模型已经运行过检查修正程序")
    compare.main()
    
    print("流程已完成")

if __name__ == "__main__":
    main()