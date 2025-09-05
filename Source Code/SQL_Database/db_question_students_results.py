import mysql.connector

# 连接到数据库
db = mysql.connector.connect(
    host="10.20.7.5",
    user="root",
    password="123456", 
    port=3306,
    database="JAPGPT"  
)
cursor = db.cursor()

# 删除旧表（注意：正式运行后不要直接删除旧表！！！）
drop_questions_table_query = "DROP TABLE IF EXISTS questions"
drop_students_table_query = "DROP TABLE IF EXISTS students"
drop_exam_results_table_query = "DROP TABLE IF EXISTS exam_results"

select_mistake_query = """
    SELECT * FROM exam_results
    WHERE is_correct = 0
    """

select_all_query = """
    SELECT * FROM exam_results
    """

insert_questions_query = """
INSERT INTO questions (question_index, content, correct_answer, type, level, is_gpt)
VALUES (%s, %s, %s, %s, %s, %s)
"""

insert_students_query = """
INSERT INTO students (student_no, name, email)
VALUES (%s, %s, %s)
"""

insert_exam_results_query = """
INSERT INTO exam_results (question_id, student_id, student_answer, is_correct)
VALUES (%s, %s, %s, %s)
"""

try:
    # 删除表，注意删除顺序，避免外键约束冲突
    cursor.execute(drop_exam_results_table_query)  # 先删除依赖于其他表的表
    cursor.execute(drop_students_table_query)  # 再删除学生表
    cursor.execute(drop_questions_table_query)  # 最后删除题目表

    print("旧表已删除")

    # 创建 `questions` 表
    create_questions_table_query = """
    CREATE TABLE questions (
        question_id INT AUTO_INCREMENT PRIMARY KEY,
        question_index VARCHAR(255) NOT NULL UNIQUE COMMENT '用户查找题目时的索引号（题目编号）',
        content TEXT NOT NULL COMMENT '题目内容',
        correct_answer TEXT NOT NULL COMMENT '题目的正确答案',
        type TEXT NOT NULL COMMENT '题目类型，存储多个知识点',
        level ENUM('N4', 'N5') NOT NULL COMMENT '难度级别',
        is_gpt BOOLEAN NOT NULL DEFAULT 0 COMMENT '是否由AI生成'
    )
    """
    cursor.execute(create_questions_table_query)
    print("questions 表已创建")

    # 创建 `students` 表
    create_students_table_query = """
    CREATE TABLE students (
        student_id INT AUTO_INCREMENT PRIMARY KEY,
        student_no BIGINT NOT NULL UNIQUE COMMENT '学生学号',
        name VARCHAR(100) NOT NULL COMMENT '学生姓名',
        email VARCHAR(255) NOT NULL UNIQUE COMMENT '学生的电子邮箱'
    )
    """
    cursor.execute(create_students_table_query)
    print("students 表已创建")

    # 创建 `exam_results` 表
    create_exam_results_table_query = """
    CREATE TABLE exam_results (
        result_id INT AUTO_INCREMENT PRIMARY KEY,
        question_id INT NOT NULL,
        student_id INT NOT NULL,
        student_answer TEXT NOT NULL,
        is_correct BOOLEAN NOT NULL,
        FOREIGN KEY (question_id) REFERENCES questions(question_id) ON DELETE CASCADE,
        FOREIGN KEY (student_id) REFERENCES students(student_id) ON DELETE CASCADE
    )
    """
    cursor.execute(create_exam_results_table_query)
    print("exam_results 表已创建")

except mysql.connector.Error as err:
    print(f"Error: {err}")

finally:
    # 提交更改并关闭连接
    db.commit()
    cursor.close()
    db.close()
    print("数据库操作完成")