import mysql.connector
def join_search():
    
    #连接数据库
    db = mysql.connector.connect(
    host="localhost",
    user="root",
    password="123",
    database="JAPGPT"
    )

    #多次查询输入
    '''
    Please input the name or SID info for student 1, or input "end" to finish the query: Lam Tzy Mang
    Please input the name or SID info for student 2, or input "end" to finish the query: Li Ching Yin Cathy
    Please input the name or SID info for student 3, or input "end" to finish the query: end
    '''
    inputs = []
    i = 1
    while True:
        query = input(f'Please input the NAME OR SID info for student {i}, or input "end" to finish the query: ')
        if query == 'end':
            break
        inputs.append(query)
        i += 1
    inputs.append('end')
    '''
        #逐行信息内连接查询
        j = 1
        while inputs[j-1] != 'end':

            info = inputs[j-1].split(',')
            Column_student = ''
            Column_question = ''
            constrain_student = ''
            constrain_question = ''

            info = [i.strip() for i in info]
    '''

    Column_student = ''
    constrain_student = ''

        #整理查询信息
    for query in inputs:
        try:
            query = int(query)
        except:
            pass
        if isinstance(query, int):
            Column_student = 'students.student_no,'
            constrain_student = 'students.student_no =' + str(query) + ' AND '
            
        else:
            Column_student = 'students.name,'
            constrain_student = 'students.name = \'' + query + '\' AND '


        inner_search_students_exam_constrain = ''

        #内连接
        if Column_student != '':
            inner_search_students_exam_constrain=f"""(SELECT students.student_id FROM students WHERE {constrain_student.rstrip(' AND ')}) = exam_results.student_id"""

        whole_constrain = ''
        for i in (constrain_student.rstrip(' AND '),  inner_search_students_exam_constrain):
            if i != '':
                whole_constrain += ' AND ' + i
        whole_constrain = whole_constrain.lstrip(' AND ')

        inner_search_query =f"""
        SELECT students.student_no, students.name, students.email, questions.question_index, questions.type, questions.level, questions.is_gpt, questions.content, questions.correct_answer, exam_results.student_answer, exam_results.is_correct FROM students, questions, exam_results
        WHERE {whole_constrain} AND students.student_id = exam_results.student_id AND questions.question_id = exam_results.question_id;
        """

        cursor = db.cursor()
        cursor.execute(inner_search_query)
        result = cursor.fetchall()
        name = ''
        print(end='\n\n\n\n\n')
        for row in result:
            if name != row[1]:
                name = row[1]
                print("student info: ", "student_no:",row[0]," ", "name:",row[1]," ", "email:",row[2],end = '\n')
            if row[10] == 1:
                print("question info: ", "question_index:",row[3]," ", "type:",row[4]," ", "level:",row[5]," ", "is_gpt:",row[6]," ",end = '\n')
                print("question content: ",row[7],end = '\n')
                print("correct answer: ",row[8], ";" ,"student answer: ",row[9], end = '\n')
                print("------------------------------------------------------")
