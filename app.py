from os import linesep
import random
import time
from os import path
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import sys


def get_normalize_spacing_in_nos(no1, no2):
    no1_str = str(no1)
    no2_str = str(no2)
    diff = len(no1_str)-len(no2_str)
    max_len = len(no1_str) if diff > 0 else len(no2_str)
    nos_tuple = (no1_str, " "*diff+no2_str,
                 max_len) if diff > 0 else (" "*abs(diff)+no1_str, no2_str, max_len)
    return nos_tuple


test_taken_on = datetime.now().strftime('%d/%m/%y %H:%M:%S')
max_no = int(input("Enter the max. range :"))
no_of_questions = int(input("no of questions you want :"))
list_of_questions = []
total_correct_answers = 0

for i in range(no_of_questions):
    no1 = random.randrange(1, max_no)
    no2 = random.randrange(1, max_no)
    #print(f"Q.{i+1}) solve: {no1}  X  {no2} ")
    tuple_no_as_string = get_normalize_spacing_in_nos(no1, no2)
    print(f"Q.{i+1})", end="\n\n")
    print(" "*20, tuple_no_as_string[0])
    print(" "*18, "X "+tuple_no_as_string[1])
    print('_'*(20+tuple_no_as_string[2]))
    start_time = time.time()
    output_entered = input()
    output_by_user = int(output_entered) if output_entered.isdigit() else None
    end_time = time.time()
    time_utilized = end_time - start_time
    output_by_computer = no1 * no2
    is_answer_correct = False
    if output_by_computer == output_by_user:
        is_answer_correct = True
        total_correct_answers = total_correct_answers + 1
        print("Congratulations. you are correct. ", end="\n\n")
    else:
        print("Wrong answer submitted. Answer is ",
              output_by_computer, end="\n\n")

    list_of_questions.append({"question": f"{no1}  X  {no2}", "expected": output_by_computer,
                             "submitted": output_by_user, "correct": is_answer_correct, "time_taken": time_utilized, "test_date": test_taken_on})


print(f"Total correct Answers : {total_correct_answers} \n")
print("Report Card for Today: \n\n")
correct = "👍"
wrong = "❌"
for item in list_of_questions:
    print(f" {item['question']}   -  {item['expected']}  - {item['submitted']}  -   { correct if item['correct'] else  wrong}  - Time Taken {item['time_taken']}")

df = pd.DataFrame(list_of_questions)
if path.exists("result.xlsx"):
    reader = pd.read_excel('result.xlsx', engine='openpyxl')
    writer = pd.ExcelWriter('result.xlsx', engine='openpyxl', mode='a')
    # try to open an existing workbook
    writer.book = load_workbook('result.xlsx')
    # copy existing sheets
    writer.sheets = dict((ws.title, ws) for ws in writer.book.worksheets)
    df.to_excel(writer, index=False, header=False, startrow=len(reader)+1)
    writer.close()
else:
    df.to_excel("result.xlsx", index=False)
