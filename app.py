from os import linesep
import random
import time
from os import path
import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
import sys
from enum import Enum


class QuestionTypes(Enum):
    Addition = "+"
    Multiplication = "X"
    Division = "%"
    Subtraction = "-"


def get_result(no1: int, no2: int, question_type: QuestionTypes):
    if question_type ==QuestionTypes.Addition:
        return no1+no2
    elif question_type ==QuestionTypes.Subtraction:
        return no1-no2
    elif question_type ==QuestionTypes.Multiplication:
        return no1*no2
    elif question_type ==QuestionTypes.Division:
        return no1//no2

def get_normalize_spacing_in_nos(no1, no2):
    no1_str = str(no1)
    no2_str = str(no2)
    diff = len(no1_str)-len(no2_str)
    max_len = len(no1_str) if diff > 0 else len(no2_str)
    nos_tuple = (no1_str, " "*diff+no2_str,
                 max_len) if diff > 0 else (" "*abs(diff)+no1_str, no2_str, max_len)
    return nos_tuple

def display_question(no1: int, no2: int, question_type: QuestionTypes):
    if question_type != QuestionTypes.Division:
        tuple_no_as_string = get_normalize_spacing_in_nos(no1, no2)
        print(f"Q.{i+1})", end="\n\n")
        print(" "*20, tuple_no_as_string[0])
        print(" "*18, question_type.value+" "+tuple_no_as_string[1])
        print('_'*(20+tuple_no_as_string[2]))
    else:
        no1=str(no1)
        no2=str(no2)
        no1_length = len(no1)
        no2_length = len(no2)
        x=3
        print(" "*(no2_length+1+x)+"_"*((no1_length*2)+x))
        print(no2+" "*x+"|"+" "*x+no1)
        print("_"*(no2_length+x))

def save_report(list_of_questions:list):
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

#to handle negative nos.
def convert_string_to_numeric(string:str):
    output=int(string) if string.isdigit() else None
    if output is None:
        try:
            output=int(string)
        except:
            pass
    return output

test_taken_on = datetime.now().strftime('%d/%m/%y %H:%M:%S')
max_no = int(input("Enter the max. range :"))
no_of_questions = int(input("no of questions you want :"))
print("- select type of Questions - \n")
choice_random = input("random[y/n]")

question_types = []
if choice_random.lower() == 'n':
    if input("Add(y/n)").lower() == 'y':
        question_types.append(QuestionTypes.Addition)
    if input("Subtract(y/n)").lower() == 'y':
        question_types.append(QuestionTypes.Subtraction)
    if input("Multiply(y/n)").lower() == 'y':
        question_types.append(QuestionTypes.Multiplication)
    if input("Divide(y/n)").lower() == 'y':
        question_types.append(QuestionTypes.Division)
else:
    question_types = [operation for operation in QuestionTypes]


list_of_questions = []
total_correct_answers = 0

for i in range(no_of_questions):
    no1 = random.randrange(1, max_no)
    no2 = random.randrange(1, max_no)
    question_type = random.choice(question_types)
    display_question(no1,no2,question_type)
    start_time = time.time()
    output_entered = input()
    end_time = time.time()
    time_taken = end_time - start_time
    output_by_user = convert_string_to_numeric(output_entered)
    output_by_computer = get_result(no1, no2,question_type)
    is_answer_correct = False
    if output_by_computer == output_by_user:
        is_answer_correct = True
        total_correct_answers = total_correct_answers + 1
        print("Congratulations. you are correct. ", end="\n\n")
    else:
        print("Wrong answer submitted. Answer is ",
              output_by_computer, end="\n\n")

    list_of_questions.append({"question": f"{no1}  {question_type.value}  {no2}", "expected": output_by_computer,
                             "submitted": output_by_user, "correct": is_answer_correct, "time_taken": time_taken, "test_date": test_taken_on})


print(f"Total correct Answers : {total_correct_answers} \n")
print("Report Card for Today: \n\n")
correct = "👍"
wrong = "❌"
for item in list_of_questions:
    print(f" {item['question']}   || {item['expected']}  || {item['submitted']}  ||  { correct if item['correct'] else  wrong}  || Time Taken {item['time_taken']}")

save_report(list_of_questions)