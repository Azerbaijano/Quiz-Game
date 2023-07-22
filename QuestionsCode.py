import openpyxl as open
import os

# Make sure excel file is in same folder.
current_directory = os.path.dirname(os.path.realpath(__file__))
file_path = os.path.join(current_directory, "Questions.xlsx")
workbook = open.load_workbook(file_path)
worksheet = workbook['Sheet1']


def askquestion(questionno):
    cell = worksheet["A{}".format(questionno + 1)]
    level = cell.value
    cell = worksheet["B{}".format(questionno + 1)]
    question = cell.value
    cell = worksheet["C{}".format(questionno + 1)]
    options = cell.value
    optionlist = options.split(',')
    cell = worksheet["D{}".format(questionno + 1)]
    c_answers = cell.value
    c_answerslist = c_answers.split(',')
    cell = worksheet["E{}".format(questionno + 1)]
    w_answers = cell.value
    w_answerslist = w_answers.split(',')
    print("For ${}: {}".format(level, question))
    for i in optionlist:
        print(i)
    valid = False
    while valid is False:
        answer = input("Enter your answer:").upper().strip()
        if answer in c_answerslist:
            print("You got the correct answer. You now have ${}.".format(level))
            valid = True
            return True
        elif answer in w_answerslist:
            print("You got the wrong answer. You won nothing.")
            print("The correct answer was {}.".format(c_answerslist[1]))    
            return False
        else:
            print("Not a valid answer.")

questionno = 1
while askquestion(questionno):
    questionno += 1
cell = worksheet["A{}".format(questionno + 1)]
amountwon = cell.value
print("You won ${}.".format(amountwon))

