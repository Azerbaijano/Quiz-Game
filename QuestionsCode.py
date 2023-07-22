import openpyxl as open

# Make sure excel file is in same folder.
file_path = r'C:\Users\ericl\OneDrive - Onslow College\2022\11 DIT\School\Questions.xlsx'
workbook = open.load_workbook(file_path)
worksheet = workbook['Sheet1']
list = ["A", "B", "C", "D", "E"]

questionno = 1

cell = worksheet["{}{}".format(list[0], questionno + 1)]
level = cell.value
cell = worksheet["{}{}".format(list[1], questionno + 1)]
question = cell.value
cell = worksheet["{}{}".format(list[2], questionno + 1)]
options = cell.value
optionlist = options.split(',')
cell = worksheet["{}{}".format(list[3], questionno + 1)]
c_answers = cell.value
c_answerslist = c_answers.split(',')
cell = worksheet["{}{}".format(list[4], questionno + 1)]
w_answers = cell.value
w_answerslist = w_answers.split(',')



def askquestion(level, question, options, c_answers, w_answers):
    print("For ${}: {}".format(level, question))
    for i in options:
        print(i)
    valid = False
    while valid is False:
        answer = input("Enter your answer:").upper().strip()
        if answer in c_answers:
            print("You got the correct answer. You now have ${}.".format(level))
            valid = True
        elif answer in w_answers:
            print("You got the wrong answer. You won nothing.")
            print("The correct answer was {}.".format(c_answers[1]))
            valid = True
            return level
        else:
            print("Not a valid answer.")
winning = True
while winning is True:
    winning = askquestion(level, question, optionlist, c_answerslist, w_answerslist)
print("You won ${}.".format(winning))
