import telebot
import json
import gspread
from gspread.cell import Cell
gc = gspread.service_account(filename='key.json')

# Open a spreadsheet by title
sh = gc.open("Rating")
wk = sh.sheet1

bot = telebot.TeleBot('token')
try:
    with open("students.json", "r") as file:
        students = json.load(file)
        i=1
        for student in students:
            print(("{}) {}\t{}\t{}\t{}\n".format(i, students[student][0], students[student][2], student, students[student][1])))
            i+=1
except:
    students={}
def create(students):
    sorted_students = sorted(students.items(), key=lambda x: x[1][2])
    return [student[1] for student in sorted_students]

def update(students):
    try:
        cells=[]
        wk.clear()
        cells.append(Cell(row=1, col=1, value='Position'))
        cells.append(Cell(row=1, col=2, value='Group'))
        cells.append(Cell(row=1, col=3, value='Name'))
        cells.append(Cell(row=1, col=4, value='Points'))
        cells.append(Cell(row=1, col=5, value='Telegram name'))
        #wk.update('A1', "Position")
        #wk.update('B1', "Group")
       #wk.update('C1', "Name")
        #wk.update('D1', "Points")
        #wk.update('E1', "Telegram name")
        wk.format('A1:E1', {'textFormat': {'bold': True}})
        wk.format('A2:E17', {'backgroundColor': {'red': 0,'green': 1,'blue': 0}})
        for i in range(0, 31):
            if i<len(students):
                cells.append(Cell(i+2,1, value=i+1))
                cells.append(Cell(i+2,2, students[i][0]))
                cells.append(Cell(i+2,3, students[i][1]))
                cells.append(Cell(i+2,4, students[i][2]))
                cells.append(Cell(i+2,5, students[i][4]))
            else:
                cells.append(Cell(i+2,1, value=""))
                cells.append(Cell(i+2,2, value=""))
                cells.append(Cell(i+2,3, value=""))
                cells.append(Cell(i+2,4, value=""))
                cells.append(Cell(i+2,5, value=""))
        wk.update_cells(cells)
    except Exception as e:
        print(e)

@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    text=message.text
    name = str(message.from_user.first_name)
    if message.from_user.last_name: name = name + " " + str(message.from_user.last_name)
    user_id = str(message.from_user.id)
    try:
        if user_id!="800918003":
            print(f"New message from {name}({user_id}): {text}")
            f = open("log.txt", "a")
            f.write(f"New message from {name}({user_id}): {text}\n")
            f.close
    except Exception as e:
        print(e)
    if not user_id in students:
        students[user_id] = ["", "", 0, 1, name]
        bot.send_message(user_id, "Send group name\n1 for 1-AKIT\n2 for 2-AKIT")
    else:
        if message.text.startswith("/"):
            command = message.text[1:]
            if command == "edit":
                bot.send_message(user_id, "Send group name\n1 for 1-AKIT\n2 for 2-AKIT")
                students[user_id][3]=1
            elif user_id=="800918003" and command.find("delete")!=-1:
                try:
                    del students[command[command.find(" ")+1:]]
                    with open("students.json", "w") as file:
                        json.dump(students, file)
                    bot.send_message(user_id, "Deleted")
                    update(create(students)[::-1])
                    bot.send_message(user_id, "Updated")
                except Exception as e:
                    print(e)
            elif user_id=="800918003" and command.find("list")!=-1:
                try:
                    res=""
                    i=1
                    for student in students:
                        res= res +("{}) {}\t{}\t{}\t{}\n".format(i, students[student][0], students[student][2], student, students[student][1]))
                        i+=1
                    bot.send_message(user_id, res)
                except Exception as e:
                    print(e)
            elif user_id=="800918003" and command.find("update")!=-1:
                try:
                    update(create(students)[::-1])
                    bot.send_message(user_id, "Updated")
                except Exception as e:
                    print(e)
            elif user_id=="800918003" and command.find("run")!=-1:
                try:
                    print("Code:\n\n{}\n\n--------------------------\n\nResult:\n".format(command[command.find(" ")+1:]))
                    exec(command[command.find(" ")+1:])
                except Exception as e:
                    print(e)

        elif students[user_id][3]==1:
            if text.isnumeric() and (int(text)==1 or int(text)==2):
                students[user_id][0] = text + "-AKIT"
                bot.send_message(user_id, "Good, your group has been added, now send your name")
                students[user_id][3]=2
            else:
                bot.send_message(user_id, "Bad data\nSend group name\n1 for 1-AKIT\n2 for 2-AKIT")

        elif students[user_id][3]==2:
            students[user_id][1] = text
            students[user_id][3]=3
            bot.send_message(user_id, "Good, your name has been added, now send your points")
        elif students[user_id][3]==3:
            if text.isnumeric():
                students[user_id][2] = int(text)
                bot.send_message(user_id, "Good, your points has been added")
                students[user_id][3]=4
            else:
                bot.send_message(user_id, "Bad data\nsend your points(number)")
    if students[user_id][3]==4 and user_id!="800918003":
        bot.send_message(user_id, "Your data has been added send /edit to edit it\nLink to rating: https://docs.google.com/spreadsheets/d/1p53_Pb_9qb_TUOZD7L0W-G29INtR8w32uVFAD1qDOwc/edit#gid=0")
        update(create(students)[::-1])

    with open("students.json", "w") as file:
                json.dump(students, file)
bot.polling(none_stop=True, interval=0)
