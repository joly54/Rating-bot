import telebot
import json
import gspread
from gspread.cell import Cell
gc = gspread.service_account(filename='key.json')
import datetime
from openpyxl.utils import get_column_letter
# Open a spreadsheet by title
sh = gc.open("Rating")
wk = sh.sheet1

bot = telebot.TeleBot('#')
try:
    with open("students.json", "r") as file:
        students = json.load(file)
        i=1
        for student in students:
            print(("{}) {}\t{}\t{}\t{}\n".format(i, students[student][0], students[student][2], student, students[student][1])))
            i+=1
except:
    students={}
def change_db():
    global students
    for student in students:
        if students[student][3]==1:
            students[student][3]="Waiting for group"
        elif students[student][3]==2:
            students[student][3]="Waiting for name"
        elif students[student][3]==3:
            students[student][3]="Waiting for points"
        elif students[student][3]==4:
            students[student][3]="Registration finished"
        print(students[student])
        with open("students.json", "w") as file:
                    json.dump(students, file)
def create(students):
    sorted_students = sorted(students.items(), key=lambda x: x[1][2])
    return [student[1] for student in sorted_students]

def setup():
    try:
        def generate_color_array(steps):
            colors = []
            for i in range(steps):
                green = 1 - (i / steps) * 0.6
                blue = (i / steps) * 0.1
                red = i / steps
                colors.append([red, green, blue])
            return colors
        def line_format(row, col, row2, col2, r, g, b):
            wk.format(f'{get_column_letter(col)}{row}:{get_column_letter(col2)}{row2}', {'backgroundColor': {'red': r,'green': g,'blue': b}})
        colors=generate_color_array(17)
        for i in range(1,17):
            line_format(i+1,1,i+1,5, colors[i][0], colors[i][1], colors[i][2])
    except Exception as e:
        print(e)
        bot.send_message(800918003, str(e))
setup()
def red_col(row, col, row2, col2):
    wk.format(f'{get_column_letter(col)}{row}:{get_column_letter(col2)}{row2}', {'backgroundColor': {'red': 0.85,'green': 0.85,'blue': 0.85}})
def update(students):
    try:
        cells=[]
        #wk.clear()
        cells.append(Cell(row=1, col=1, value='Position'))
        cells.append(Cell(row=1, col=2, value='Group'))
        cells.append(Cell(row=1, col=3, value='Name'))
        cells.append(Cell(row=1, col=4, value='Points'))
        cells.append(Cell(row=1, col=5, value='Telegram name'))
        wk.format('A1:E1', {'textFormat': {'bold': True}})
        wk.format('A1:A50', {'textFormat': {'bold': True}})
        wk.format('A1:E1', {'backgroundColor': {'red': 1,'green': 1,'blue': 1}})
        setup()
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
        bot.send_message(800918003, str(e))
update(create(students)[::-1])
@bot.message_handler(content_types=['text'])
def get_text_messages(message):
    text=message.text
    try:
        name = str(message.from_user.first_name)
        if message.from_user.last_name: name = name + " " + str(message.from_user.last_name)
        user_id = str(message.from_user.id)
        if not user_id in students:
            students[user_id] = ["", "", 0, "Waiting for group", name]
            bot.send_message(user_id, "Send group name\n1 for 1-AKIT\n2 for 2-AKIT")
        else:
            if message.text.startswith("/"):
                command = message.text[1:]
                if command == "edit":
                    bot.send_message(user_id, "Send group name\n1 for 1-AKIT\n2 for 2-AKIT")
                    students[user_id][3]="Waiting for group"
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
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("get")!=-1:
                    try:
                        bot.send_message(800918003, str(students[command[command.find(" ")+1:]]))
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("list")!=-1:
                    try:
                        res = ""
                        i=1
                        bot.send_message(user_id, "Preparing...")
                        for student in students:
                            res = res + ("{}) Group: {}\nName: {}\nPoints: {}\nID: <code>{}</code>\nTG name: {}\nStatus: {}\n".format(i, students[student][0], students[student][1], students[student][2], student, students[student][4],students[student][3]))
                            if str((bot.get_chat(student)).username)!="None":
                                res = res + f"Username: @{(bot.get_chat(student)).username}\n"
                            res = res + "\n"
                            i+=1
                        bot.send_message(user_id, res, parse_mode='HTML')
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif command.find("link")!=-1:
                    try:
                        bot.send_message(user_id, "Link to rating: https://docs.google.com/spreadsheets/d/1p53_Pb_9qb_TUOZD7L0W-G29INtR8w32uVFAD1qDOwc/edit#gid=0", disable_web_page_preview=True)
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("update")!=-1:
                    try:
                        update(create(students)[::-1])
                        bot.send_message(user_id, "Updated")
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("edit_user")!=-1:
                    try:
                        new_data=(command[command.find(" ")+1:]).split("; ")
                        print(new_data)
                        #del students[new_data[0]]
                        if len(new_data)==5:
                            students[str(new_data[0])]=[new_data[1], new_data[2], int(new_data[3]), 4, new_data[4]]
                        elif len(new_data)==4:
                            students[str(new_data[0])]=[new_data[1], new_data[2], int(new_data[3]), 4, students[str(new_data[0])][3]]
                        update(create(students)[::-1])
                        bot.send_message(user_id, "Updated")
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("run")!=-1:
                    try:
                        print("Code:\n\n{}\n\n--------------------------\n\nResult:\n".format(command[command.find(" ")+1:]))
                        exec(command[command.find(" ")+1:])
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id=="800918003" and command.find("logs")!=-1:
                    try:
                        f = open("log.txt","rb")
                        bot.send_document(message.chat.id,f)
                    except Exception as e:
                        print(e)
                        bot.send_message(800918003, str(e))
                elif user_id!="800918003" and (command.find("logs")!=-1 or command.find("logs")!=-1 or command.find("run")!=-1 or command.find("update")!=-1 or command.find("list")!=-1 or command.find("delete")!=-1):
                    bot.send_message(user_id, "Access denied")
            elif students[user_id][3]=="Waiting for group":
                if text.isnumeric() and (int(text)==1 or int(text)==2):
                    students[user_id][0] = text + "-AKIT"
                    bot.send_message(user_id, "Good, your group has been added, now send your name")
                    students[user_id][3]="Waiting for name"
                else:
                    bot.send_message(user_id, "Bad data\nSend group name\n1 for 1-AKIT\n2 for 2-AKIT")

            elif students[user_id][3]=="Waiting for name":
                students[user_id][1] = text
                students[user_id][3]="Waiting for points"
                bot.send_message(user_id, "Good, your name has been added, now send your points(without history and english)")
            elif students[user_id][3]=="Waiting for points":
                if text.isnumeric():
                    students[user_id][2] = int(text)
                    bot.send_message(user_id, "Good, your points has been added")
                    students[user_id][3]="Registration finished"
                    update(create(students)[::-1])
                    bot.send_message(user_id, "Link to rating: https://docs.google.com/spreadsheets/d/1p53_Pb_9qb_TUOZD7L0W-G29INtR8w32uVFAD1qDOwc/edit#gid=0", disable_web_page_preview=True)
                else:
                    bot.send_message(user_id, "Bad data\nsend your points(number)")
            elif students[user_id][3]=="Registration finished" and user_id!="800918003":
                bot.send_message(user_id, "Link to rating: https://docs.google.com/spreadsheets/d/1p53_Pb_9qb_TUOZD7L0W-G29INtR8w32uVFAD1qDOwc/edit#gid=0", disable_web_page_preview=True)
        try:
            now = datetime.datetime.now()
            # Use the datetime.strftime method to specify the format of the date and time
            date_time = now.strftime("Time %H-%M-%S Date %d-%m-%y")
            print(f"New message from\t\t{name}({user_id}):\t\t{text}\t{date_time}\t\tStatus: {students[user_id][3]}\n")
            f = open("log.txt", "a")
            f.write(f"New message from\t\t{name}({user_id}):\t\t{text}\t{date_time}\t\tStatus: {students[user_id][3]}\n")
            f.close
        except Exception as e:
            print(e)
            bot.send_message(800918003, str(e))
        with open("students.json", "w") as file:
                    json.dump(students, file)
    except Exception as e:
        print(e)
        bot.send_message(800918003, str(e))

bot.polling(none_stop=True, interval=0)
