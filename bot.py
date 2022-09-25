import telebot
from telebot import types
import pylightxl as xl

db = xl.readxl(fn='bot/1.xlsx')
sheet = db.ws(ws='Лист1')
col2num = xl.pylightxl.utility_columnletter2num
num2col = xl.pylightxl.utility_num2columnletters

bot = telebot.TeleBot("")



days = ("Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота")
lessons = ("9.00 - 10.20", "10.30 - 11.50", "12.20 - 13.40", "13.50 - 15.10", "15.20 - 16.40", "16.50 - 18.10", "18.20 - 19.40")
group_list = []

@bot.message_handler(commands=["start"])
def make_keybord(message, res = False):
    group_keyboard = types.ReplyKeyboardMarkup(one_time_keyboard = True, resize_keyboard = True, row_width=3)
    for i in range(col2num("D"),col2num("DE"),2):
        group_name = sheet.address(address=num2col(i)+'16')
        group_list.append(group_name)
    if message.text in group_list:
        schedule_message(message)
    else:
        for i in range(0,len(group_list),3):
            print(i)
            key0 = types.KeyboardButton(text=group_list[i])
            try:
                key1 = types.KeyboardButton(text=group_list[i+1])
            except:
                key1 = types.KeyboardButton("Руки из жопы")
            try:
                key2 = types.KeyboardButton(text=group_list[i+2])
            except:
                key2 = types.KeyboardButton("Для вида")
            group_keyboard.add(key0, key1, key2)

    bot.send_message(message.from_user.id, "Выберите группу", reply_markup=group_keyboard)
    

@bot.message_handler(content_types=["text"])
def get_text_messages(message):
    if message.text == "Привет":
        bot.send_message(message.from_user.id,"Пока :)")
    elif message.text == "/help":
        bot.send_message(message.from_user.id,"Напиши /start и выбери группу")
    elif message.text in group_list:
        schedule_message(message)
    else:
        make_keybord(message)

def schedule_message(message):
    for i in range(0,6):
        output = ""
        start_parse = str(21+i*7)
        end_parse = str(27+i*7)
        column = num2col(group_list.index(message.text)*2+4)

        cell = sheet.range(address=column+start_parse+':'+column+end_parse)
        print(cell)
        for (j,_) in enumerate(cell):
            if cell[j][0] != '':
                cab = sheet.address(num2col(col2num(column)+1)+ str(int(start_parse)+j)) or "ну тут хз конечно"
                output += lessons[j] + " | " + ' '.join(cell[j][0].split()) + " - " + str(cab) + '\n' #Clear space
        if output != "":
            bot.send_message(message.from_user.id, "*"+days[i]+"\n"+"*" + output, parse_mode= 'Markdown')


bot.polling(none_stop=True, interval=0)
