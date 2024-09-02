import openpyxl
import telebot as tb
from telebot import types

def schedule_read(week, day):
    wb = openpyxl.load_workbook('Расписание_При-123.xlsx')
    ws_Denominator = wb['Знаменатель']
    ws_Numerator = wb['Числитель']

    schedule = []
    time_list = []

    worksheet = ws_Denominator if week == 'Знаменатель' else ws_Numerator

    for cell in worksheet[day][1:]:
        schedule.append(cell.value)
    for cell in worksheet['A'][1:]:
        time_list.append(cell.value)

    return schedule, time_list

def tg_bot():
    bot = tb.TeleBot('7394583406:AAF213ES0M2gh28OwNf_9Ff4IkWp1fMDTRg')

    @bot.message_handler(commands=['start'])
    def start(message):
        markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
        itembtn1 = types.KeyboardButton('Расписание')
        markup.add(itembtn1)
        bot.send_message(message.chat.id, "Что вы хотите узнать?:", reply_markup=markup)
        bot.register_next_step_handler(message, check_week)

    @bot.message_handler(func=lambda message: True)
    def handle_all_messages(message):
        if message.text != '/start':
            bot.send_message(message.chat.id, 'Чтобы начать, введите команду /start.')
        else:
            start(message)

    def check_week(message):
        if message.text == 'Расписание':
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            itembtn1 = types.KeyboardButton('Знаменатель')
            itembtn2 = types.KeyboardButton('Числитель')
            markup.add(itembtn1, itembtn2)
            bot.send_message(message.chat.id, 'Неделя по знаменателю или числителю? :)', reply_markup=markup)
            bot.register_next_step_handler(message, check_day)
        else:
            bot.send_message(message.chat.id, 'Укажите, что хотите узнать -.-')
            bot.register_next_step_handler(message, check_week)

    def check_day(message):
        if message.text in ['Знаменатель', 'Числитель']:
            week = message.text
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            itembtn1 = types.KeyboardButton('Понедельник')
            itembtn2 = types.KeyboardButton('Вторник')
            itembtn3 = types.KeyboardButton('Среда')
            itembtn4 = types.KeyboardButton('Четверг')
            itembtn5 = types.KeyboardButton('Пятница')
            markup.add(itembtn1, itembtn2, itembtn3, itembtn4, itembtn5)
            bot.send_message(message.chat.id, 'Выберите день недели =)', reply_markup=markup)
            bot.register_next_step_handler(message, lambda msg: schedule_write(msg, week))
        else:
            bot.send_message(message.chat.id, 'Укажите знаменатель или числитель =[')
            bot.register_next_step_handler(message, check_day)

    def schedule_write(message, week):
        if message.text in ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница']:
            day_map = {
                'Понедельник': 'B',
                'Вторник': 'C',
                'Среда': 'D',
                'Четверг': 'E',
                'Пятница': 'F'
            }
            day = day_map.get(message.text)
            schedule, time = schedule_read(week, day)
            for i, place in enumerate(schedule):
                if place is None:
                    continue
                bot.send_message(message.chat.id, f'{i+1} пара ({time[i]}) - {place}')
            markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
            itembtn1 = types.KeyboardButton('Расписание')
            markup.row(itembtn1)
            bot.send_message(message.chat.id, "Еще что-то хотите узнать?:", reply_markup=markup)
            bot.register_next_step_handler(message, check_week)
        else:
            bot.send_message(message.chat.id, 'Укажите день недели >:')
            bot.register_next_step_handler(message, lambda msg: schedule_write(msg, week))

    bot.polling(none_stop=True, interval=0)

if __name__ == '__main__':
    tg_bot()