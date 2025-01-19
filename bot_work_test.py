import telebot
from dotenv import load_dotenv
import os
from telebot import types
import openpyxl

load_dotenv(r'C:\Users\kasus\Desktop\pythonapps\calculator_work\data.env')

TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)

user_data = {}

dropoff_places = [
    "moscow", "petersburg", "novosibirsk", "ekaterinburg", "tolyatti", "rostov",
]

pol_places = [
    "ningbo", "shanghai", "qingdao", "tianjin", "xiamen", "dalian", "shenzhen", "nansha",
]

class Calculator:
    def __init__(self, file_name_read):
        self.file_name_read = file_name_read

    def parse_excel(self, pol, container, dropoff):
        workbook = openpyxl.load_workbook(self.file_name_read, data_only=True)
        sheet = workbook.active

        results = []
        for row in sheet.iter_rows(min_row=2, max_col=20, values_only=True): 
            if len(row) != 20:
                continue

            country, pol_row, pod, railway_terminal, container_row, max_weight, usage_type, freight, railway, railway_protection, convertation, shipping_line, dropoff_row, validity_from, validity_till, total, total_sale_AB, stability, comment, currency_course = row

            pol_row = (pol_row or "").strip().lower()
            container_row = (container_row or "").strip().lower()
            dropoff_row = (dropoff_row or "").strip().lower()

            if pol_row == pol and container_row == container and dropoff_row == dropoff:
                result = (f"Найдено совпадение:\n"
                          f"Страна: {country}\n"
                          f"POL: {pol_row}\n"
                          f"POD: {pod}\n"
                          f"ЖД терминал: {railway_terminal}\n"
                          f"Контейнер: {container_row}\n"
                          f"Пользование ктк: {usage_type}\n"
                          f"Максимальный вес: {max_weight}\n"
                          f"Ставка фрахт: {freight} USD\n"
                          f"Ставка ЖД: {railway} RUB\n"
                          f"Ставка охрана ЖД: {railway_protection} RUB\n"
                          f"Конвертация: {convertation} %\n"
                          f"Линия: {shipping_line}\n"
                          f"Место сдачи порожнего: {dropoff_row}\n"
                          f"Валидность от: {validity_from}\n"
                          f"Валидность до: {validity_till}\n"
                          f"Общая ставка: {total} USD\n"
                          f"Продажа A/B: {total_sale_AB} USD\n"
                          f"Стабильность: {stability}\n"
                          f"Комментарий: {comment}\n"
                          f"Курс валют(RUB/USD): {currency_course}\n")
                results.append(result)

        return results

@bot.message_handler(commands=['start'])
def start_bot(message):
    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton('Тарифы', callback_data="count"),
        types.InlineKeyboardButton('Порты выхода', callback_data="show_pol"),
        types.InlineKeyboardButton('Точки назначения', callback_data="show_drop")
    )
    bot.send_message(message.chat.id, f"Привет!, {message.from_user.username} \n"
                    "Это бот для расчета тарифов!\n"
                    "Для работы с ботом для начала нужно нажать на кнопку выбора далее ввести данные в формате:\n"
                    "Порт отправления - только латиницей(например shanghai)\n"
                    "Тип контейнера - только слитно и латиницей(например - 20dc)\n"
                    "Место сдачи порожнего - только латиницей(например - moscow)\n"
                    "\n"
                    "Что делают кнопки: \n"
                     "Тарифы: начать ввод данных для расчета\n"
                     "Порты выхода: Список доступных портов отправления\n"
                     "Точки назначения: Список доступных мест сдачи порожнего\n"
                    "\n"
                     "Команды, доступные в боте:\n"
                     "/start - Перезапуск бота\n"
                     "/show_pol - Показать список портов отправления\n"
                     "/show_drop - Показать список мест сдачи порожнего", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'count')
def handle_count(call):
    user_data[call.from_user.id] = {}
    bot.send_message(call.message.chat.id, 'Введите порт отправления: ')
    bot.register_next_step_handler(call.message, get_pol)

def get_pol(message):
    user_data[message.from_user.id]['pol'] = message.text.strip().lower()
    bot.send_message(message.chat.id, 'Введите тип контейнера: ')
    bot.register_next_step_handler(message, get_container)

def get_container(message):
    user_data[message.from_user.id]['container'] = message.text.strip().lower()
    bot.send_message(message.chat.id, 'Введите место сдачи порожнего контейнера: ')
    bot.register_next_step_handler(message, get_dropoff)

def get_dropoff(message):
    user_data[message.from_user.id]['dropoff'] = message.text.strip().lower()
    pol = user_data[message.from_user.id]['pol']
    container = user_data[message.from_user.id]['container']
    dropoff = user_data[message.from_user.id]['dropoff']

    calc = Calculator(file_name_read='new_calc_draft1.xlsx')
    results = calc.parse_excel(pol, container, dropoff)

    if results:
        for result in results:
            bot.send_message(message.chat.id, result)
    else:
        bot.send_message(message.chat.id, 'Совпадений не найдено((')

    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton('Начать заново', callback_data="start"),
        types.InlineKeyboardButton('Посчитать еще раз', callback_data="count")
    )
    bot.send_message(message.chat.id, 'Что бы продолжить, выберите действие:', reply_markup=markup)

    user_data.pop(message.from_user.id, None)

@bot.callback_query_handler(func=lambda call: call.data == 'show_pol')
def show_pol_callback(call):
    places_list = "\n".join([f"- {place}" for place in pol_places])

    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton('Начать заново', callback_data="start"),
        types.InlineKeyboardButton('Порты выхода', callback_data="show_pol"),
        types.InlineKeyboardButton('Точки назначения', callback_data="show_drop")
    )

    bot.send_message(call.message.chat.id, f"Список портов отправления:\n{places_list}", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'show_drop')
def show_drop_callback(call):
    places_list = "\n".join([f"- {place}" for place in dropoff_places])

    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton('Начать заново', callback_data="start"),
        types.InlineKeyboardButton('Порты выхода', callback_data="show_pol"),
        types.InlineKeyboardButton('Точки назначения', callback_data="show_drop")
    )

    bot.send_message(call.message.chat.id, f"Список точек назначения:\n{places_list}", reply_markup=markup)

@bot.callback_query_handler(func=lambda call: call.data == 'start')
def restart_bot(call):
    start_bot(call.message)

if __name__ == '__main__':
    bot.polling()
