import telebot
from dotenv import load_dotenv
import os
from telebot import types
import openpyxl


load_dotenv(r'C:\Users\kasus\Desktop\pythonapps\calculator_work\data.env')

TELEGRAM_BOT_TOKEN = os.getenv('TELEGRAM_BOT_TOKEN')
bot = telebot.TeleBot(TELEGRAM_BOT_TOKEN)

user_data = {}


class Calculator:
    def __init__(self, file_name_read):
        self.file_name_read = file_name_read

    def parse_excel(self):
        workbook = openpyxl.load_workbook(self.file_name_read)
        sheet = workbook.active

        results = []
        for row in sheet.iter_rows(min_row=2, values_only=True): 
            country, pol_row, pod, container_row, railway, freight, unloading, service, shipping, dropoff_row = row

            if pol_row.strip().lower() == pol and dropoff_row.strip().lower() == dropoff:
                if container == container_row:
                    result = (f"Match found for 20DC: {{'country': {country}, 'pol': {pol}, 'pod': {pod}, 'container': {container}, \n"
                              f"'railway rate': {railway}, 'freight price': {freight}, 'unloading price': {unloading}, \n"
                              f"'shipping line': {shipping}, 'dropoff place': {dropoff}}}")
                    results.append(result)
                return results

@bot.message_handler(commands=['start'])
def start_bot(message):
    markup = types.InlineKeyboardMarkup()
    markup.add(
        types.InlineKeyboardButton('Тарифы', callback_data="count"), 
    )
    bot.send_message(message.chat.id, f"Привет!, {message.from_user.username} \n"
                    "Это бот для расчета тарифов!\n"
                    "Для работы с ботом для начала нужно нажать на кнопку выбора далее ввести данные в формате:\n"
                    "Порт отправления - только латиницей(например shanghai)\n"
                    "Тип контейнера - только слитно и латиницей(например - 20dc)\n"
                    "Место сдачи порожнего - только латиницей(например - moscow)\n"
                    "Команды доступные в данном боте -"
                    "/start - команда которая вызывает данное сообщение\n", reply_markup=markup)


@bot.callback_query_handler(func= lambda call: call.data == 'count')
def handle_count(call):
    user_data[call.from_user.id] = {}
    bot.send_message(call.message.chat.id, 'Введите порт отправления: ')
    bot.register_next_step_handler(call.message, get_pol)

def get_pol(message):
    user_data[message.from_user.id]['pol'] = message.text.strip().lower()
    bot.send_message(message.chat.id, 'Введите тип контейнера: ')
    bot.register_next_step_handler(message, get_container)

def get_container(message):
    user_data[message.from_user.id]['pol'] = message.text.strip().lower()
    bot.send_message(message.chat.id, 'Введите тип контейнера: ')
    bot.register_next_step_handler(message, get_container)


@bot.message_handler(commands=['generate'])
def generate(message):
    markup = types.InlineKeyboardMarkup()
    markup.add(types.InlineKeyboardButton('Сгенерировать логи', callback_data="generate_logs"))
    bot.send_message(message.chat.id, f'Для генерации логов нажмите на кнопку.', reply_markup=markup)



@bot.message_handler(commands=['count'])
def count(message):
    calc = Calculator(file_name_read='calc_draft.xlsx') 
    calc.parse_excel()
    markup = types.InlineKeyboardMarkup(row_width=2)
    markup.add(
    types.InlineKeyboardButton('Еще раз', callback_data="start"))
    bot.send_message(message.chat.id, 'test', reply_markup=markup)


bot.polling()