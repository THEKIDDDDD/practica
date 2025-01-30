import pandas as pd
import telebot
from telebot import types
import time
from telebot import apihelper


bot = telebot.TeleBot('7832112372:AAGnYZDYoOM4w5Z6y9hk8wM2ftmj_QPNj9s')

is_running = True

def send_message(chat_id, message):
    bot.send_message(chat_id, message)

def analyze_and_notify(file_path, chat_id):
    try:
        df = pd.read_excel(file_path, header=[0, 1])
        print("Файл успешно загружен")
        print(df.head())
        df = df.fillna(0)

        message = ""
        total_assigned_all = 0
        total_planned_all = 0

        for index, row in df.iterrows():
            teacher_name = row[('ФИО преподавателя', 'Unnamed: 1_level_1')]
            total_assigned = (row[('Месяц', 'Выдано')] +
                              row[('Неделя', 'Выдано')] +
                              row[('День', 'Выдано')])
            total_planned = (row[('Месяц', 'План')] +
                             row[('Неделя', 'План')] +
                             row[('День', 'План')])
            if total_planned == 0:
                raise ValueError("Некорректные данные в столбцах 'Выдано' или 'План'")

            percentage_assigned = (total_assigned / total_planned) * 100
            print(f"Процент выданных заданий для {teacher_name}: {percentage_assigned:.2f}%")
            message += f"Процент выданных домашних заданий для {teacher_name}: {percentage_assigned:.2f}%\n"
            if percentage_assigned < 70:
                message += f"Внимание! {teacher_name} выдал(а) менее 70% домашних заданий. Пожалуйста, обратите внимание на это.\n\n"
            else:
                message += f"Отлично! {teacher_name} выдал(а) более 70% домашних заданий. Продолжайте в том же духе!\n\n"

            total_assigned_all += total_assigned
            total_planned_all += total_planned

        if total_planned_all == 0:
            raise ValueError("Некорректные данные в столбцах 'Выдано' или 'План'")

        overall_percentage_assigned = (total_assigned_all / total_planned_all) * 100
        overall_message = f"Общий процент выданных домашних заданий всех преподавателей: {overall_percentage_assigned:.2f}%\n"
        if overall_percentage_assigned < 70:
            overall_message += "Внимание! Общий процент выданных домашних заданий менее 70%. Пожалуйста, обратите внимание на это.\n\nС уважением,\nВаш чат-бот"
        else:
            overall_message += "Отлично! Общий процент выданных домашних заданий более 70%. Продолжайте в том же духе!\n\nС уважением,\nВаш чат-бот"

        send_message(chat_id, message)
        send_message(chat_id, overall_message)
    except ValueError as ve:
        print(f"Ошибка при анализе данных: {ve}")
        send_message(chat_id, f"Ошибка при анализе данных: {ve}. Пожалуйста, проверьте файл и попробуйте еще раз.")
    except Exception as e:
        print(f"Ошибка при анализе данных: {e}")
        send_message(chat_id, "Произошла ошибка при анализе данных. Пожалуйста, проверьте файл и попробуйте еще раз.")

@bot.message_handler(commands=['start'])
def handle_start(message):
    bot.reply_to(message, 'Привет! Отправьте мне файл Excel для анализа.')
    print("Команда /start выполнена")

@bot.message_handler(commands=['restart'])
def handle_restart(message):
    global is_running
    is_running = True
    bot.reply_to(message, 'Бот перезапущен.')
    print("Команда /restart выполнена")

@bot.message_handler(commands=['stop'])
def handle_stop(message):
    global is_running
    is_running = False
    bot.reply_to(message, 'Бот остановлен.')
    print("Команда /stop выполнена")

@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_info = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        with open('received_file.xlsx', 'wb') as new_file:
            new_file.write(downloaded_file)

        analyze_and_notify('received_file.xlsx', message.chat.id)
        bot.reply_to(message, 'Файл получен и обработан.')
        print("Файл получен и обработан")
    except Exception as e:
        print(f"Ошибка при обработке файла: {e}")
        bot.reply_to(message, 'Произошла ошибка при обработке файла. Пожалуйста, попробуйте еще раз.')

def main():
    global is_running
    print("Бот запущен и ожидает сообщений")
    while is_running:
        try:
            bot.polling(none_stop=True, interval=0, timeout=30)
        except telebot.apihelper.ApiException as e:
            print(f"Ошибка API: {e.result.text}")
            time.sleep(5)
        except Exception as e:
            print(f"Ошибка при запуске бота: {e}")
            time.sleep(5)

if __name__ == '__main__':
    main()
