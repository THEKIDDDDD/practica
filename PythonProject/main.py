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
        total_assigned = (df[('Месяц', 'Выдано')].sum() +
                          df[('Неделя', 'Выдано')].sum() +
                          df[('День', 'Выдано')].sum())
        total_planned = (df[('Месяц', 'План')].sum() +
                         df[('Неделя', 'План')].sum() +
                         df[('День', 'План')].sum())
        if total_planned == 0:
            raise ValueError("Некорректные данные в столбцах 'Выдано' или 'План'")

        percentage_assigned = (total_assigned / total_planned) * 100
        print(f"Процент выданных заданий: {percentage_assigned:.2f}%")
        message = f"Процент выданных домашних заданий: {percentage_assigned:.2f}%\n"
        if percentage_assigned < 70:
            message += "Внимание! Выдано менее 70% домашних заданий. Пожалуйста, обратите внимание на это.\n\nС уважением,\nВаш чат-бот"
        else:
            message += "Отлично! Выдано более 70% домашних заданий. Продолжайте в том же духе!\n\nС уважением,\nВаш чат-бот"

        send_message(chat_id, message)
    except ValueError as ve:
        print(f"Ошибка при анализе данных: {ve}")
        send_message(chat_id, f"Ошибка при анализе данных: {ve}. Пожалуйста, проверьте файл и попробуйте еще раз.")
    except Exception as e:
        print(f"Ошибка при анализе данных: {e}")
        send_message(chat_id, "Произошла ошибка при анализе данных. Пожалуйста, проверьте файл и попробуйте еще раз.")

# start
@bot.message_handler(commands=['start'])
def handle_start(message):
    bot.reply_to(message, 'Привет! Отправьте мне файл Excel для анализа.')
    print("Команда /start выполнена")
# restart
@bot.message_handler(commands=['restart'])
def handle_restart(message):
    global is_running
    is_running = True
    bot.reply_to(message, 'Бот перезапущен.')
    print("Команда /restart выполнена")
# stop
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
