import unittest
from unittest.mock import MagicMock, patch
import pandas as pd
import telebot

from __main__ import send_message, analyze_and_notify, handle_start, handle_restart, handle_stop, handle_document

class TestBot(unittest.TestCase):

    def setUp(self):
        self.bot = telebot.TeleBot('7832112372:AAGnYZDYoOM4w5Z6y9hk8wM2ftmj_QPNj9s')
        self.chat_id = 123456789

    @patch('your_bot_file.bot.send_message')
    def test_send_message(self, mock_send_message):
        message = "Test message"
        send_message(self.chat_id, message)
        mock_send_message.assert_called_once_with(self.chat_id, message)

    @patch('your_bot_file.send_message')
    def test_analyze_and_notify(self, mock_send_message):
        data = {
            ('Месяц', 'Выдано'): [10, 20],
            ('Неделя', 'Выдано'): [5, 15],
            ('День', 'Выдано'): [2, 8],
            ('Месяц', 'План'): [50, 60],
            ('Неделя', 'План'): [25, 35],
            ('День', 'План'): [10, 20]
        }
        df = pd.DataFrame(data)
        df.to_excel('test_file.xlsx', header=[0, 1])

        analyze_and_notify('test_file.xlsx', self.chat_id)
        mock_send_message.assert_called_once()

    @patch('your_bot_file.bot.reply_to')
    def test_handle_start(self, mock_reply_to):
        message = MagicMock()
        handle_start(message)
        mock_reply_to.assert_called_once_with(message, 'Привет! Отправьте мне файл Excel для анализа.')

    @patch('your_bot_file.bot.reply_to')
    def test_handle_restart(self, mock_reply_to):
        message = MagicMock()
        handle_restart(message)
        mock_reply_to.assert_called_once_with(message, 'Бот перезапущен.')

    @patch('your_bot_file.bot.reply_to')
    def test_handle_stop(self, mock_reply_to):
        message = MagicMock()
        handle_stop(message)
        mock_reply_to.assert_called_once_with(message, 'Бот остановлен.')

    @patch('your_bot_file.bot.get_file')
    @patch('your_bot_file.bot.download_file')
    @patch('your_bot_file.bot.reply_to')
    @patch('your_bot_file.analyze_and_notify')
    def test_handle_document(self, mock_analyze_and_notify, mock_reply_to, mock_download_file, mock_get_file):
        message = MagicMock()
        message.document.file_id = 'file_id'
        file_info = MagicMock()
        file_info.file_path = 'file_path'
        mock_get_file.return_value = file_info
        mock_download_file.return_value = b'file_content'

        handle_document(message)
        mock_get_file.assert_called_once_with('file_id')
        mock_download_file.assert_called_once_with('file_path')
        mock_analyze_and_notify.assert_called_once_with('received_file.xlsx', message.chat.id)
        mock_reply_to.assert_called_once_with(message, 'Файл получен и обработан.')

if __name__ == '__main__':
    unittest.main()
