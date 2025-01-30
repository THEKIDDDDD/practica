import unittest
from unittest.mock import patch, MagicMock
import pandas as pd
import os

from your_bot_file import analyze_and_notify, send_message

class TestBotFunctions(unittest.TestCase):

    @patch('your_bot_file.bot.send_message')
    def test_send_message(self, mock_send_message):
        chat_id = 123456
        message = "Test message"
        send_message(chat_id, message)
        mock_send_message.assert_called_once_with(chat_id, message)

    @patch('your_bot_file.bot.send_message')
    def test_analyze_and_notify(self, mock_send_message):
        test_data = {
            ('ФИО преподавателя', 'Unnamed: 1_level_1'): ['Teacher1', 'Teacher2'],
            ('Месяц', 'Выдано'): [10, 20],
            ('Неделя', 'Выдано'): [5, 10],
            ('День', 'Выдано'): [2, 5],
            ('Месяц', 'План'): [20, 40],
            ('Неделя', 'План'): [10, 20],
            ('День', 'План'): [5, 10]
        }
        df = pd.DataFrame(test_data)
        df.to_excel('test_file.xlsx', header=[0, 1], index=False)

        chat_id = 123456
        analyze_and_notify('test_file.xlsx', chat_id)

        self.assertEqual(mock_send_message.call_count, 2)

        os.remove('test_file.xlsx')

if __name__ == '__main__':
    unittest.main()
