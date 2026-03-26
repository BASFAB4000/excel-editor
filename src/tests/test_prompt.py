import unittest
from unittest.mock import patch, MagicMock
from excel_editor.prompt import prompt_user_for_input

class TestPrompt(unittest.TestCase):

    @patch('builtins.input', side_effect=['test.xlsx', 'Sheet1', 'System1', 'Edit'])
    def test_prompt_user_for_input(self, mock_input):
        file_path, sheet_name, system, action = prompt_user_for_input()
        self.assertEqual(file_path, 'test.xlsx')
        self.assertEqual(sheet_name, 'Sheet1')
        self.assertEqual(system, 'System1')
        self.assertEqual(action, 'Edit')

    @patch('builtins.input', side_effect=['invalid_path.xlsx', 'Sheet1', 'System1', 'Add'])
    def test_prompt_user_for_input_invalid_path(self, mock_input):
        file_path, sheet_name, system, action = prompt_user_for_input()
        self.assertEqual(file_path, 'invalid_path.xlsx')
        self.assertEqual(sheet_name, 'Sheet1')
        self.assertEqual(system, 'System1')
        self.assertEqual(action, 'Add')

if __name__ == '__main__':
    unittest.main()