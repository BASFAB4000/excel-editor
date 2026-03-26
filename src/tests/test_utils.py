import unittest
from excel_editor.utils import validate_file_path, check_file_access

class TestUtils(unittest.TestCase):

    def test_validate_file_path_valid(self):
        valid_path = "test.xlsx"
        self.assertTrue(validate_file_path(valid_path))

    def test_validate_file_path_invalid(self):
        invalid_path = "invalid_path/test.xlsx"
        self.assertFalse(validate_file_path(invalid_path))

    def test_check_file_access_readable(self):
        readable_path = "test.xlsx"
        self.assertTrue(check_file_access(readable_path, 'r'))

    def test_check_file_access_writable(self):
        writable_path = "test.xlsx"
        self.assertTrue(check_file_access(writable_path, 'w'))

    def test_check_file_access_invalid(self):
        invalid_path = "invalid_path/test.xlsx"
        self.assertFalse(check_file_access(invalid_path, 'r'))

if __name__ == '__main__':
    unittest.main()