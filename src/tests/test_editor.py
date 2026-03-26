import unittest
from excel_editor.editor import edit_excel_file
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

class TestExcelEditor(unittest.TestCase):

    def setUp(self):
        self.test_file = 'test.xlsx'
        self.test_sheet = 'Sheet1'
        self.create_test_excel_file()

    def create_test_excel_file(self):
        workbook = load_workbook(self.test_file)
        sheet = workbook.create_sheet(self.test_sheet)
        sheet['A1'] = 'Name'
        sheet['B1'] = 'Age'
        sheet['A2'] = 'Alice'
        sheet['B2'] = 30
        sheet['A3'] = 'Bob'
        sheet['B3'] = 25
        workbook.save(self.test_file)

    def test_edit_entry(self):
        edit_excel_file(self.test_file, self.test_sheet, 'Alice', 'Age', 31)
        workbook = load_workbook(self.test_file)
        sheet = workbook[self.test_sheet]
        self.assertEqual(sheet['B2'].value, 31)

    def tearDown(self):
        import os
        os.remove(self.test_file)

if __name__ == '__main__':
    unittest.main()