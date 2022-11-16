import os
from xlsxopera import Notebook

file = "test.xlsx"


class TestFile:

    # @staticmethod
    # def _active_book():
    #     return Notebook("test.xlsx")

    @staticmethod
    def _tests_book():
        return Notebook(file, "tests")

    @staticmethod
    def _rows_and_cols_book():
        return Notebook(file, "rows and cols")

    @staticmethod
    def _math_book():
        return Notebook(file, "math")

    @staticmethod
    def test_xlsx_file():
        for file_in_dir in os.listdir():
            if file_in_dir.endswith('.xlsx'):
                assert file_in_dir == file

    def test_tests_sheet_exist(self):
        assert self._tests_book().headers() == [
            'ID', 'Status', 'Test Name', 'Test Description',
        ]

    def test_rows_and_cols_sheet_exist(self):
        assert self._rows_and_cols_book().headers() == [
            'COLUMN A', 'COLUMN B', 'COLUMN C', 'COLUMN D', 'COLUMN E', 'COLUMN F'
        ]
