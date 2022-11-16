from xlsxopera import Notebook


class TestRowsAndCols:

    @staticmethod
    def _rows_and_cols_book():
        return Notebook("test.xlsx", "rows and cols")

    def test_get_data_in_list_by_rows(self):
        """
        Get first 2 values and last 2 values in sheet. Listing by rows get first rows from position A1 and B1.
        Same logic for last 2 values.
        """
        assert True if self._rows_and_cols_book().data_rows()[0:2] == [
            'COLUMN A', 'COLUMN B'
        ] and self._rows_and_cols_book().data_rows()[-2::] == [
            'John Smith', None
        ] else False

    def test_get_data_in_list_by_columns(self):
        """
        Get first 2 values and last 2 values in sheet. Listing by columns get first rows from position A1 and A2.
        Same logic for last 2 values.
        """
        assert True if self._rows_and_cols_book().data_cols()[0:2] == [
            'COLUMN A', 'value_1_column_A'
        ] and self._rows_and_cols_book().data_cols()[-2::] == [
            'value_1_column_F_center_vertically', None
        ] else False

    def test_get_rows_2_and_3_in_list(self):
        """
        Get lists row from 2 to 4 and compare.
        """
        assert self._rows_and_cols_book().list_rows(start=2, end=4) == [
            [
                'value_1_column_A',
                'value_1_column_B',
                'value_1_column_C',
                'value_1_column_D',
                'value_1_column_E',
                'value_1_column_F'
            ],
            [
                'value_2_column_A',
                'value_2_column_B',
                'value_2_column_C',
                'value_2_column_D',
                'value_2_column_E',
                'value_2_column_F'
            ],
            [
                'value_3_column_A',
                'value_3_column_B',
                'value_3_column_C',
                'value_3_column_D',
                'value_3_column_E',
                'value_3_column_F'
            ]
        ]
