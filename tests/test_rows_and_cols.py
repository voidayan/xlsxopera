from xlsxopera import Notebook


class TestRowsAndCols:

    @staticmethod
    def _rows_and_cols_book():
        return Notebook("test.xlsx", "rows and cols")

    @staticmethod
    def _data_convert_book():
        return Notebook("test.xlsx", "data convert")

    def test_get_cell_coordinates_by_value(self):
        assert self._rows_and_cols_book().value_coordinates(
            "value_1_column_F"
        ) == (2, 6)
        assert self._rows_and_cols_book().value_coordinates(
            "value_3_column_A"
        ) == (4, 1)

    def test_get_cell_position_by_value(self):
        assert self._rows_and_cols_book().value_position(
            "value_3_column_C"
        ) == "C4"
        assert self._rows_and_cols_book().value_position(
            "value_2_column_E"
        ) == "E3"

    def test_find_cell_by_value(self):
        assert True if self._rows_and_cols_book().find("value_2_column_B") \
                       and self._rows_and_cols_book().find("Value_that_not_exist") is False else False

    def test_get_values_below_the_header(self):
        assert self._rows_and_cols_book().header_values("COLUMN D") == [
            'value_1_column_D', 'value_2_column_D', 'value_3_column_D',
            'value_1_column_D_align_center', 'value_1_column_D_center_vertically',
            'environment'
        ]
        assert self._rows_and_cols_book().header_values("value_1_column_E_align_center") == [
            'value_1_column_E_center_vertically', 'John Smith'
        ]

    def test_get_headers(self):
        assert self._rows_and_cols_book().headers() == [
            "COLUMN A", "COLUMN B", "COLUMN C", "COLUMN D", "COLUMN E", "COLUMN F"
        ]

    def test_get_all_data_converted_to_dict(self):
        assert self._data_convert_book().data_convert() == {
            "COLUMN A":
                [
                    'value_1_column_A', 'value_2_column_A', 'value_3_column_A'
                ],
            "COLUMN B":
                [
                    'value_1_column_B', 'value_2_column_B', 'value_3_column_B'
                ],
        }

    def test_get_cols_2_3_in_dict(self):
        """
        Get dict of cols from 2 to 3 and compare.
        """
        assert self._rows_and_cols_book().dict_cols(start=3, end=4) == {
            1:
                [
                    'value_2_column_A', 'value_3_column_A'
                ],
            2:
                [
                    'value_2_column_B', 'value_3_column_B'
                ],
            3:
                [
                    'value_2_column_C', 'value_3_column_C'
                ],
            4:
                [
                    'value_2_column_D', 'value_3_column_D'
                ],
            5:
                [
                    'value_2_column_E', 'value_3_column_E'
                ],
            6:
                [
                    'value_2_column_F', 'value_3_column_F'
                ],
        }

    def test_get_cols_2_3_4_in_list(self):
        """
        Get lists of cols from 2 to 4 and compare.
        """
        assert self._rows_and_cols_book().list_cols(start=2, end=4) == [
            [
                'value_1_column_A', 'value_2_column_A', 'value_3_column_A'
            ],
            [
                'value_1_column_B', 'value_2_column_B', 'value_3_column_B'
            ],
            [
                'value_1_column_C', 'value_2_column_C', 'value_3_column_C'
            ],
            [
                'value_1_column_D', 'value_2_column_D', 'value_3_column_D'
            ],
            [
                'value_1_column_E', 'value_2_column_E', 'value_3_column_E'
            ],
            [
                'value_1_column_F', 'value_2_column_F', 'value_3_column_F'
            ]
        ]

    def test_get_rows_2_3_in_dict(self):
        """
        Get dict of rows from 2 to 4.
        """
        assert self._rows_and_cols_book().dict_rows(start=3, end=4) == {
            3:
            [
                'value_2_column_A',
                'value_2_column_B',
                'value_2_column_C',
                'value_2_column_D',
                'value_2_column_E',
                'value_2_column_F'
            ],
            4:
            [
                'value_3_column_A',
                'value_3_column_B',
                'value_3_column_C',
                'value_3_column_D',
                'value_3_column_E',
                'value_3_column_F'
            ]
        }

    def test_get_rows_2_3_4_in_list(self):
        """
        Get lists rows from 2 to 4.
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
