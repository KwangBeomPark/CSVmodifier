import csv
import shutil
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd

from csv_modifier import CSVModifierApp, ParsedNumber


class _Value:
    def __init__(self, value=''):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class _Root:
    def update(self):
        pass

    def update_idletasks(self):
        pass


class TestOutput(unittest.TestCase):
    def test_output_name_includes_timestamp_and_avoids_same_minute_overwrite(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.csv'
            source.touch()
            timestamp = datetime(2030, 1, 2, 3, 4)
            first = Path(CSVModifierApp._build_output_path(source, 'CSV', timestamp))
            self.assertEqual(first.name, 'processed_output_20300102_0304.csv')
            first.touch()
            second = Path(CSVModifierApp._build_output_path(source, 'CSV', timestamp))

        self.assertEqual(second.name, 'processed_output_20300102_0304_02.csv')

    def test_csv_formatter_preserves_each_cell_decimal_scale(self):
        values = [
            CSVModifierApp.parse_number('500,00', 'Polish'),
            CSVModifierApp.parse_number('1,2', 'Polish'),
            CSVModifierApp.parse_number('2,345', 'Polish'),
        ]
        self.assertEqual(
            [CSVModifierApp._format_csv_value(value, ',') for value in values],
            ['500,00', '1,2', '2,345'],
        )

    def test_excel_output_preserves_formats_and_text_protects_large_numbers(self):
        values = [
            CSVModifierApp.parse_number('500,00', 'Polish'),
            CSVModifierApp.parse_number('1,2', 'Polish'),
            CSVModifierApp.parse_number('9999999999999999,99', 'Polish'),
            CSVModifierApp.parse_number('1234567890123456', 'Polish'),
            CSVModifierApp.parse_number('100000000000000,00', 'Polish'),
        ]
        self.assertTrue(all(isinstance(value, ParsedNumber) for value in values))
        frame = pd.DataFrame([values], columns=['two', 'one', 'large_decimal', 'large_integer', 'safe_zeroes'])

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'output.xlsx'
            protected_count = CSVModifierApp._write_excel_output(frame, path)
            workbook = openpyxl.load_workbook(path, data_only=False)
            sheet = workbook.active

        self.assertEqual(protected_count, 2)
        self.assertEqual(sheet['A2'].value, 500)
        self.assertEqual(sheet['A2'].data_type, 'n')
        self.assertEqual(sheet['A2'].number_format, '0.00')
        self.assertEqual(sheet['B2'].value, 1.2)
        self.assertEqual(sheet['B2'].number_format, '0.0')
        self.assertEqual(sheet['C2'].value, '9999999999999999,99')
        self.assertEqual(sheet['D2'].value, '1234567890123456')
        self.assertEqual(sheet['E2'].value, 100000000000000)
        self.assertEqual(sheet['E2'].number_format, '0.00')

    def test_sample_file_regression_keeps_all_columns_and_formats(self):
        repository_root = Path(__file__).resolve().parents[1]
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / 'input.csv'
            shutil.copyfile(repository_root / 'test_data.csv', source)

            app = CSVModifierApp.__new__(CSVModifierApp)
            app.filepath = _Value(str(source))
            app.delimiter = _Value(',')
            app.num_format = _Value('Polish')
            app.max_cols = _Value('5')
            app.out_format = _Value('CSV')
            app.log_text = _Value()
            app.root = _Root()
            app._set_progress = lambda *args, **kwargs: None

            with patch('csv_modifier.messagebox.showerror') as show_error, patch(
                'csv_modifier.messagebox.showwarning'
            ) as show_warning, patch('csv_modifier.messagebox.showinfo'):
                app.process_csv()

            show_error.assert_not_called()
            show_warning.assert_not_called()
            output_files = list(Path(directory).glob('processed_output_*.csv'))
            self.assertEqual(len(output_files), 1)
            output = output_files[0]
            with output.open(encoding='utf-8-sig', newline='') as file:
                rows = list(csv.reader(file, delimiter=';'))

        self.assertEqual(rows[0], ['Id', 'Name', 'Balance', 'Date', 'Notes'])
        self.assertEqual(rows[1][2], '1234,56')
        self.assertEqual(rows[2][2], '500,00')
        self.assertEqual(rows[3][2], '1000000,00')
        self.assertEqual(rows[1][4], 'First line Second line')
