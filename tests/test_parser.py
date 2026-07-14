import unittest
import decimal
from csv_modifier import CSVModifierApp, ParsedNumber


class TestCSVModifierParser(unittest.TestCase):
    def test_polish_valid_numbers(self):
        cases = [
            ("1234", "1234", 0),
            ("1234,56", "1234.56", 2),
            ("1 234,56", "1234.56", 2),
            ("1.234,56", "1234.56", 2),
            ("1.234.567,89", "1234567.89", 2),
            ("-1 234,56", "-1234.56", 2),
            ("+1 234,56", "1234.56", 2),
            ("0,5", "0.5", 1),
            (",5", "0.5", 1),
            ("1\u00A0234,56", "1234.56", 2), # NBSP
            ("1\u202F234,56", "1234.56", 2), # Narrow NBSP
            ("\u22121 234,56", "-1234.56", 2), # Unicode minus
            ("500,00", "500.00", 2), # trailing zeros
            ("12345678901234,56", "12345678901234.56", 2)
        ]

        for val, expected_dec_str, expected_orig_dec in cases:
            with self.subTest(val=val):
                res = CSVModifierApp.parse_number(val, 'Polish')
                self.assertIsInstance(res, ParsedNumber)
                self.assertEqual(res.value, decimal.Decimal(expected_dec_str))
                self.assertEqual(res.orig_decimals, expected_orig_dec)
                self.assertEqual(res.orig_text, val)

    def test_polish_invalid_numbers(self):
        cases = [
            "1 23,45", # wrong group
            "1 234.567,89", # mixed group separators
            "1,234.56", # English format
            "1234,", # incomplete fraction
            "1 234,56 zł", # currency
            "1 234 PLN", # currency
            "(1 234,56)", # parens
            "00123", # ID leading zero
            "0000123456", # ID leading zeros
        ]

        for val in cases:
            with self.subTest(val=val):
                res = CSVModifierApp.parse_number(val, 'Polish')
                self.assertEqual(res, val) # Should return original string

    def test_english_valid_numbers(self):
        cases = [
            ("1234", "1234", 0),
            ("1234.56", "1234.56", 2),
            ("1,234.56", "1234.56", 2),
            ("-1,234.56", "-1234.56", 2),
            ("+1,234.56", "1234.56", 2),
            ("0.5", "0.5", 1),
            (".5", "0.5", 1),
            ("500.00", "500.00", 2),
        ]

        for val, expected_dec_str, expected_orig_dec in cases:
            with self.subTest(val=val):
                res = CSVModifierApp.parse_number(val, 'English')
                self.assertIsInstance(res, ParsedNumber)
                self.assertEqual(res.value, decimal.Decimal(expected_dec_str))
                self.assertEqual(res.orig_decimals, expected_orig_dec)
                self.assertEqual(res.orig_text, val)

    def test_english_invalid_numbers(self):
        cases = [
            "1,23.45",
            "1 234,56",
            "1,234,567.89.",
            "1234.",
            "00123",
        ]

        for val in cases:
            with self.subTest(val=val):
                res = CSVModifierApp.parse_number(val, 'English')
                self.assertEqual(res, val)

    def test_large_integer_is_preserved_as_decimal_for_safe_export_handling(self):
        result = CSVModifierApp.parse_number("1234567890123456", "Polish")
        self.assertIsInstance(result, ParsedNumber)
        self.assertEqual(result.value, decimal.Decimal("1234567890123456"))
        self.assertTrue(CSVModifierApp._requires_excel_text(result))

if __name__ == '__main__':
    unittest.main()
