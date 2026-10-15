"""Unit tests for the independent financial-checks tool.

Only pure helpers are unit-tested here. End-to-end checks require a
recalculated workbook produced by a real spreadsheet engine and were run
against the Marksans candidates during development (see docs/editable-input-map.md
section 5); they are not part of the offline suite.
"""
from decimal import Decimal
from pathlib import Path
import unittest

from modelmaker import Package
from financial_checks import check, value


class FinancialChecksHelperTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        root = Path(__file__).resolve().parent
        cls.package = Package((root / "originals" / "Marksans Keshav.xlsx").read_bytes())

    def test_value_reads_cached_numbers(self):
        self.assertEqual(value(self.package, "DCF 2nd way", "D10"), Decimal("0.06"))
        self.assertEqual(value(self.package, "WACC-DCF", "J39"), Decimal("0.12107404471298937"))

    def test_value_missing_cell_is_none(self):
        self.assertIsNone(value(self.package, "DCF 2nd way", "ZZ999"))

    def test_check_arithmetic_and_tolerances(self):
        result = check("add", Decimal("2"), Decimal("1") + Decimal("1"))
        self.assertTrue(result["pass"])
        result = check("within_tolerance", Decimal("2"), Decimal("2.4"), tolerance=Decimal("0.5"))
        self.assertTrue(result["pass"])
        result = check("too_far", Decimal("2"), Decimal("3"), tolerance=Decimal("0.1"))
        self.assertFalse(result["pass"])
        result = check("missing", None, Decimal("1"))
        self.assertFalse(result["pass"])


if __name__ == "__main__":
    unittest.main()
