from sheets.workbook import Workbook
import unittest


class TestBooleanFunctionality(unittest.TestCase):
    def test_boolean(self):
        """Unit tests for boolean literals."""

        wb = Workbook()
        wb.new_sheet()
        wb.set_cell_contents("Sheet1", "A1", "true")
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A1"), bool))
        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))

        wb.set_cell_contents("Sheet1", "A2", "=FaLsE")
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A2"), bool))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))

        wb.set_cell_contents("Sheet1", "A3", "TRUE      ")
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A2"), bool))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))

        wb.set_cell_contents("Sheet1", "A4", "=      FaLsE")
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A4"), bool))
        self.assertFalse(wb.get_cell_value("Sheet1", "A4"))


class TestCompareFunctionality(unittest.TestCase):
    def test_comparison(self):
        """Unit tests for formulas with comparisons."""

        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=1<2")
        wb.set_cell_contents("Sheet1", "A2", "=1+1<>3")
        wb.set_cell_contents("Sheet1", "A3", "=\"BLUE\" = \"blue\"")
        wb.set_cell_contents("Sheet1", "A4", "=\"BLUE\" < \"blue\"")
        wb.set_cell_contents("Sheet1", "A5", "=\"BLUE\" > \"blue\"")

        # Based on our testing, all major spreadsheet applications perform
        # string comparison against the lowercase versions of the strings.
        # Therefore, we will also follow this convention.
        # In other words, "a" < "[" will evaluate to FALSE.
        wb.set_cell_contents("Sheet1", "A6", "=\"a\" < \"[\"")

        # Booleans are always greater than strings,
        # and strings are always greater than numbers.
        wb.set_cell_contents("Sheet1", "A7", "=\"12\" > 12")
        wb.set_cell_contents("Sheet1", "A8", "=\"TRUE\" > FALSE")

        wb.set_cell_contents("Sheet1", "A9", "=FALSE < TRUE")

        wb.set_cell_contents("Sheet1", "C1", "poop type")
        wb.set_cell_contents("Sheet1", "C2", "poop")
        wb.set_cell_contents("Sheet1", "C3", "=C1 = C2 & \" type\"")

        # If the non-empty operand is a string,
        # the empty operand becomes a zero-length string ““.
        wb.set_cell_contents("Sheet1", "D1", "")
        wb.set_cell_contents("Sheet1", "E1", "=D1=F1")
        # If the non-empty operand is a number, the empty operand becomes a 0.
        wb.set_cell_contents("Sheet1", "D2", "1")
        wb.set_cell_contents("Sheet1", "E2", "=D2>F2")
        # If the non-empty operand is a Boolean,
        # the empty operand becomes FALSE.
        wb.set_cell_contents("Sheet1", "D3", "False")
        wb.set_cell_contents("Sheet1", "E3", "=D3 == F3")

        # If both cells being compared are empty, then they compare as equal.
        wb.set_cell_contents("Sheet1", "E4", "=Z1 = Z2")

        # TODO: Test error propagation

        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A1"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A2"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A3"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A4"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A5"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A6"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A7"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A8"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "A9"), bool))

        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "C3"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "E1"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "E2"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "E3"), bool))
        self.assertTrue(isinstance(wb.get_cell_value("Sheet1", "E4"), bool))

        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A3"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A4"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A5"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A6"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A7"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A8"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A9"))

        self.assertTrue(wb.get_cell_value("Sheet1", "C3"))

        self.assertTrue(wb.get_cell_value("Sheet1", "E1"))
        self.assertTrue(wb.get_cell_value("Sheet1", "E2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "E3"))
        self.assertTrue(wb.get_cell_value("Sheet1", "E4"))
