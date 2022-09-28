from sheets.workbook import Workbook
import unittest


class TestSheetMovingFunctionality(unittest.TestCase):
    def test_move(self):
        """Unit test for moving sheets."""

        wb = Workbook()

        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()
        _, name3 = wb.new_sheet()

        # Move Sheet1 to the end
        wb.move_sheet(name, 2)
        self.assertEqual(wb.list_sheets(), ["Sheet2", "Sheet3", "Sheet1"])

        # Move to the same place
        wb.move_sheet(name3, 1)
        self.assertEqual(wb.list_sheets(), ["Sheet2", "Sheet3", "Sheet1"])

        wb.move_sheet(name2, 2)
        self.assertEqual(wb.list_sheets(), ["Sheet3", "Sheet1", "Sheet2"])

        # Check workbook attributes are being updated

        wb.move_sheet(name, 0)
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet3", "Sheet2"])

    def test_move_case_insensitive(self):
        """Test that the sheet name match is case-insensitive."""

        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()
        wb.new_sheet()

        wb.move_sheet("sHeEt1", 2)
        self.assertEqual(wb.list_sheets(), ["Sheet2", "Sheet3", "Sheet1"])

        wb.move_sheet("sheet3", 1)
        self.assertEqual(wb.list_sheets(), ["Sheet2", "Sheet3", "Sheet1"])

        wb.move_sheet("SHEET2", 2)
        self.assertEqual(wb.list_sheets(), ["Sheet3", "Sheet1", "Sheet2"])

    def test_move_errors(self):
        """Test that the proper errors are raised."""

        wb = Workbook()

        _, name = wb.new_sheet()

        # If the specified sheet name is not found, a KeyError is raised.
        self.assertRaises(KeyError, wb.move_sheet, "Sheet4", 0)

        # If the index is outside the valid range, an IndexError is raised.
        self.assertRaises(IndexError, wb.move_sheet, "Sheet1", 100)
        self.assertRaises(IndexError, wb.move_sheet, "Sheet1", -10)
