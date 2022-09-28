from decimal import Decimal
from sheets.workbook import Workbook
import unittest


class TestSheetRenamingFunctionality(unittest.TestCase):
    def test_rename_sheet(self):
        """Unit tests for sheet renaming."""

        wb = Workbook()
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "a2", "2")
        wb.set_cell_contents(name, "a3", "=Sheet1!a1")
        wb.set_cell_contents(name2, "b1", "=Sheet1!a1+9")
        wb.set_cell_contents(name2, "b2", "=(Sheet1!a1 + Sheet1!a2)*9")
        wb.set_cell_contents(name2, "b3", "=(-Sheet1!a1+Sheet1!a2)*9")
        wb.set_cell_contents(name2, "b4", "0")
        wb.set_cell_contents(name2, "b5", "=Sheet1!a1*Sheet2!b4")

        wb.rename_sheet("Sheet1", "hello")

        self.assertEqual(wb.get_cell_instance("hello", "a3").contents,
                         "=hello!a1")
        self.assertEqual(wb.get_cell_value("hello", "a3"), Decimal(1))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b1").contents,
                         "=hello!a1+9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b1"), Decimal(10))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b2").contents,
                         "=(hello!a1+hello!a2)*9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b2"), Decimal(27))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b3").contents,
                         "=(-hello!a1+hello!a2)*9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b3"), Decimal(9))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b5").contents,
                         "=hello!a1*Sheet2!b4")
        self.assertEqual(wb.get_cell_value("Sheet2", "b5"), Decimal(0))

        # Quoted sheet name
        wb.rename_sheet("hello", "foo bar")

        self.assertEqual(wb.get_cell_instance("foo bar", "a3").contents,
                         "='foo bar'!a1")
        self.assertEqual(wb.get_cell_value("foo bar", "a3"), Decimal(1))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b1").contents,
                         "='foo bar'!a1+9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b1"), Decimal(10))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b2").contents,
                         "=('foo bar'!a1+'foo bar'!a2)*9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b2"), Decimal(27))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b3").contents,
                         "=(-'foo bar'!a1+'foo bar'!a2)*9")
        self.assertEqual(wb.get_cell_value("Sheet2", "b3"), Decimal(9))
        self.assertEqual(wb.get_cell_instance("Sheet2", "b5").contents,
                         "='foo bar'!a1*Sheet2!b4")

    def test_rename_casing(self):
        """Tests renaming a sheet follows the casing and uniqueness rules.

        The sheet_name match is case-insensitive and the case of the renamed
        sheet is preserved by the workbook, but a sheet cannot be renamed to
        another existing sheet case-insensitive.
        """

        wb = Workbook()
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()

        self.assertRaises(ValueError, wb.rename_sheet, "Sheet1", "Sheet2")
        self.assertRaises(ValueError, wb.rename_sheet, "Sheet1", "sHeEt2")
        wb.rename_sheet("Sheet1", "sHeEt1")
        self.assertEqual(wb.sheet_arr[0].name, "sHeEt1")
        self.assertRaises(ValueError, wb.rename_sheet, "Sheet2", "Sheet1")
        wb.rename_sheet("Sheet1", "Sheet1")
        self.assertEqual(wb.sheet_arr[0].name, "Sheet1")
