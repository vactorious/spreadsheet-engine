from decimal import Decimal
from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
import unittest


class TestSheetCopyingFunctionality(unittest.TestCase):
    def test_copy(self):
        """Unit test for copying sheets."""

        wb = Workbook()

        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()

        copy_idx, copy_name = wb.copy_sheet(name)
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1"])
        self.assertEqual(copy_idx, 2)
        self.assertEqual(copy_name, "Sheet1_1")

        # Copy again
        copy_idx2, copy_name2 = wb.copy_sheet(name)
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1",
                                            "Sheet1_2"])
        self.assertEqual(copy_idx2, 3)
        self.assertEqual(copy_name2, "Sheet1_2")

        # Copy a copied sheet
        copy_idx3, copy_name3 = wb.copy_sheet(name + "_1")
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1",
                                            "Sheet1_2", "Sheet1_1_1"])
        self.assertEqual(copy_idx3, 4)
        self.assertEqual(copy_name3, "Sheet1_1_1")

    def test_copy_case_insensitive(self):
        """Test that the sheet name match is case-insensitive."""

        wb = Workbook()
        wb.new_sheet()
        wb.new_sheet()

        wb.copy_sheet("sHeEt1")
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1"])

        # Copy again
        wb.copy_sheet("sheet1")
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1",
                                            "Sheet1_2"])

        # Copy a copied sheet
        wb.copy_sheet("SHEET1_1")
        self.assertEqual(wb.list_sheets(), ["Sheet1", "Sheet2", "Sheet1_1",
                                            "Sheet1_2", "Sheet1_1_1"])

    def test_copy_keyerror(self):
        """Test that the proper errors are raised."""

        wb = Workbook()
        wb.new_sheet()

        # If the specified sheet name is not found, a KeyError is raised.
        self.assertRaises(KeyError, wb.copy_sheet, "Sheet4")

    def test_copy_cells(self):
        """Test that cells are properly copied to the copy sheet."""

        wb = Workbook()

        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "b1", "a")
        wb.set_cell_contents(name, "c1", "=10*4")
        wb.set_cell_contents(name, "d1", "=#ERROR!")
        wb.set_cell_contents(name, "e1", "=#ERROR!")

        _, copy_name = wb.copy_sheet(name)

        self.assertEqual(wb.get_cell_contents(copy_name, "a1"), "1")
        self.assertEqual(wb.get_cell_contents(copy_name, "b1"), "a")
        self.assertEqual(wb.get_cell_contents(copy_name, "c1"), "=10*4")
        self.assertEqual(wb.get_cell_contents(copy_name, "d1"), "=#ERROR!")

        a1_val = wb.get_cell_value(copy_name, "a1")
        b1_val = wb.get_cell_value(copy_name, "b1")
        c1_val = wb.get_cell_value(copy_name, "c1")
        d1_val = wb.get_cell_value(copy_name, "d1")

        self.assertEqual(a1_val, 1)
        self.assertEqual(b1_val, "a")
        self.assertEqual(c1_val, 40)
        self.assertTrue(isinstance(d1_val, CellError))
        self.assertEqual(d1_val.get_type(), CellErrorType.PARSE_ERROR)

    def test_ref_to_copy(self):
        """Tests that a cell referencing the copy sheet in the original sheet
        will be updated with a CIRC_REF error after copying."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # References an undefined sheet at first.
        wb.set_cell_contents(name, "a1", "=Sheet1_1!a1")
        a1_val = wb.get_cell_value(name, "a1")
        self.assertTrue(isinstance(a1_val, CellError))
        self.assertEqual(a1_val.get_type(), CellErrorType.BAD_REFERENCE)

        _, copy_name = wb.copy_sheet(name)
        self.assertEqual(copy_name, "Sheet1_1")

        # References itself so should raise CIRCULAR_REFERENCE error.
        a1_copy_val = wb.get_cell_value(copy_name, "a1")
        self.assertTrue(isinstance(a1_copy_val, CellError))
        self.assertEqual(a1_copy_val.get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)

        # Check a1 now after copy was created. CIRC_REF should propogate.
        a1_val = wb.get_cell_value(name, "a1")
        self.assertTrue(isinstance(a1_val, CellError))
        self.assertEqual(a1_val.get_type(), CellErrorType.CIRCULAR_REFERENCE)

    def test_ref_to_copy_update(self):
        """Tests that a cell referencing the copy sheet will be resolved in the
        original aftering setting contents for the copy sheet."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # References an undefined sheet at first.
        wb.set_cell_contents(name, "a1", "=Sheet1_1!b1")
        a1_val = wb.get_cell_value(name, "a1")
        self.assertTrue(isinstance(a1_val, CellError))
        self.assertEqual(a1_val.get_type(), CellErrorType.BAD_REFERENCE)

        # After the sheet is copied and we set contents for the copy's b1 cell,
        # the original a1 celll's error should resolve accordingly.
        _, copy_name = wb.copy_sheet(name)
        self.assertEqual(copy_name, "Sheet1_1")

        wb.set_cell_contents(copy_name, "b1", "5")
        b1_val = wb.get_cell_value(copy_name, "b1")
        self.assertEqual(b1_val, Decimal("5"))

        a1_val = wb.get_cell_value(name, "a1")
        self.assertEqual(a1_val, Decimal("5"))

    def test_independent_copy(self):
        """Tests that a copy is updated independently from its original."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "b2", "2")
        _, copy_name = wb.copy_sheet(name)

        wb.set_cell_contents(name, "a1", "5")
        wb.set_cell_contents(copy_name, "b2", "6")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "5")
        self.assertEqual(wb.get_cell_contents(copy_name, "a1"), "1")
        self.assertEqual(wb.get_cell_contents(name, "b2"), "2")
        self.assertEqual(wb.get_cell_contents(copy_name, "b2"), "6")

        self.assertEqual(wb.get_cell_value(name, "a1"), Decimal("5"))
        self.assertEqual(wb.get_cell_value(copy_name, "a1"), Decimal("1"))
        self.assertEqual(wb.get_cell_value(name, "b2"), Decimal("2"))
        self.assertEqual(wb.get_cell_value(copy_name, "b2"), Decimal("6"))
