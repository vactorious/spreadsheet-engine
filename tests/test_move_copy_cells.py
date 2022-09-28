from sheets.workbook import Workbook
import unittest
from decimal import Decimal


class TestCellMovingFunctionality(unittest.TestCase):
    def test_simple_move(self):
        """Unit test for a simple cell move."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        wb.move_cells(name, "a1", "c1", "a2")

        self.assertEqual(wb.get_cell_contents(name, "a1"), None)
        self.assertEqual(wb.get_cell_contents(name, "b1"), None)
        self.assertEqual(wb.get_cell_contents(name, "c1"), None)
        self.assertEqual(wb.get_cell_contents(name, "a2"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "b2"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "c2"), "=a2*b2")

    def test_overlap_move(self):
        """Unit test for a cell move when source and target areas overlap."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        wb.move_cells(name, "a1", "c1", "B1")

        self.assertEqual(wb.get_cell_contents(name, "a1"), None)
        self.assertEqual(wb.get_cell_contents(name, "b1"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "c1"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "d1"), "=b1*c1")

    def test_absolute_cell_refs_move(self):
        """Unit test for moving cells with absolute cell references."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "b1", "2")

        # No part of the cell ref is modified by moving
        wb.set_cell_contents(name, "c1", "=$a$1+$b$1")
        wb.move_cells(name, "c1", "c1", "c2")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "1")
        self.assertEqual(wb.get_cell_contents(name, "b1"), "2")
        self.assertEqual(wb.get_cell_contents(name, "c1"), None)
        self.assertEqual(wb.get_cell_contents(name, "c2"), "=$a$1+$b$1")
        self.assertEqual(wb.get_cell_value(name, "c2"), Decimal(3))

        # Column is not modified by moving, but the row is modified
        wb.set_cell_contents(name, "d1", "=$a1+$b1")
        wb.move_cells(name, "d1", "d1", "y8")

        self.assertEqual(wb.get_cell_contents(name, "d1"), None)
        self.assertEqual(wb.get_cell_contents(name, "y8"), "=$a8+$b8")
        self.assertEqual(wb.get_cell_value(name, "y8"), Decimal(0))

        # Row is not modified by moving, but the column is modified
        wb.set_cell_contents(name, "a9", "=a$1+b$1")
        wb.move_cells(name, "a9", "a9", "j9")

        self.assertEqual(wb.get_cell_contents(name, "a9"), None)
        self.assertEqual(wb.get_cell_contents(name, "j9"), "=j$1+k$1")
        self.assertEqual(wb.get_cell_value(name, "j9"), Decimal(0))

    def test_out_of_bounds_move(self):
        """Unit test for a cell move to out of bounds."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        self.assertRaises(ValueError, wb.move_cells, name, "a1", "c1", "zzzz1")

    def test_move_cell_other_sheet(self):
        """Unit test for cell moving to other sheets."""
        # Test that sheet names are case insensitive

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")
        _, name2 = wb.new_sheet()

        wb.move_cells(name, "a1", "c1", "a2", "Sheet2")

        self.assertEqual(wb.get_cell_contents(name, "a1"), None)
        self.assertEqual(wb.get_cell_contents(name, "b1"), None)
        self.assertEqual(wb.get_cell_contents(name, "c1"), None)
        self.assertEqual(wb.get_cell_contents(name2, "a2"), "'123")
        self.assertEqual(wb.get_cell_contents(name2, "b2"), "5.3")
        self.assertEqual(wb.get_cell_contents(name2, "c2"), "=a2*b2")

    def test_move_cell_nonexistent_sheet(self):
        """Unit test for cell moving to nonexistent sheets."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        self.assertRaises(KeyError,
                          wb.move_cells, name, "a1", "c1", "a2", "dne")

    def test_move_cell_invalid_loc(self):
        """Unit test for cell moving to invalid cell locations."""
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")
        wb.set_cell_contents(name, "a2", "'456")
        wb.set_cell_contents(name, "b2", "6.7")
        wb.set_cell_contents(name, "c2", "=a2/b2")

        wb.move_cells(name, "a1", "c2", "zzzx9998")

        self.assertEqual(wb.get_cell_contents(name, "a1"), None)
        self.assertEqual(wb.get_cell_contents(name, "b1"), None)
        self.assertEqual(wb.get_cell_contents(name, "c1"), None)
        self.assertEqual(wb.get_cell_contents(name, "a2"), None)
        self.assertEqual(wb.get_cell_contents(name, "b2"), None)
        self.assertEqual(wb.get_cell_contents(name, "c2"), None)

        self.assertRaises(ValueError,
                          wb.move_cells, name, "a1", "c1", "zzzzz99999")
        self.assertRaises(ValueError,
                          wb.move_cells, name, "a1", "c1", "zzzz9999")
        self.assertRaises(ValueError,
                          wb.move_cells, name, "a1", "c1", "zzzy9999")

        self.assertEqual(wb.get_cell_contents(name, "zzzx9998"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "zzzy9998"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "zzzz9998"),
                         "=zzzx9998*zzzy9998")
        self.assertEqual(wb.get_cell_contents(name, "zzzx9999"), "'456")
        self.assertEqual(wb.get_cell_contents(name, "zzzy9999"), "6.7")
        self.assertEqual(wb.get_cell_contents(name, "zzzz9999"),
                         "=zzzx9999/zzzy9999")

    def test_move_cell_becomes_invalid(self):
        """Unit test for cell moving to invalid cell locations."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "2.2")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")
        wb.set_cell_contents(name, "a2", "4.5")
        wb.set_cell_contents(name, "b2", "3.1")
        wb.set_cell_contents(name, "c2", "=a2*b2")

        wb.move_cells(name, "c1", "c2", "b1")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "2.2")
        self.assertEqual(wb.get_cell_contents(name, "b1"), "=#REF!*a1")
        self.assertEqual(wb.get_cell_contents(name, "c1"), None)
        self.assertEqual(wb.get_cell_contents(name, "a2"), "4.5")
        self.assertEqual(wb.get_cell_contents(name, "b2"), "=#REF!*a2")
        self.assertEqual(wb.get_cell_contents(name, "c2"), None)


class TestCellCopyingFunctionality(unittest.TestCase):
    def test_simple_copy(self):
        """Unit test for a simple cell copy."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        wb.copy_cells(name, "a1", "c1", "a2")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "b1"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "c1"), "=a1*b1")
        self.assertEqual(wb.get_cell_contents(name, "a2"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "b2"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "c2"), "=a2*b2")

    def test_overlap_copy(self):
        """Unit test for a cell copy when source and target areas overlap."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        wb.set_cell_contents(name, "b1", "5.3")
        wb.set_cell_contents(name, "c1", "=a1*b1")

        wb.copy_cells(name, "a1", "c1", "b1")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "b1"), "'123")
        self.assertEqual(wb.get_cell_contents(name, "c1"), "5.3")
        self.assertEqual(wb.get_cell_contents(name, "d1"), "=b1*c1")

    def test_absolute_cell_refs_copy(self):
        """Unit test for copying cells with absolute cell references."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "b1", "2")

        # No part of the cell ref is modified by moving
        wb.set_cell_contents(name, "c1", "=$a$1+$b$1")
        wb.copy_cells(name, "c1", "c1", "c2")

        self.assertEqual(wb.get_cell_contents(name, "a1"), "1")
        self.assertEqual(wb.get_cell_contents(name, "b1"), "2")
        self.assertEqual(wb.get_cell_contents(name, "c1"), "=$a$1+$b$1")
        self.assertEqual(wb.get_cell_contents(name, "c2"), "=$a$1+$b$1")
        self.assertEqual(wb.get_cell_value(name, "c2"), Decimal(3))

        # Column is not modified by moving, but the row is modified
        wb.set_cell_contents(name, "d1", "=$a1+$b1")
        wb.copy_cells(name, "d1", "d1", "y8")

        self.assertEqual(wb.get_cell_contents(name, "d1"), "=$a1+$b1")
        self.assertEqual(wb.get_cell_contents(name, "y8"), "=$a8+$b8")
        self.assertEqual(wb.get_cell_value(name, "y8"), Decimal(0))

        # Row is not modified by moving, but the column is modified
        wb.set_cell_contents(name, "a9", "=a$1+b$1")
        wb.copy_cells(name, "a9", "a9", "j9")

        self.assertEqual(wb.get_cell_contents(name, "a9"), "=a$1+b$1")
        self.assertEqual(wb.get_cell_contents(name, "j9"), "=j$1+k$1")
        self.assertEqual(wb.get_cell_value(name, "j9"), Decimal(0))
