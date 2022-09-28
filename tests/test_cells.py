from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
from sheets.cell import CellType
import unittest
from decimal import Decimal
import sys


RECURSION_LIMIT = sys.getrecursionlimit()


class TestCellsFunctionality(unittest.TestCase):
    def test_setting_and_getting_cell_contents(self):
        """Unit test for setting and getting cell contents."""

        wb = Workbook()
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "10")
        wb.set_cell_contents(name, "b1", "20")

        wb.set_cell_contents(name, "c1", None)
        wb.set_cell_contents(name, "d1", "")
        wb.set_cell_contents(name, "e1", " ")
        wb.set_cell_contents(name, "f1", "   ")

        self.assertEqual(wb.get_cell_value(name, "a1"), Decimal("10"))
        wb.set_cell_contents(name, "g1", "'")
        wb.set_cell_contents(name, "h1", "''")
        wb.set_cell_contents(name, "i1", "\"")

    def test_getting_cell_values(self):
        """Unit test for setting cell contents and getting their values."""

        pass

    def test_cell_types(self):
        """Unit test for cell types."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "b1", "'123")
        wb.set_cell_contents(name, "b2", "asdf")
        wb.set_cell_contents(name, "c1", "'foo bar")
        wb.set_cell_contents(name, "d1", "")
        wb.set_cell_contents(name, "e1", None)
        wb.set_cell_contents(name, "f1", "=a1+4")
        wb.set_cell_contents(name, "g1", "=a1+b1")
        wb.set_cell_contents(name, "g2", "=a1+b2")
        wb.set_cell_contents(name, "h1", "foo bar")
        wb.set_cell_contents(name, "i1", "#REF")
        wb.set_cell_contents(name, "j1", "#REF!")
        wb.set_cell_contents(name, "k1", "#REF!  ")

        self.assertEqual(wb.get_cell_instance(name, "a1").val_type,
                         CellType.DECIMAL)
        self.assertEqual(wb.get_cell_instance(name, "b1").val_type,
                         CellType.STRING)
        self.assertEqual(wb.get_cell_instance(name, "b2").val_type,
                         CellType.STRING)
        self.assertEqual(wb.get_cell_instance(name, "c1").val_type,
                         CellType.STRING)
        self.assertEqual(wb.get_cell_instance(name, "d1").val_type,
                         CellType.NONE)
        self.assertEqual(wb.get_cell_instance(name, "e1").val_type,
                         CellType.NONE)
        self.assertEqual(wb.get_cell_instance(name, "f1").val_type,
                         CellType.FORMULA)
        self.assertEqual(wb.get_cell_instance(name, "g1").val_type,
                         CellType.FORMULA)
        self.assertEqual(wb.get_cell_instance(name, "g2").val_type,
                         CellType.ERROR)
        self.assertEqual(wb.get_cell_instance(name, "h1").val_type,
                         CellType.STRING)
        self.assertEqual(wb.get_cell_instance(name, "i1").val_type,
                         CellType.STRING)
        self.assertEqual(wb.get_cell_instance(name, "j1").val_type,
                         CellType.ERROR)
        self.assertEqual(wb.get_cell_instance(name, "k1").val_type,
                         CellType.ERROR)

    def test_literal_parser(self):
        """Unit test for parsing literals and getting cell values."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "a2", "9")
        wb.set_cell_contents(name, "a3", "'")
        wb.set_cell_contents(name, "a4", "''")
        wb.set_cell_contents(name, "a5", "'=a1+a2")
        wb.set_cell_contents(name, "a6", "#DIV/0!")

        self.assertEqual(wb.get_cell_value(name, "a5"), "=a1+a2")
        self.assertTrue(isinstance(wb.get_cell_value(name, "a6"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a6").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

    def test_formula_parser(self):
        """Tests for formula parser errors (#ERROR!)."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "a2", "9")

        wb.set_cell_contents(name, "a3", "=a1+a2")

        self.assertEqual(wb.get_cell_value(name, "a3"), Decimal("10"))

        _, name2 = wb.new_sheet("Foo Bar")
        wb.set_cell_contents(name2, "a1", "2")

        wb.set_cell_contents(name, "a4", "=Sheet1!a1!a1")
        wb.set_cell_contents(name, "a5", f"=(a1+a2)*{name2}!a1")
        wb.set_cell_contents(name, "a6", f"=(a1+a2)*\"{name2}\"!a1")
        wb.set_cell_contents(name, "a7", f"=(a1+a2)*'{name2}'!a1")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a4"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a4").get_type(),
                         CellErrorType.PARSE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a5"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a5").get_type(),
                         CellErrorType.PARSE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a6"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a6").get_type(),
                         CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value(name, "a7"), Decimal("20"))

        # Case-insensitive cell references
        wb.set_cell_contents(name, "b1", f"='{name2}'!a1")
        wb.set_cell_contents(name, "b2", f"='{name2.lower()}'!a1")

        self.assertEqual(wb.get_cell_value(name, "b1"),
                         wb.get_cell_value(name, "b2"))

    def test_circ_ref_errors(self):
        """Tests for circular reference cell type errors (#CIRREF!)."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # Simple cycle
        wb.set_cell_contents(name, "A1", "=B1")
        wb.set_cell_contents(name, "B1", "=C1+1")
        wb.set_cell_contents(name, "C1", "=A1+2")
        err_a, err_b, err_c = wb.get_cell_value(name, "A1"),\
            wb.get_cell_value(name, "B1"), wb.get_cell_value(name, "C1")

        self.assertTrue(isinstance(err_a, CellError))
        self.assertTrue(isinstance(err_b, CellError))
        self.assertTrue(isinstance(err_c, CellError))
        self.assertEqual(err_a.get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(err_b.get_type(), CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(err_c.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # Error should propagate to those referencing a cycle
        wb.set_cell_contents(name, "D1", "=B1")
        err_d = wb.get_cell_value(name, "D1")

        self.assertTrue(isinstance(err_d, CellError))
        self.assertEqual(err_d.get_type(), CellErrorType.CIRCULAR_REFERENCE)

        # # Algorithm to find strongly connection components (SCC) shouldn't
        # # be implemented recursively to avoid hitting max recursion stack.
        # cells = []
        # for i in range(1, RECURSION_LIMIT + 10):
        #     wb.set_cell_contents(name, f"e{i}", f"=e{i+1}")
        #     cells.append(wb.get_cell_instance(name, f"e{i}"))
        # try:
        #     wb.set_cell_contents(name, f"e{i}", "=e1")
        # except RecursionError as e:
        #     self.fail(e)

        # self.assertTrue(all(isinstance(v.val, CellError) for v in cells))
        # self.assertTrue(all(v.val.get_type() ==
        #                     CellErrorType.CIRCULAR_REFERENCE for v in cells))

    def test_bad_ref_errors(self):
        """Tests for bad cell reference errors in formulas (#REF!)."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # Error propagation
        wb.set_cell_contents(name, "a1", "#REF!")
        wb.set_cell_contents(name, "a2", "=#REF!")
        wb.set_cell_contents(name, "a3", "=#REF! + 5")
        a1_val = wb.get_cell_value(name, "a1")
        a2_val = wb.get_cell_value(name, "a2")
        a3_val = wb.get_cell_value(name, "a3")

        self.assertTrue(isinstance(a1_val, CellError))
        self.assertEqual(a1_val.get_type(), CellErrorType.BAD_REFERENCE)
        self.assertTrue(isinstance(a2_val, CellError))
        self.assertEqual(a2_val.get_type(), CellErrorType.BAD_REFERENCE)
        self.assertTrue(isinstance(a3_val, CellError))
        self.assertEqual(a3_val.get_type(), CellErrorType.BAD_REFERENCE)

        _, name2 = wb.new_sheet("sheet 2")
        wb.set_cell_contents(name2, "a1", "3")
        wb.set_cell_contents(name, "a4", "=Sheet2!a1+1")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a4"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a4").get_type(),
                         CellErrorType.BAD_REFERENCE)

        wb.set_cell_contents(name, "d1", "=ZZZZ9999")
        wb.set_cell_contents(name, "d2", "=AAAAA1")
        wb.set_cell_contents(name, "d3", "=A10000")

        self.assertEqual(wb.get_cell_value(name, "d1"), None)
        self.assertTrue(isinstance(wb.get_cell_value(name, "d2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "d2").get_type(),
                         CellErrorType.BAD_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "d3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "d3").get_type(),
                         CellErrorType.BAD_REFERENCE)

    def test_bad_name_errors(self):
        """Tests for unrecognizable function names in formulas (#NAME?)."""

        pass

    def test_type_errors(self):
        """Tests for operations performed on unsupported types (#VALUE!)."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # Arithmetic only work with numbers
        wb.set_cell_contents(name, "a1", "abc")
        wb.set_cell_contents(name, "a2", "=3*a1")
        wb.set_cell_contents(name, "a3", "=a1+3")
        wb.set_cell_contents(name, "a4", "foo")
        wb.set_cell_contents(name, "a5", "bar")
        wb.set_cell_contents(name, "a6", "6")
        wb.set_cell_contents(name, "a7", "9")
        wb.set_cell_contents(name, "a8", "=a4 & a5")
        wb.set_cell_contents(name, "a9", "=a4 & a6")
        wb.set_cell_contents(name, "a10", "=a6 & a7")
        wb.set_cell_contents(name, "a11", "=(68+1) & a4 & a5")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a3").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertEqual(wb.get_cell_value(name, "a8"), "foobar")
        self.assertEqual(wb.get_cell_value(name, "a9"), "foo6")
        self.assertEqual(wb.get_cell_value(name, "a10"), "69")
        self.assertEqual(wb.get_cell_value(name, "a11"), "69foobar")

    def test_divide_by_zero(self):
        """Tests for operations that divide by zero (#DIV/0!)."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "a2", "0")

        wb.set_cell_contents(name, "a3", "=a1/a2")
        wb.set_cell_contents(name, "a4", "=a1/0")
        wb.set_cell_contents(name, "a5", "#DIV/0!")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a3").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a4"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a4").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a5"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a5").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

    def test_strip(self):
        """Tests that strings are properly stripped."""

        wb = Workbook()
        idx, name = wb.new_sheet()

        wb.set_cell_contents(name, "s1", "   this string should be stripped")
        self.assertEqual(wb.get_cell_contents(name, "s1"),
                         "this string should be stripped")
        self.assertEqual(wb.get_cell_value(name, "s1"),
                         "this string should be stripped")

        wb.set_cell_contents(name, "s2", "this string should be stripped   ")
        self.assertEqual(wb.get_cell_contents(name, "s2"),
                         "this string should be stripped")
        self.assertEqual(wb.get_cell_value(name, "s2"),
                         "this string should be stripped")

        wb.set_cell_contents(name, "s3", "   first")
        wb.set_cell_contents(name, "s4", "second   ")
        wb.set_cell_contents(name, "s5", "=s3&s4")
        self.assertEqual(wb.get_cell_value(name, "s5"), "firstsecond")

    def test_type_conversions(self):
        """Tests that values are automatically converted to the proper type."""

        wb = Workbook()
        _, name = wb.new_sheet()
        wb.set_cell_contents(name, "a1", "'123")
        self.assertEqual(wb.get_cell_value(name, "a1"), "123")

        wb.set_cell_contents(name, "b1", "5.3")
        self.assertEqual(wb.get_cell_value(name, "b1"), Decimal("5.3"))

        wb.set_cell_contents(name, "c1", "=A1*B1")
        self.assertEqual(wb.get_cell_value(name, "c1"),
                         Decimal("651.9"))

        wb.set_cell_contents(name, "a2", "'     123")
        self.assertEqual(wb.get_cell_value(name, "a2"), "     123")

        # Should have the same behavior as c1
        wb.set_cell_contents(name, "c2", "=A2*B1")
        self.assertEqual(wb.get_cell_value(name, "c2"),
                         Decimal("651.9"))

        # Should be able to evaluate formulas using strings that can be
        # interpreted as numbers.
        wb.set_cell_contents(name, "f4", "1")
        wb.set_cell_contents(name, "f5", "'123")
        wb.set_cell_contents(name, "f1", "=F4+F5")
        self.assertEqual(wb.get_cell_value(name, "f1"),
                         Decimal("124"))

        wb.set_cell_contents(name, "a10", "2")
        wb.set_cell_contents(name, "b10", "'     10")
        wb.set_cell_contents(name, "c10", "=B10-A10")
        self.assertEqual(wb.get_cell_value(name, "c10"),
                         Decimal("8"))

        wb.set_cell_contents(name, "a20", "2")
        wb.set_cell_contents(name, "b20", "'20.")
        wb.set_cell_contents(name, "c20", "=B20*A20")
        self.assertEqual(wb.get_cell_value(name, "c20"),
                         Decimal("40"))

        wb.set_cell_contents(name, "a30", "'2     ")
        wb.set_cell_contents(name, "b30", "30")
        wb.set_cell_contents(name, "c30", "=B30 / A30")
        self.assertEqual(wb.get_cell_value(name, "c30"),
                         Decimal("15"))

    def test_using_empty_cell(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        # If an empty cell is used as a number, it should be 0.
        wb.set_cell_contents(name, "n1", "100")
        wb.set_cell_contents(name, "n2", "=E5 + N1")
        self.assertEqual(wb.get_cell_value(name, "n2"),
                         Decimal("100"))

        # If an empty cell is used as a string,
        # it should be a zero-length string ““.
        wb.set_cell_contents(name, "s1", "s2 should be the same string")
        wb.set_cell_contents(name, "s2", "=G5 & S1")
        self.assertEqual(wb.get_cell_value(name, "s2"),
                         "s2 should be the same string")

    def test_topological_sort(self):
        """Unit tests for the topological sorting algorithm, used to evaluate
        the order in which cells need to be evaluated."""

        wb = Workbook()
        _, name = wb.new_sheet()

        # Regardless of the order in which these cells' contents are set...
        wb.set_cell_contents(name, "a1", "=b1+c1")
        wb.set_cell_contents(name, "b1", "=c1")
        wb.set_cell_contents(name, "c1", "=1")

        self.assertEqual(wb.get_cell_value(name, "a1"), Decimal(2))
        self.assertEqual(wb.get_cell_value(name, "b1"), Decimal(1))
        self.assertEqual(wb.get_cell_value(name, "c1"), Decimal(1))

        # Regardless of the order in which these cells' contents are set...
        wb.set_cell_contents(name, "c2", "=1")
        wb.set_cell_contents(name, "b2", "=c2")
        wb.set_cell_contents(name, "a2", "=b2+c2")

        self.assertEqual(wb.get_cell_value(name, "a2"), Decimal(2))
        self.assertEqual(wb.get_cell_value(name, "b2"), Decimal(1))
        self.assertEqual(wb.get_cell_value(name, "c2"), Decimal(1))


if __name__ == "__main__":
    unittest.main()
