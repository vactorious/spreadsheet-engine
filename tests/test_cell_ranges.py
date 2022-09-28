from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
from decimal import Decimal
import unittest


class TestCellRangeFunctionality(unittest.TestCase):
    def test_cell_ranges(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        # test parse error for ranges in anything other than function
        wb.set_cell_contents("Sheet1", "C1", "=C2:C4")
        wb.set_cell_contents("Sheet1", "E1", "=IF(C2:C4, 1, 2)")
        # test type error for ranges in functions that don't handle ranges
        # (like exact)
        wb.set_cell_contents("Sheet1", "D1", "=EXACT(A2:A5, B2:B5)")

        self.assertTrue(isinstance(wb.get_cell_value(name, "C1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "C1").get_type(),
                         CellErrorType.PARSE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "D1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_range_min(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "33")
        wb.set_cell_contents(name, "A2", "22")
        wb.set_cell_contents(name, "A3", "11")
        wb.set_cell_contents(name, "A4", "44")
        wb.set_cell_contents(name, "B1", "32")
        wb.set_cell_contents(name, "B2", "21")
        wb.set_cell_contents(name, "B3", "10")
        wb.set_cell_contents(name, "B4", "43")
        wb.set_cell_contents(name, "A5", "=MIN(A1:B4)")
        wb.set_cell_contents(name, "B5", "=MIN(Sheet1!B4:A1)")
        wb.set_cell_contents(name, "A6", "=MIN(B1:A4)")
        wb.set_cell_contents(name, "B6", "=MIN(A4:B1)")
        wb.set_cell_contents(name, "F1", "=MIN(A4:B1, 1)")

        wb.set_cell_contents(name, "C1", "=MIN(L15:M20)")

        wb.set_cell_contents(name, "D1", "lmao")
        wb.set_cell_contents(name, "D2", "=MIN(D1)")
        wb.set_cell_contents(name, "E1", "=MIN()")

        self.assertEqual(wb.get_cell_value(name, "A5"), Decimal('10'))
        self.assertEqual(wb.get_cell_value(name, "B5"), Decimal('10'))
        self.assertEqual(wb.get_cell_value(name, "A6"), Decimal('10'))
        self.assertEqual(wb.get_cell_value(name, "B6"), Decimal('10'))
        self.assertEqual(wb.get_cell_value(name, "F1"), Decimal('1'))
        self.assertEqual(wb.get_cell_value(name, "C1"), Decimal('0'))

        self.assertTrue(isinstance(wb.get_cell_value(name, "D2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_range_max(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "33")
        wb.set_cell_contents(name, "A2", "22")
        wb.set_cell_contents(name, "A3", "11")
        wb.set_cell_contents(name, "A4", "44")
        wb.set_cell_contents(name, "B1", "32")
        wb.set_cell_contents(name, "B2", "21")
        wb.set_cell_contents(name, "B3", "10")
        wb.set_cell_contents(name, "B4", "43")
        wb.set_cell_contents(name, "A5", "=MAX(A1:B4)")
        wb.set_cell_contents(name, "F1", "=MAX(A4:B1, 100)")

        wb.set_cell_contents(name, "C1", "=MAX(L15:M20)")

        wb.set_cell_contents(name, "D1", "lmao")
        wb.set_cell_contents(name, "D2", "=MAX(D1)")
        wb.set_cell_contents(name, "E1", "=MAX()")

        self.assertEqual(wb.get_cell_value(name, "A5"), Decimal('44'))
        self.assertEqual(wb.get_cell_value(name, "F1"), Decimal('100'))
        self.assertEqual(wb.get_cell_value(name, "C1"), Decimal('0'))

        self.assertTrue(isinstance(wb.get_cell_value(name, "D2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_range_sum(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "33")
        wb.set_cell_contents(name, "A2", "22")
        wb.set_cell_contents(name, "A3", "11")
        wb.set_cell_contents(name, "A4", "44")
        wb.set_cell_contents(name, "B1", "32")
        wb.set_cell_contents(name, "B2", "21")
        wb.set_cell_contents(name, "B3", "10")
        wb.set_cell_contents(name, "B4", "43")
        wb.set_cell_contents(name, "A5", "=SUM(A1:B4)")
        wb.set_cell_contents(name, "F1", "=SUM(A4:B1, 100)")

        wb.set_cell_contents(name, "C1", "=SUM(L15:M20)")

        wb.set_cell_contents(name, "D1", "lmao")
        wb.set_cell_contents(name, "D2", "=SUM(D1)")
        wb.set_cell_contents(name, "E1", "=SUM()")

        self.assertEqual(wb.get_cell_value(name, "A5"), Decimal('216'))
        self.assertEqual(wb.get_cell_value(name, "F1"), Decimal('316'))
        self.assertEqual(wb.get_cell_value(name, "C1"), Decimal('0'))

        self.assertTrue(isinstance(wb.get_cell_value(name, "D2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_range_average(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "33")
        wb.set_cell_contents(name, "A2", "22")
        wb.set_cell_contents(name, "A3", "11")
        wb.set_cell_contents(name, "A4", "44")
        wb.set_cell_contents(name, "B1", "32")
        wb.set_cell_contents(name, "B2", "21")
        wb.set_cell_contents(name, "B3", "10")
        wb.set_cell_contents(name, "B4", "43")
        wb.set_cell_contents(name, "A5", "=AVERAGE(A1:B4)")
        wb.set_cell_contents(name, "F1", "=AVERAGE(A4:B1, 99)")

        wb.set_cell_contents(name, "C1", "=AVERAGE(L15:M20)")

        wb.set_cell_contents(name, "D1", "lmao")
        wb.set_cell_contents(name, "D2", "=AVERAGE(D1)")
        wb.set_cell_contents(name, "E1", "=AVERAGE()")

        self.assertEqual(wb.get_cell_value(name, "A5"), Decimal('27'))
        self.assertEqual(wb.get_cell_value(name, "F1"), Decimal('35'))

        self.assertTrue(isinstance(wb.get_cell_value(name, "C1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "C1").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

        self.assertTrue(isinstance(wb.get_cell_value(name, "D2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_hlookup(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "Not")
        wb.set_cell_contents(name, "A2", "1")
        wb.set_cell_contents(name, "A3", "2")
        wb.set_cell_contents(name, "A4", "3")
        wb.set_cell_contents(name, "B1", "Numbers")
        wb.set_cell_contents(name, "B3", "101")
        wb.set_cell_contents(name, "B4", "102")
        wb.set_cell_contents(name, "C1", "=HLOOKUP(\"Numbers\", A1:B4, 4)")

        self.assertEqual(wb.get_cell_value(name, "C1"), Decimal('102'))

    def test_vlookup(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "Not")
        wb.set_cell_contents(name, "A2", "1")
        wb.set_cell_contents(name, "A3", "2")
        wb.set_cell_contents(name, "A4", "3")
        wb.set_cell_contents(name, "B1", "Numbers")
        wb.set_cell_contents(name, "B3", "101")
        wb.set_cell_contents(name, "B4", "102")
        wb.set_cell_contents(name, "C1", "=VLOOKUP(\"Not\", A1:B4, 2)")

        self.assertEqual(wb.get_cell_value(name, "C1"), "Numbers")

    def test_if_range(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "B1", "FALSE")
        wb.set_cell_contents(name, "D1", "1")
        wb.set_cell_contents(name, "D4", "100")
        wb.set_cell_contents(name, "A1", "=SUM(IF(B1, C1:C5, D1:D10))")

        self.assertEqual(wb.get_cell_value(name, "A1"), Decimal('101'))

    def test_choose_range(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "B1", "2")
        wb.set_cell_contents(name, "D1", "1")
        wb.set_cell_contents(name, "D4", "100")
        wb.set_cell_contents(name, "A1", "=SUM(CHOOSE(B1, C1:C5, D1:D10))")

        self.assertEqual(wb.get_cell_value(name, "A1"), Decimal('101'))

    def test_indirect_range(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        # wb.set_cell_contents(name, "A1", \
        #       "=IFERROR(VLOOKUP(A2, INDIRECT(B1 & \"!A2:J100\"), 7), \"\")")
        assert True

    # def test_range_cycles(self):
    #     wb = Workbook()
    #     _, name = wb.new_sheet()

    #     wb.set_cell_contents(name, "A1", "=SUM(A2:A4)")
    #     wb.set_cell_contents(name, "A2", "1")
    #     wb.set_cell_contents(name, "A3", "2")
    #     wb.set_cell_contents(name, "A4", "=A1")

    #     self.assertTrue(isinstance(wb.get_cell_value(name, "A1"), CellError))
    #     self.assertEqual(wb.get_cell_value(name, "A1").get_type(),
    #                      CellErrorType.CIRCULAR_REFERENCE)
    #     self.assertTrue(isinstance(wb.get_cell_value(name, "A4"), CellError))
    #     self.assertEqual(wb.get_cell_value(name, "A4").get_type(),
    #                      CellErrorType.CIRCULAR_REFERENCE)
