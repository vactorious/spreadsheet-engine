from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
import sheets
from decimal import Decimal
import unittest


class TestFunctionFunctionality(unittest.TestCase):
    def test_bool_and(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=and(True)")
        wb.set_cell_contents("Sheet1", "A2", "=and(False)")
        wb.set_cell_contents("Sheet1", "A3", "=and(True, False)")
        wb.set_cell_contents("Sheet1", "A4", "=and(True, True, 1, \"hi\")")
        wb.set_cell_contents("Sheet1", "A5", "=and(1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A6", "=and(0, 1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A7", "=and(0, 0, False, False)")

        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A3"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A4"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A5"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A6"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A7"))

    def test_bool_or(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=or(True)")
        wb.set_cell_contents("Sheet1", "A2", "=or(False)")
        wb.set_cell_contents("Sheet1", "A3", "=or(True, False)")
        wb.set_cell_contents("Sheet1", "A4", "=or(True, True, 1, \"hi\")")
        wb.set_cell_contents("Sheet1", "A5", "=or(1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A6", "=or(0, 1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A7", "=or(0, 0, False, False)")

        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A3"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A4"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A5"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A6"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A7"))

    def test_bool_not(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=not(True)")
        wb.set_cell_contents("Sheet1", "A2", "=not(False)")
        wb.set_cell_contents("Sheet1", "A3", "=not(0)")
        wb.set_cell_contents("Sheet1", "A4", "=not(1)")
        wb.set_cell_contents("Sheet1", "A5", "=not(\"hi\")")

        self.assertFalse(wb.get_cell_value("Sheet1", "A1"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A3"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A4"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A5"))

    def test_bool_xor(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=xor(True)")
        wb.set_cell_contents("Sheet1", "A2", "=xor(False)")
        wb.set_cell_contents("Sheet1", "A3", "=xor(True, False)")
        wb.set_cell_contents("Sheet1", "A4", "=xor(True, True, 1, \"hi\")")
        wb.set_cell_contents("Sheet1", "A5", "=xor(1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A6", "=xor(0, 1, 2, 3, 4)")
        wb.set_cell_contents("Sheet1", "A7", "=xor(0, 0, False, False)")
        wb.set_cell_contents("Sheet1", "A8", "=xor(True, 0, 1, 2, 3, 4)")

        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A3"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A4"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A5"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A6"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A7"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A8"))

    def test_string_match_exact(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=exact(\"hi\", \"hi\")")
        wb.set_cell_contents("Sheet1", "A2", "=exact(\"hi\", \"Hi\")")
        wb.set_cell_contents("Sheet1", "A3", "=exact(True, True)")
        wb.set_cell_contents("Sheet1", "A4", "=exact(1, 1)")

        self.assertTrue(wb.get_cell_value("Sheet1", "A1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "A2"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A3"))
        self.assertTrue(wb.get_cell_value("Sheet1", "A4"))

    def test_conditional_if(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=if(True, \"hi\")")
        wb.set_cell_contents("Sheet1", "A2", "=if(True, \"Hi\", 9)")
        wb.set_cell_contents("Sheet1", "A3", "=if(False, True)")
        wb.set_cell_contents("Sheet1", "A4", "=if(False, 1, 0)")
        wb.set_cell_contents("Sheet1", "A5", "=if(0, 1)")
        wb.set_cell_contents("Sheet1", "A6", "=if(1, 3)")
        wb.set_cell_contents("Sheet1", "A7", "=if(\"hi\", \"yes\", \"no\")")

        self.assertEqual(wb.get_cell_value("Sheet1", "A1"), "hi")
        self.assertEqual(wb.get_cell_value("Sheet1", "A2"), "Hi")
        self.assertEqual(wb.get_cell_value("Sheet1", "A3"), False)
        self.assertEqual(wb.get_cell_value("Sheet1", "A4"), 0)
        self.assertEqual(wb.get_cell_value("Sheet1", "A5"), False)
        self.assertEqual(wb.get_cell_value("Sheet1", "A6"), 3)
        self.assertEqual(wb.get_cell_value("Sheet1", "A7"), "yes")

    def test_conditional_iferror(self):
        wb = Workbook()
        wb.new_sheet()

    def test_conditional_choose(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=choose(1, \"a\")")
        wb.set_cell_contents("Sheet1", "A2", "=choose(2, \"Hi\", 9)")

        self.assertEqual(wb.get_cell_value("Sheet1", "A1"), "a")
        self.assertEqual(wb.get_cell_value("Sheet1", "A2"), 9)

    def test_informational_isblank(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "B1", "=isblank(A1)")

        self.assertTrue(wb.get_cell_value("Sheet1", "B1"))

        wb.set_cell_contents("Sheet1", "A1", "apple")
        wb.set_cell_contents("Sheet1", "C1", "=isblank(A1)")

        self.assertFalse(wb.get_cell_value("Sheet1", "B1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "C1"))

        wb.set_cell_contents("Sheet1", "F1", "=ISBLANK(\"\")")
        wb.set_cell_contents("Sheet1", "F2", "=ISBLANK(FALSE)")
        wb.set_cell_contents("Sheet1", "F3", "=ISBLANK(0)")

        self.assertFalse(wb.get_cell_value("Sheet1", "F1"))
        self.assertFalse(wb.get_cell_value("Sheet1", "F2"))
        self.assertFalse(wb.get_cell_value("Sheet1", "F3"))

    def test_informational_iserror(self):
        wb = Workbook()
        wb.new_sheet()

    def test_informational_version(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=version()")
        self.assertEqual(wb.get_cell_value("Sheet1", "A1"), sheets.version)

    def test_indirect(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "B2")
        wb.set_cell_contents("Sheet1", "A2", "420")
        wb.set_cell_contents("Sheet1", "A3", "=indirect(A1)")

    def test_empty_cell_arguments(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "B1", "=IF(TRUE, A1)")
        wb.set_cell_contents("Sheet1", "C1", "=IF(FALSE, A1)")
        wb.set_cell_contents("Sheet1", "B2", "'")
        wb.set_cell_contents("Sheet1", "C2", "=EXACT(A2, B2)")

        self.assertEqual(wb.get_cell_value("Sheet1", "B1"), Decimal('0'))
        self.assertEqual(wb.get_cell_value("Sheet1", "C1"), False)
        self.assertEqual(wb.get_cell_value("Sheet1", "B2"), "")
        self.assertEqual(wb.get_cell_value("Sheet1", "C2"), True)

    def test_iserror_iferror_cycles(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=B1")
        wb.set_cell_contents("Sheet1", "B1", "=A1")
        wb.set_cell_contents("Sheet1", "C1", "=ISERROR(B1)")
        wb.set_cell_contents("Sheet1", "A2", "=ISERROR(B2)")
        wb.set_cell_contents("Sheet1", "B2", "=ISERROR(A2)")
        wb.set_cell_contents("Sheet1", "C2", "=ISERROR(B2)")
        wb.set_cell_contents("Sheet1", "A3", "=B1")
        wb.set_cell_contents("Sheet1", "B3", "=INDIRECT(\"A1\")")

        self.assertTrue(isinstance(wb.get_cell_value(name, "A1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "A1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "B1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "B1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "C1"), True)
        self.assertTrue(isinstance(wb.get_cell_value(name, "A2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "A2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "B2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "B2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertEqual(wb.get_cell_value("Sheet1", "C2"), True)
        self.assertTrue(isinstance(wb.get_cell_value(name, "A3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "A3").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "B3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "B3").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)

    # TODO: convert the EvalExpressions Transformer to Interpreter.
    #       This test will fail with a Transformer.
    # def test_conditional_functions_and_cycles(self):
    #     wb = Workbook()
    #     _, name = wb.new_sheet()

    #     wb.set_cell_contents("Sheet1", "A1", "=IF(A2, B1, C1")
    #     wb.set_cell_contents("Sheet1", "B1", "=A1")
    #     wb.set_cell_contents("Sheet1", "C1", "5")
    #     wb.set_cell_contents("Sheet1", "A2", "FALSE")

    #     self.assertEqual(wb.get_cell_value("Sheet1", "A1"), 5)
    #     self.assertEqual(wb.get_cell_value("Sheet1", "B1"), 5)
    #     self.assertEqual(wb.get_cell_value("Sheet1", "C1"), 5)

    def test_bad_name(self):
        wb = Workbook()
        wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=andd()")
        wb.set_cell_contents("Sheet1", "B1", "=ornot()")
        wb.set_cell_contents("Sheet1", "C1", "=notor()")

    def test_wrong_number_of_args(self):
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents("Sheet1", "A1", "=and()")
        wb.set_cell_contents("Sheet1", "B1", "=or()")
        wb.set_cell_contents("Sheet1", "C1", "=not()")
        wb.set_cell_contents("Sheet1", "C2", "=not(1, 2)")
        wb.set_cell_contents("Sheet1", "D1", "=xor()")
        wb.set_cell_contents("Sheet1", "E1", "=exact(\"foo\")")
        wb.set_cell_contents("Sheet1", "E2", "=exact(\"a\", \"a\", \"a\")")
        wb.set_cell_contents("Sheet1", "F1", "=if(true)")
        wb.set_cell_contents("Sheet1", "F2", "=if(true, 1, 2, 3)")
        wb.set_cell_contents("Sheet1", "G1", "=iferror()")
        wb.set_cell_contents("Sheet1", "G2", "=iferror(1, 2, 3)")
        wb.set_cell_contents("Sheet1", "H1", "=choose()")
        wb.set_cell_contents("Sheet1", "H2", "=choose(1)")
        wb.set_cell_contents("Sheet1", "I1", "=isblank()")
        wb.set_cell_contents("Sheet1", "I2", "=isblank(1, 2)")
        wb.set_cell_contents("Sheet1", "J1", "=iserror()")
        wb.set_cell_contents("Sheet1", "J2", "=iserror(1, 2)")
        wb.set_cell_contents("Sheet1", "K1", "=version(1)")
        wb.set_cell_contents("Sheet1", "L1", "=indirect(\"a1\", \"b1\")")

        self.assertTrue(isinstance(wb.get_cell_value(name, "A1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "A1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "B1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "B1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "C1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "C1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "C2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "C2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "D1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "D1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "E2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "E2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "F1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "F1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "F2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "F2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "G1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "G1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "G2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "G2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "H1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "H1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "H2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "H2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "I1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "I1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "I2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "I2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "J1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "J1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "J2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "J2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "K1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "K1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "L1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "L1").get_type(),
                         CellErrorType.TYPE_ERROR)
