from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
import unittest
from decimal import Decimal


class TestFailedProject1(unittest.TestCase):
    def test_missing_sheet_added_back(self):  # FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=Sheet2!a1")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
                         CellErrorType.BAD_REFERENCE)

        # Add missing sheet
        _, name2 = wb.new_sheet("Sheet2")

        self.assertEqual(wb.get_cell_value(name, "a1"), None)

    def test_negative_errors(self):  # FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "#ERROR!")
        wb.set_cell_contents(name, "a2", "#CIRCREF!")
        wb.set_cell_contents(name, "a3", "#REF!")
        wb.set_cell_contents(name, "a4", "#NAME?")
        wb.set_cell_contents(name, "a5", "#VALUE!")
        wb.set_cell_contents(name, "a6", "#DIV/0!")
        wb.set_cell_contents(name, "b1", "=-a1")
        wb.set_cell_contents(name, "b2", "=-a2")
        wb.set_cell_contents(name, "b3", "=-a3")
        wb.set_cell_contents(name, "b4", "=-a4")
        wb.set_cell_contents(name, "b5", "=-a5")
        wb.set_cell_contents(name, "b6", "=-a6")

        self.assertTrue(isinstance(wb.get_cell_value(name, "b1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b1").get_type(),
                         CellErrorType.PARSE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b3").get_type(),
                         CellErrorType.BAD_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b4"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b4").get_type(),
                         CellErrorType.BAD_NAME)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b5"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b5").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b6"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b6").get_type(),
                         CellErrorType.DIVIDE_BY_ZERO)

    def test_multiple_errors_in_cycle(self):  # FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=#REF! + a2")
        wb.set_cell_contents(name, "a2", "=a1")
        wb.set_cell_contents(name, "b1", "=#VALUE! + b2")
        wb.set_cell_contents(name, "b2", "=b1")
        wb.set_cell_contents(name, "c1", "=7/0 + c2")
        wb.set_cell_contents(name, "c2", "=c1")
        wb.set_cell_contents(name, "d1", "=#NAME? + d2")
        wb.set_cell_contents(name, "d2", "=d1")
        wb.set_cell_contents(name, "e1", "=#ERROR! + e2")
        wb.set_cell_contents(name, "e2", "=e1")
        wb.set_cell_contents(name, "f1", "=#ERROR! + f2")
        wb.set_cell_contents(name, "f2", "=f3 * 7/0")
        wb.set_cell_contents(name, "f3", "=f1 + asdf!a1")
        wb.set_cell_contents(name, "g1", "=g1 / 0")
        wb.set_cell_contents(name, "g2", "=g2 + #REF!")
        wb.set_cell_contents(name, "g3", "=g3 + #DIV/0!")
        wb.set_cell_contents(name, "g4", "=g4 * #REF!")
        wb.set_cell_contents(name, "g5", "=g5 + dne!g5")  # This was the issue.
        wb.set_cell_contents(name, "g6", "=G6 + dne!g5")  # This was the issue.

        self.assertTrue(isinstance(wb.get_cell_value(name, "a1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "b2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "b2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "c1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "c1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "c2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "c2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "d1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "d1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "d2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "d2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "e1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "e1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "e2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "e2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "f1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "f1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "f2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "f2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "f3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "f3").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g2").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g3").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g4"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g4").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g5"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g5").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(wb.get_cell_value(name, "g6"), CellError))
        self.assertEqual(wb.get_cell_value(name, "g6").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)

    def test_cell_self_reference(self):  # FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=A1")

        self.assertTrue(isinstance(wb.get_cell_value(name, "a1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
                         CellErrorType.CIRCULAR_REFERENCE)

    def test_formula_operation_error_literals(self):  # NOT FIXED
        wb = Workbook()
        _, name = wb.new_sheet()
        error_str = {
            "#ERROR!": CellErrorType.PARSE_ERROR,
            "#CIRCREF!": CellErrorType.CIRCULAR_REFERENCE,
            "#REF!": CellErrorType.BAD_REFERENCE,
            "#NAME?": CellErrorType.BAD_NAME,
            "#VALUE!": CellErrorType.TYPE_ERROR,
            "#DIV/0!": CellErrorType.DIVIDE_BY_ZERO
        }

        for error in error_str:
            wb.set_cell_contents(name, "a1", f"=4 + {error}")
            wb.set_cell_contents(name, "a2", f"=(4 + 0) + {error}")
            wb.set_cell_contents(name, "a3", f"=(4 + {error}) + 0")
            wb.set_cell_contents(name, "a4", f"=4 + 4 - {error}")
            wb.set_cell_contents(name, "a5", f"={error} + 4")
            wb.set_cell_contents(name, "a6", f"=-{error} + 4")

            wb.set_cell_contents(name, "s1", f"=4 - {error}")
            wb.set_cell_contents(name, "s2", f"=4 - -{error}")
            wb.set_cell_contents(name, "s3", f"=({error} - 5)")

            wb.set_cell_contents(name, "m1", f"=4 * {error}")

            wb.set_cell_contents(name, "d1", f"=4 / {error}")

            wb.set_cell_contents(name, "p1", f"=({error})")

            wb.set_cell_contents(name, "sf1", f"={error}")

            wb.set_cell_contents(name, "u1", f"=-{error}")

            self.assertTrue(isinstance(wb.get_cell_value(name, "a1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "a2"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a2").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "a3"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a3").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "a4"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a4").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "a5"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a5").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "a6"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "a6").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "s1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "s1").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "s2"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "s2").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "s3"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "s3").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "m1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "m1").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "d1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "d1").get_type(),
                             error_str[error])

            self.assertTrue(isinstance(wb.get_cell_value(name, "p1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "p1").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "sf1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "sf1").get_type(),
                             error_str[error])
            self.assertTrue(isinstance(wb.get_cell_value(name, "u1"),
                            CellError))
            self.assertEqual(wb.get_cell_value(name, "u1").get_type(),
                             error_str[error])

        wb.set_cell_contents(name, "d2", "'#DIV/0!'")
        wb.set_cell_contents(name, "d3", "\"#DIV/0!\"")
        wb.set_cell_contents(name, "d4", "'#DIV/0!")

        self.assertEqual(wb.get_cell_value(name, "d2"), "#DIV/0!'")
        self.assertEqual(wb.get_cell_value(name, "d3"), "\"#DIV/0!\"")
        self.assertEqual(wb.get_cell_value(name, "d4"), "#DIV/0!")

    def test_add_cell_number_whitespaces(self):  # FAIL
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "d1", "=4+ a1")
        wb.set_cell_contents(name, "d2", "=4             + a1           ")
        wb.set_cell_contents(name, "d3", "=4 + a1\n")
        wb.set_cell_contents(name, "d4", "=a1             + 4           ")
        wb.set_cell_contents(name, "d5", "=a1 + 4\n")
        wb.set_cell_contents(name, "d6", "=a1             - 4           ")
        wb.set_cell_contents(name, "d7", "=a1 * 4\n")
        wb.set_cell_contents(name, "d8", "=     a1 * 4\n")
        wb.set_cell_contents(name, "d9", "     =a1 + 4")
        wb.set_cell_contents(name, "d10", "=4 + \ta1")
        wb.set_cell_contents(name, "d11", "=\t4 \n\t\t+ a1")
        wb.set_cell_contents(name, "d12", "=4 +a1")
        wb.set_cell_contents(name, "d13", "=a1 +4")
        wb.set_cell_contents(name, "d14", "=    4 +a1")

        self.assertEqual(wb.get_cell_value(name, "d1"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d2"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d3"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d4"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d5"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d6"), Decimal(-3))
        self.assertEqual(wb.get_cell_value(name, "d7"), Decimal(4))
        self.assertEqual(wb.get_cell_value(name, "d8"), Decimal(4))
        self.assertEqual(wb.get_cell_value(name, "d9"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d10"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d11"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d12"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d13"), Decimal(5))
        self.assertEqual(wb.get_cell_value(name, "d14"), Decimal(5))

    def test_concat_num_computation_string(self):  # NOT FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "e1", "=6 & \"nine\"")
        wb.set_cell_contents(name, "e2", "=(2+4) & \"nine\"")
        wb.set_cell_contents(name, "e3", "=2 + 4 & \"nine\"")
        wb.set_cell_contents(name, "e4", "=1 * 2 + 4 & \"nine\"")
        wb.set_cell_contents(name, "e5", "=2 / 4 & \"nine\"")
        wb.set_cell_contents(name, "e6", "=2 - 4 & \"nine\"")
        wb.set_cell_contents(name, "e7", "=2 * 4 & \"nine\"")
        wb.set_cell_contents(name, "a1", "=2 + (4 & \"nine\")")
        wb.set_cell_contents(name, "a2", "=2 + 4 & \"nine\" + 5")
        wb.set_cell_contents(name, "a3", "=2 + 4 & \"nine\" -9+ 5")

        # self.assertEqual(wb.get_cell_value(name, "e1"), "6nine")
        # self.assertEqual(wb.get_cell_value(name, "e2"), "6nine")
        # self.assertTrue(isinstance(wb.get_cell_value(name, "e3"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "e3").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "e4"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "e4").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "e5"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "e5").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "e6"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "e6").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "e7"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "e7").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "a1"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
        #                  CellErrorType.TYPE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "a2"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "a2").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        # self.assertTrue(isinstance(wb.get_cell_value(name, "a3"), CellError))
        # self.assertEqual(wb.get_cell_value(name, "a3").get_type(),
        #                  CellErrorType.PARSE_ERROR)
        self.assertEqual(wb.get_cell_value(name, "e1"), "6nine")
        self.assertEqual(wb.get_cell_value(name, "e2"), "6nine")
        self.assertEqual(wb.get_cell_value(name, "e3"), "6nine")
        self.assertEqual(wb.get_cell_value(name, "e4"), "6nine")
        self.assertEqual(wb.get_cell_value(name, "e5"), "0.5nine")
        self.assertEqual(wb.get_cell_value(name, "e6"), "-2nine")
        self.assertEqual(wb.get_cell_value(name, "e7"), "8nine")
        self.assertTrue(isinstance(wb.get_cell_value(name, "a1"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a1").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a2"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a2").get_type(),
                         CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(wb.get_cell_value(name, "a3"), CellError))
        self.assertEqual(wb.get_cell_value(name, "a3").get_type(),
                         CellErrorType.TYPE_ERROR)

    def test_trailing_zeros_concat_number_string(self):  # FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "1.000")
        wb.set_cell_contents(name, "a2", "=a1 & \"hello\"")
        wb.set_cell_contents(name, "a3", "=1.000 & \"hello\"")

        self.assertEqual(wb.get_cell_value(name, "a2"), "1hello")
        self.assertEqual(wb.get_cell_value(name, "a3"), "1hello")

    def test_preserve_quoted_string_leading_whitespace(self):  # NOT FIXED
        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "'   hello world")
        wb.set_cell_contents(name, "a2", "'\nhello world")
        wb.set_cell_contents(name, "a3", "''   hello world")
        wb.set_cell_contents(name, "a4", "   '   hello world")
        wb.set_cell_contents(name, "a5", "'   \"   hello world")
        wb.set_cell_contents(name, "a6", "\"    hello world\"")
        wb.set_cell_contents(name, "a7", "'hello world   ")

        self.assertEqual(wb.get_cell_value(name, "a1"), "   hello world")
        self.assertEqual(wb.get_cell_value(name, "a2"), "\nhello world")
        self.assertEqual(wb.get_cell_value(name, "a3"), "'   hello world")
        self.assertEqual(wb.get_cell_value(name, "a4"), "   hello world")
        self.assertEqual(wb.get_cell_value(name, "a5"), "   \"   hello world")
        self.assertEqual(wb.get_cell_value(name, "a6"), "\"    hello world\"")
        self.assertEqual(wb.get_cell_value(name, "a7"), "hello world   ")
