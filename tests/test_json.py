from sheets.workbook import Workbook
from sheets.error import CellError, CellErrorType
import unittest
from decimal import Decimal
import os
import json


class TestLoadJSON(unittest.TestCase):
    def test_load_workbook(self):
        """Unit test for loading a basic workbook with one sheet."""

        if not os.path.exists("tests/json/test.json"):
            data = {
                "sheets": [
                    {
                        "name": "Sheet1",
                        "cell-contents": {
                            "A1": "'123",
                            "b1": "5.3",
                            "C1": "=A1*B1"
                        }
                    }
                ]
            }
            with open("tests/json/test.json", "w") as outfile:
                json.dump(data, outfile, indent=4)

        with open("tests/json/test.json", "r") as f:
            wb = Workbook.load_workbook(f)
            self.assertEqual(len(wb.sheet_arr), 1)
            self.assertEqual(wb.get_sheet_extent("Sheet1"), (3, 1))
            self.assertEqual(wb.get_cell_contents("Sheet1", "a1"), "'123")
            self.assertEqual(wb.get_cell_contents("Sheet1", "b1"), "5.3")
            self.assertEqual(wb.get_cell_contents("Sheet1", "c1"), "=A1*B1")
            self.assertEqual(wb.get_cell_value("Sheet1", "a1"), "123")
            self.assertEqual(wb.get_cell_value("Sheet1", "b1"), Decimal("5.3"))
            self.assertEqual(wb.get_cell_value("Sheet1", "c1"),
                             Decimal("651.9"))

    def test_load_sheets(self):
        """Unit test for loading a json with multiple sheets."""

        if not os.path.exists("tests/json/testsheets.json"):
            data = {
                "sheets": [
                    {
                        "name": "Sheet1",
                        "cell-contents": {
                            "A1": "'123",
                            "B1": "5.3",
                            "C1": "=A1*B1"
                        }
                    },
                    {
                        "name": "test",
                        "cell-contents": {
                            "A1": "'123",
                            "B1": "5.3",
                            "E7": "=Sheet1!c1"
                        }
                    }
                ]
            }
            with open("tests/json/testsheets.json", "w") as outfile:
                json.dump(data, outfile, indent=4)

        with open("tests/json/testsheets.json", "r") as f:
            wb = Workbook.load_workbook(f)
            self.assertEqual(len(wb.sheet_arr), 2)
            self.assertEqual(wb.get_sheet_extent("test"), (5, 7))
            self.assertEqual(wb.get_cell_contents("test", "a1"), "'123")
            self.assertEqual(wb.get_cell_contents("test", "b1"), "5.3")
            self.assertEqual(wb.get_cell_contents("test", "e7"), "=Sheet1!c1")
            self.assertEqual(wb.get_cell_value("test", "e7"), Decimal("651.9"))

    def test_load_bad_workbook(self):
        """Test that an error occurs if the workbook json is bad."""

        # Use a testbad.json that is an invalid json format.
        with open("tests/json/testjunk.json", "r") as f:
            self.assertRaises(json.JSONDecodeError, Workbook.load_workbook, f)

    def test_load_keyerror(self):
        """Test that a key error occurs if the json is missing a needed key."""

        if not os.path.exists("tests/json/testmissingname.json"):
            data = {
                "sheets": [
                    {
                        "cell-contents": {
                            "A1": "'123",
                            "b1": "5.3",
                            "C1": "=A1*B1"
                        }
                    }
                ]
            }
            with open("tests/json/testmissingname.json", "w") as outfile:
                json.dump(data, outfile, indent=4)
        with open("tests/json/testmissingname.json", "r") as f:
            self.assertRaises(KeyError, Workbook.load_workbook, f)

        if not os.path.exists("tests/json/testmissingcells.json"):
            data = {
                "sheets": [
                    {
                        "name": "Sheet1"
                    }
                ]
            }
            with open("tests/json/testmissingcells.json", "w") as outfile:
                json.dump(data, outfile, indent=4)
        with open("tests/json/testmissingcells.json", "r") as f:
            self.assertRaises(KeyError, Workbook.load_workbook, f)

        if not os.path.exists("tests/json/testempty.json"):
            data = {}
            with open("tests/json/testempty.json", "w") as outfile:
                json.dump(data, outfile, indent=4)
        with open("tests/json/testempty.json", "r") as f:
            self.assertRaises(KeyError, Workbook.load_workbook, f)

    def test_load_decodeerror(self):
        """Test that a JSONDecodeError occurs if the json cannot be parsed."""

        with open("tests/json/testblank.json", "r") as f:
            self.assertRaises(json.JSONDecodeError, Workbook.load_workbook, f)

    def test_load_bad_type(self):
        """Test that a TypeError is raised when handling an unexpected type."""

        with open("tests/json/testbadtype1.json", "r") as f:
            self.assertRaises(TypeError, Workbook.load_workbook, f)

        with open("tests/json/testbadtype2.json", "r") as f:
            self.assertRaises(TypeError, Workbook.load_workbook, f)

    def test_load_bad_formula(self):
        """Test that cells with bad formulas have their value set to a
        parse error, as usual."""

        if not os.path.exists("tests/json/testbadformula.json"):
            data = {
                "sheets": [
                    {
                        "name": "Sheet1",
                        "cell-contents": {
                            "E1": "=should be a parse error"

                        }
                    }
                ]
            }
            with open("tests/json/testbadformula.json", "w") as outfile:
                json.dump(data, outfile, indent=4)

        with open("tests/json/testbadformula.json", "r") as f:
            wb = Workbook.load_workbook(f)
            e1_contents = wb.get_cell_contents("Sheet1", "e1")
            e1_val = wb.get_cell_value("Sheet1", "e1")

            self.assertEqual(e1_contents, "=should be a parse error")
            self.assertTrue(isinstance(e1_val, CellError))
            self.assertEqual(e1_val.get_type(), CellErrorType.PARSE_ERROR)


class TestSaveJSON(unittest.TestCase):

    def test_save_workbook(self):
        """Unit test for saving a basic workbook with one sheet."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "'123")
        wb.set_cell_contents(name, "B1", "5.3")
        wb.set_cell_contents(name, "C1", "=A1*B1")

        with open("tests/json/testsave.json", "w") as f:
            wb.save_workbook(f)

        # Check if the saved json is the same as the reference test json.
        with open("tests/json/testsave.json", "r") as saved_f:
            saved = json.load(saved_f)

            if not os.path.exists("tests/json/test.json"):
                data = {
                    "sheets": [
                        {
                            "name": "Sheet1",
                            "cell-contents": {
                                "A1": "'123",
                                "b1": "5.3",
                                "C1": "=A1*B1"
                            }
                        }
                    ]
                }
                with open("tests/json/test.json", "w") as outfile:
                    json.dump(data, outfile, indent=4)
            with open("tests/json/test.json", "r") as ref_f:
                ref = json.load(ref_f)
                self.assertEqual(saved, ref)

    def test_save_sheets(self):
        """Unit test for saving a basic workbook with multiple sheets."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "A1", "'123")
        wb.set_cell_contents(name, "B1", "5.3")
        wb.set_cell_contents(name, "C1", "=A1*B1")

        _, name2 = wb.new_sheet("test")
        wb.set_cell_contents(name2, "A1", "'123")
        wb.set_cell_contents(name2, "B1", "5.3")
        wb.set_cell_contents(name2, "E7", "=Sheet1!c1")

        with open("tests/json/testsavesheets.json", "w") as f:
            wb.save_workbook(f)

        # Check if the saved json is the same as the reference test json.
        with open("tests/json/testsavesheets.json", "r") as saved_f:
            saved = json.load(saved_f)

            if not os.path.exists("tests/json/testsheets.json"):
                data = {
                    "sheets": [
                        {
                            "name": "Sheet1",
                            "cell-contents": {
                                "A1": "'123",
                                "B1": "5.3",
                                "C1": "=A1*B1"
                            }
                        },
                        {
                            "name": "test",
                            "cell-contents": {
                                "A1": "'123",
                                "B1": "5.3",
                                "E7": "=Sheet1!c1"
                            }
                        }
                    ]
                }
                with open("tests/json/testsavesheets.json", "w") as outfile:
                    json.dump(data, outfile, indent=4)
            with open("tests/json/testsheets.json", "r") as ref_f:
                ref = json.load(ref_f)
                self.assertEqual(saved, ref)
