from sheets.workbook import Workbook
import unittest
import cProfile
from pstats import Stats


class TestPerformanceStress(unittest.TestCase):
    def setUp(self):
        self.profiler = cProfile.Profile()
        self.profiler.enable()
        print("\n<<<---")

    def tearDown(self):
        p = Stats(self.profiler)
        p.strip_dirs()
        p.sort_stats('cumtime')
        p.print_stats(0.1)
        print("\n--->>>")

    def test_one_long_chain(self):
        """Integration test for propagating updates through long chains of cell
        references, where each cell depends only on one other cell."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=1")
        for i in range(2, 9999):
            wb.set_cell_contents(name, f"a{i}", f"=a{i-1}")

        wb.set_cell_contents(name, "a1", "=2")

    def test_many_short_chains(self):
        """Integration test for cells that are referenced by many other cells,
        with shallow chains but large amounts of cell updates."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=1")
        for i in range(2, 9999):
            if i % 3 == 0:
                wb.set_cell_contents(name, f"a{i}", "=a1")
            else:
                wb.set_cell_contents(name, f"a{i}", f"=a{i-1}")

        wb.set_cell_contents(name, "a1", "=2")

    def test_one_large_cycle(self):
        """Integration test for large cycles that contain many cells."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=1")
        for i in range(2, 9999):
            wb.set_cell_contents(name, f"a{i}", f"=a{i-1}")

        wb.set_cell_contents(name, "a1", "=a9999")

    def test_many_small_cycles(self):
        """Integration test for many small cycles each containing a small
        number of cells."""

        wb = Workbook()
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=1")
        for i in range(2, 9999):
            if i % 3 == 0:
                wb.set_cell_contents(name, f"a{i}", f"=a{i-2}")
            elif i % 3 == 1:
                wb.set_cell_contents(name, f"a{i}", "=1")
            else:
                wb.set_cell_contents(name, f"a{i}", f"=a{i-1}")

    def test_one_cell_many_cycles(self):
        """Integration test for cell that is part of many different cycles.
        """

        wb = Workbook()
        _, name = wb.new_sheet()

        # b1=b2, b2=a1, a1=b1
        # b3=b4, b4=a1, a1=b3
        # b5=b6, b6=a1, a1=b5
        # ...
        for i in range(1, 9999, 2):
            wb.set_cell_contents(name, f"b{i}", f"=b{i+1}")
            wb.set_cell_contents(name, f"b{i+1}", "=a1")

        wb.set_cell_contents(name, "a1", "=" + '+'.join([f"b{i}"
                             for i in range(1, 9999, 2)]))


class TestPerformanceStress2(unittest.TestCase):
    def setUp(self):
        self.profiler = cProfile.Profile()
        self.profiler.enable()
        print("\n<<<---")

    def tearDown(self):
        p = Stats(self.profiler)
        p.strip_dirs()
        p.sort_stats('cumtime')
        p.print_stats(0.1)
        print("\n--->>>")

    large_wb = None

    def test_z0_loading_large_workbook(self):
        """Integration test for loading a very large workbook file."""
        with open("tests/json/largeworkbook.json", "r") as f:
            self.__class__.large_wb = Workbook.load_workbook(f)

    def test_z1_copying_large_sheet(self):
        """Integration test for copying a very large sheet."""
        self.__class__.large_wb.copy_sheet("Sheet2")

    def test_z4_renaming_large_sheet(self):
        """Integration test for renaming a very large sheet."""
        self.__class__.large_wb.rename_sheet("Sheet1", "bingbong")

    def test_z2_moving_large_cell_area(self):
        """Integration test for moving a very large cell area."""
        self.__class__.large_wb.move_cells("Sheet2", "A1", "SE499", "A2")

    def test_z3_copying_large_cell_area(self):
        """Integration test for copying a very large cell area."""
        self.__class__.large_wb.copy_cells("Sheet3", "A1", "SE499", "A2")


if __name__ == "__main__":
    unittest.main()
