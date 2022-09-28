from sheets.workbook import Workbook
import unittest


CHANGED_CELLS = []


def on_cells_changed(workbook, changed_cells):
    """This function gets called when cells change in the workbook that the
    function was registered on.  The changed_cells argument is an iterable
    of tuples; each tuple is of the form (sheet_name, cell_location).
    """
    if changed_cells:
        CHANGED_CELLS.append(changed_cells)
    # print(f'Cell(s) changed:  {changed_cells}')


class TestCellUpdateNotifcations(unittest.TestCase):
    def test_basic_update_notifs(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed)

        wb.set_cell_contents(name, "a1", "1")
        wb.set_cell_contents(name, "a2", "2")
        wb.set_cell_contents(name, "a3", "=a1 + a2")

        self.assertEqual(CHANGED_CELLS, [
            [(name, "a1")], [(name, "a2")], [(name, "a3")]
        ])

        wb.set_cell_contents(name, "a1", "10")

        self.assertEqual(CHANGED_CELLS[-1], [(name, "a1"), (name, "a3")])

        wb.set_cell_contents(name, "a2", "20")

        self.assertEqual(CHANGED_CELLS[-1], [(name, "a2"), (name, "a3")])

        wb.set_cell_contents(name, "a3", None)

        self.assertEqual(CHANGED_CELLS[-1], [(name, "a3")])

        wb.set_cell_contents(name, "a3", "=#REF!")

        self.assertEqual(CHANGED_CELLS[-1], [("Sheet1", "a3")])

    def test_notifs_for_value_changes(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed)

        wb.set_cell_contents(name, "a1", "1")

        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        wb.set_cell_contents(name, "a1", "=1+1-1")

        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        wb.set_cell_contents(name, "a1", "=1+1-1")

        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

    def test_robust_notifs(self):
        def on_cells_changed_error(workbook, changed_cells):
            raise ValueError

        CHANGED_CELLS.clear()
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed_error)

        try:
            wb.set_cell_contents(name, "a1", "1")
        except Exception:
            self.fail("Notification function raising an error breaks program.")

    def test_many_notif_funcs(self):
        def func1(workbook, changed_cells):
            CHANGED_CELLS.append(1)

        def func2(workbook, changed_cells):
            CHANGED_CELLS.append(2)

        def func3(workbook, changed_cells):
            CHANGED_CELLS.append(3)

        CHANGED_CELLS.clear()
        wb = Workbook()
        _, name = wb.new_sheet()
        wb.notify_cells_changed(func2)
        wb.notify_cells_changed(func1)
        wb.notify_cells_changed(func3)
        wb.notify_cells_changed(func1)

        wb.set_cell_contents(name, "a1", "1")

        self.assertEqual(CHANGED_CELLS, [2, 1, 3, 1])

        wb.set_cell_contents(name, "a1", "2")

        self.assertEqual(CHANGED_CELLS, [2, 1, 3, 1, 2, 1, 3, 1])

    def test_add_sheets_notifs(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        wb.notify_cells_changed(on_cells_changed)
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=Sheet2!a1")

        # None -> #REF!
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        _, name2 = wb.new_sheet()

        # #REF! -> None
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")], [(name, "a1")]])

    def test_del_sheets_notifs(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        wb.notify_cells_changed(on_cells_changed)
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=Sheet2!a1")
        wb.del_sheet(name2)

        # None -> #REF!
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        wb.del_sheet(name)

        # #REF! -> None
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

    def test_copy_sheets_notifs(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        wb.notify_cells_changed(on_cells_changed)
        _, name = wb.new_sheet()

        wb.set_cell_contents(name, "a1", "=Sheet1_1!b1")

        # None -> #REF!
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        wb.copy_sheet(name)

        # #REF! -> None
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")], [(name, "a1")]])

    def test_rename_sheets_notifs(self):
        CHANGED_CELLS.clear()
        wb = Workbook()
        _, name = wb.new_sheet()
        _, name2 = wb.new_sheet()
        wb.notify_cells_changed(on_cells_changed)

        wb.set_cell_contents(name, "a1", "=jesus!a1")

        # None -> #REF!
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")]])

        wb.rename_sheet(name2, "jesus")

        # #REF! -> None
        self.assertEqual(CHANGED_CELLS, [[(name, "a1")], [(name, "a1")]])
