from typing import Optional, Tuple, Dict, List
from heapq import heappush, heappop
from collections import Counter
from .cell import Cell
from . import util


class Sheet:
    """A spreadsheet class.
    Any and all operations on a workbook that may affect calculated cell
    values should cause the workbook's contents to be updated properly.
    Attributes:
        name (str): Name of the sheet. Capitalization is considered.
        extent Tuple[int][int]: A tuple (num-cols, num-rows) indicating the
        current extent of the specified spreadsheet.
        loc_to_cell (dict): A dictionary mapping from cell location to tuple
        containing the cell's content, displayed value, and the type of the
        displayed value.
    """

    def __init__(self, name: str):
        """Initialize a new empty spreadsheet."""

        self.name = name
        self.loc_to_cell: Dict[Tuple[int, int], Cell] = {}
        self.col_counter: Dict[int, int] = Counter()
        self.row_counter: Dict[int, int] = Counter()
        self.col_max_heap: List[int] = []
        self.row_max_heap: List[int] = []

    def cell_from_loc(self, loc: str) -> Optional[Cell]:
        """Given cell location, returns cell instance."""

        return self.loc_to_cell.get(util.quantify_cell_loc(loc))

    def update_cell(self, loc: str, contents: Optional[str], parser) -> Cell:
        """Update a cell at a specified location.
        The user specifies a cell and the given value is entered into that
        cell. The previous value is overwritten and the spreadsheet is updated
        accordingly. The extent will also be updated
        Args:
            loc (str): The location being updated.
            contents Optional[str]: The value to be inserted.
        Returns:
            Cell: the cell instance that was created or changed.
        """

        loc_col, loc_row = util.quantify_cell_loc(loc)
        clean_loc = (loc_col, loc_row)
        clean_contents, val, val_type, pt = util.parse_contents(contents,
                                                                parser)

        created_new_cell = False
        return_cell = None
        if clean_contents is not None:
            if clean_loc in self.loc_to_cell:
                cur_cell = self.loc_to_cell[clean_loc]
                cur_cell.contents = clean_contents
                cur_cell.val = val
                cur_cell.val_type = val_type
                cur_cell.pt = pt
                return_cell = cur_cell
            else:
                return_cell = Cell(loc_col, loc_row, clean_contents, val,
                                   val_type, self.name, pt)
                self.loc_to_cell[clean_loc] = return_cell
                created_new_cell = True
        else:
            # No contents, so we empty the cell (return for topo-sort)
            if clean_loc in self.loc_to_cell:
                return_cell = self.loc_to_cell[clean_loc]
                return_cell.make_empty()

                # Delete cell from all records
                if len(return_cell.children) == 0:
                    self.loc_to_cell.pop(clean_loc)

                    # Update the col/row counters and heaps
                    self.col_counter[loc_col] -= 1
                    self.row_counter[loc_row] -= 1
                    if self.col_counter[loc_col] == 0:
                        self.col_counter.pop(loc_col)
                    if self.row_counter[loc_row] == 0:
                        self.row_counter.pop(loc_row)
                    while self.col_max_heap and \
                            -self.col_max_heap[0] not in self.col_counter:
                        heappop(self.col_max_heap)
                    while self.row_max_heap and \
                            -self.row_max_heap[0] not in self.row_counter:
                        heappop(self.row_max_heap)
            else:
                return_cell = Cell(loc_col, loc_row, clean_contents, val,
                                   val_type, self.name, pt)
                self.loc_to_cell[clean_loc] = return_cell
                created_new_cell = True

        # If a new cell has been created, update the col/row counters and heaps
        if created_new_cell:
            if loc_col not in self.col_counter:
                heappush(self.col_max_heap, -loc_col)
            if loc_row not in self.row_counter:
                heappush(self.row_max_heap, -loc_row)
            self.col_counter[loc_col] += 1
            self.row_counter[loc_row] += 1

        return return_cell

    def get_extent(self) -> Tuple[int, int]:
        """Calculates the extent of the sheet.
        Returns:
            The extent of the sheet as a tuple of two ints.
        """

        if not self.col_max_heap or not self.row_max_heap:
            return (0, 0)
        return (-self.col_max_heap[0], -self.row_max_heap[0])
