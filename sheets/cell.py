import enum
from typing import Optional, Set, Any


class CellType(enum.Enum):
    """This enum specifies the types of values that cells can hold."""

    NONE = 0
    FORMULA = 1
    DECIMAL = 2
    STRING = 3
    ERROR = 4
    BOOLEAN = 5


class Cell:
    """Cell class."""

    def __init__(self, col: int, row: int, contents: Optional[str], val: Any,
                 val_type: CellType, sheet_name: str, pt):
        """Initialize a new cell.
        Attributes:
            col (int): The column the cell is located at.
            row (int): The row the cell is located at.

            children (Set[Cell]): The list of cells that depend on this cell.
            parents (Set[Cell]): The list of cells that this cell depends on.
        """

        self.loc = (col, row)
        self.contents = contents
        self.val = val
        self.val_type = val_type
        self.sheet_name = sheet_name
        self.pt = pt
        self.children: Set[Cell] = set()
        self.parents: Set[Cell] = set()
        self.invalid_sheet_refs: Set[str] = set()

    def make_empty(self) -> None:
        self.contents = None
        self.val = None
        self.val_type = CellType.NONE
        self.pt = None
        while self.parents:
            p = self.parents.pop()
            p.children.remove(self)
        self.invalid_sheet_refs.clear()
