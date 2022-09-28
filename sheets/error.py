import enum
from typing import Optional
from functools import total_ordering


@total_ordering
class CellErrorType(enum.Enum):
    """This enum specifies the error types that spreadsheet cells can hold."""

    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6

    def __lt__(self, other):
        if self.__class__ is other.__class__:
            return self.value < other.value
        return NotImplemented


error_str = {
    "#ERROR!": CellErrorType.PARSE_ERROR,
    "#CIRCREF!": CellErrorType.CIRCULAR_REFERENCE,
    "#REF!": CellErrorType.BAD_REFERENCE,
    "#NAME?": CellErrorType.BAD_NAME,
    "#VALUE!": CellErrorType.TYPE_ERROR,
    "#DIV/0!": CellErrorType.DIVIDE_BY_ZERO
}

error_desc = {
    CellErrorType.PARSE_ERROR: "Parse error",
    CellErrorType.CIRCULAR_REFERENCE: "Circular reference",
    CellErrorType.BAD_REFERENCE: "Invalid sheet or cell reference",
    CellErrorType.BAD_NAME: "Bad name",
    CellErrorType.TYPE_ERROR: "Type error",
    CellErrorType.DIVIDE_BY_ZERO: "Divided by zero"
}


@total_ordering
class CellError:
    """This class represents an error value from user input, cell parsing, or
    evaluation.
    """

    def __init__(self, error_type: CellErrorType, detail: Optional[str],
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception

    def get_type(self) -> CellErrorType:
        """The category of the cell error."""
        return self._error_type

    def get_detail(self) -> str:
        """More detail about the cell error."""
        return self._detail if self._detail else ""

    def get_exception(self) -> Optional[Exception]:
        """If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.
        """
        return self._exception

    def __str__(self) -> str:
        return f'ERROR[{self._error_type}, "{self._detail}"]'

    def __repr__(self) -> str:
        return self.__str__()

    def __eq__(self, other):
        if isinstance(other, CellError):
            return self._error_type == other._error_type
        return False

    def __ne__(self, other):
        return not self.__eq__(other)

    def __lt__(self, other):
        if isinstance(other, CellError):
            return self._error_type < other._error_type
        return False
