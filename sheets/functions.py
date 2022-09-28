from .error import CellError, CellErrorType
import sheets
import re
from decimal import Decimal
from typing import List, Any
import sys


error_desc = {
    CellErrorType.PARSE_ERROR: "Parse error",
    CellErrorType.CIRCULAR_REFERENCE: "Circular reference",
    CellErrorType.BAD_REFERENCE: "Invalid sheet or cell reference",
    CellErrorType.BAD_NAME: "Bad name",
    CellErrorType.TYPE_ERROR: "Type error",
    CellErrorType.DIVIDE_BY_ZERO: "Divided by zero"
}


def bool_and(args: List[Any]) -> Any:
    """Evaluates to TRUE if all arguments are TRUE.
    This function requires one or more arguments.
    All arguments are converted to Boolean values.
    """

    if len(args) == 0:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    res = True
    for a in args:
        if isinstance(a, list):
            return CellError(CellErrorType.TYPE_ERROR,
                             error_desc.get(CellErrorType.TYPE_ERROR))
        res = res and bool(a)
    return res


def bool_or(args: List[Any]) -> Any:
    """Evaluates to TRUE if any argument is TRUE.
    This function requires one or more arguments.
    All arguments are converted to Boolean values.
    """

    if len(args) == 0:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    res = False
    for a in args:
        if isinstance(a, list):
            return CellError(CellErrorType.TYPE_ERROR,
                             error_desc.get(CellErrorType.TYPE_ERROR))
        res = res or bool(a)
    return res


def bool_not(args: List[Any]) -> Any:
    """Evaluates to TRUE if the input is FALSE, and FALSE if the input is TRUE.
    This function requires exactly one argument.
    The argument is converted to a Boolean value.
    """

    if len(args) != 1 or isinstance(args[0], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    return not bool(args[0])


def bool_xor(args: List[Any]) -> Any:
    """Evaluates to TRUE if an odd number of inputs is TRUE, and FALSE
    otherwise. This function requires one or more arguments.
    All arguments are converted to Boolean values.
    """

    if len(args) == 0:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    bool_sum = 0
    for a in args:
        if isinstance(a, list):
            return CellError(CellErrorType.TYPE_ERROR,
                             error_desc.get(CellErrorType.TYPE_ERROR))
        bool_sum += bool(a)
    return bool_sum % 2 == 1


def string_match_exact(args: List[Any]) -> Any:
    """Evaluates to TRUE if the two string arguments match exactly,
    or FALSE otherwise. This test is case-sensitive. This function requires
    exactly two arguments. The arguments are converted to string values.
    """

    if len(args) != 2 or isinstance(args[0], list) or \
            isinstance(args[1], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    args[0] = args[0] if args[0] is not None else ""
    args[1] = args[1] if args[0] is not None else ""
    return str(args[0]) == str(args[1])


def conditional_if(args: List[Any]) -> Any:
    """Evaluates to value1 if cond is TRUE. If cond is false and value2
    is specified, the function evaluates to value2. If cond is false and value2
    is not specified, the function evalutes to the FALSE Boolean value.
    This function requires 2 or 3 arguments. The first argument is converted to
    a Boolean value; the other arguments are not converted.
    """

    if len(args) < 2 or len(args) > 3 or isinstance(args[0], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    if bool(args[0]):
        return args[1] if args[1] is not None else Decimal("0")
    else:
        if len(args) == 3:
            return args[2] if args[2] is not None else Decimal("0")
        return False


def conditional_iferror(args: List[Any]) -> Any:
    """Evaluates to value1 if value1 is not an error. If value1 is an error and
    value2 is specified, the function evaluates to value2. If value1 is an
    error and value2 is not specified, the function evaluates to an empty
    string ““. This function requires 1 or 2 arguments.
    """

    if len(args) < 1 or len(args) > 2 or isinstance(args[0], list) or \
            isinstance(args[1], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    if args[0] in error_desc:
        if len(args) == 1:
            return ""
        return args[1] if args[1] is not None else Decimal("0")
    return False


def conditional_choose(args: List[Any]) -> Any:
    """Uses the first value to index into the remaining arguments: if index is
    1 then value1 is returned; if index is 2 then value2 is returned, etc.
    This function requires at least 2 arguments.
    The first argument is converted to a number. If index is not an integer,
    or is 0 or less, or is beyond the end of the value list, then a TYPE_ERROR
    is produced.
    """

    if len(args) < 2 or isinstance(args[0], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    index = args[0]
    if not isinstance(index, Decimal) or index != int(index) or index < 1 \
            or index >= len(args):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    return args[int(index)] if args[int(index)] is not None else Decimal("0")


def informational_isblank(args: List[Any]) -> Any:
    """Evaluates to TRUE if its input is an empty-cell value, or FALSE
    otherwise. This function always takes exactly one argument.

    Note specifically that ISBLANK("") evaluates to FALSE,
    as do ISBLANK(FALSE) and ISBLANK(0).
    """

    if len(args) != 1 or isinstance(args[0], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    return args[0] is None


def informational_iserror(args: List[Any]) -> Any:
    """Evaluates to TRUE if its input is an error, or FALSE otherwise.
    This function always takes exactly one argument.
    """

    if len(args) != 1 or isinstance(args[0], list):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    return isinstance(args[0], CellError)


def informational_version(args: List[Any]) -> Any:
    """Takes no arguments, and returns the version of your sheets library as a
    string.
    """

    if len(args) > 0:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    return sheets.version


def indirect(args: List[Any], sheet_name: str, workbook: Any) -> Any:
    """Parses its string argument as a cell-reference (possibly including a
    sheet name and absolute/relative modifiers), and returns the value of the
    specified cell. This function always takes exactly one argument. The
    argument is converted to a string.

    If the argument string cannot be parsed as a cell reference for any reason,
    or if the argument can be parsed as a cell reference, but the
    cell-reference is invalid for some reason (other than creating a circular
    reference in the workbook), this function returns a BAD_REFERENCE error.
    """

    if len(args) != 1:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    sht_ptrn = r"[A-Za-z_][A-Za-z0-9_]*"
    quoted_sht_ptrn = r"\'[^']*\'"
    cellref_ptrn = r"[\$]?[A-Za-z]+[\$]?[1-9][0-9]*"
    groups = re.match(fr"^(?:(?P<sheet_name>{sht_ptrn}|{quoted_sht_ptrn}) \
            \!)?(?P<cell_ref>{cellref_ptrn})$", args[0])

    if not groups or not groups.group("cell_ref"):
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))
    sheet_name = groups.group("sheet_name") or sheet_name
    return workbook.get_cell_value(sheet_name, groups.group("cell_ref"))


def flatten(arr):
    """Recursively flattens a list."""

    if isinstance(arr, list):
        return [v for a in arr for v in flatten(a)]
    return [arr]


def math_min(args: List[Any]) -> Any:
    """Returns the minimum value over the set of inputs. Arguments may include
    cell-range references as well as normal expressions; values from the
    cell-range are also considered by the function. All non-empty inputs are
    converted to numbers; if any input cannot be converted to a number then the
    function returns a TYPE_ERROR. Only non-empty cells should be considered;
    empty cells should be ignored. If the function’s inputs only include empty
    cells then the function’s result is 0. This function requires at least 1
    argument.
    """

    if len(args) < 1:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    # Flatten args list first (could have cell ranges)
    args_flat = flatten(args)

    min_val = sys.maxsize
    regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
    for a in args_flat:
        if a is None:
            continue
        if not isinstance(a, Decimal):
            # Use regex to check if expression can be converted into a decimal.
            if isinstance(a, str) and re.match(regex, args[0].strip()):
                a = Decimal(a.strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        min_val = min(min_val, a)

    if min_val == sys.maxsize:
        return Decimal(0)
    return Decimal(min_val)


def math_max(args: List[Any]) -> Any:
    """Returns the maximum value over the set of inputs. Arguments may include
    cell-range references as well as normal expressions; values from the
    cell-range are also considered by the function. All non-empty inputs are
    converted to numbers; if any input cannot be converted to a number then the
    function returns a TYPE_ERROR. Only non-empty cells should be considered;
    empty cells should be ignored. If the function’s inputs only include empty
    cells then the function’s result is 0. This function requires at least 1
    argument.
    """

    if len(args) < 1:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    # Flatten args list first (could have cell ranges)
    args_flat = flatten(args)

    max_val = -sys.maxsize - 1
    regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
    for a in args_flat:
        if a is None:
            continue
        if not isinstance(a, Decimal):
            # Use regex to check if expression can be converted into a decimal.
            if isinstance(a, str) and re.match(regex, args[0].strip()):
                a = Decimal(a.strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        max_val = max(max_val, a)

    if max_val == -sys.maxsize - 1:
        return Decimal(0)
    return Decimal(max_val)


def math_sum(args: List[Any]) -> Any:
    """Returns the sum of all inputs. Arguments may include cell-range
    references as well as normal expressions; values from the cell-range are
    added into the sum. All non-empty inputs are converted to numbers; if any
    input cannot be converted to a number then the function returns a
    TYPE_ERROR. If the function’s inputs only include empty cells then the
    function’s result is 0. This function requires at least 1 argument.
    """

    if len(args) < 1:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    # Flatten args list first (could have cell ranges)
    args_flat = flatten(args)

    val_sum = 0
    regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
    for a in args_flat:
        if a is None:
            continue
        if not isinstance(a, Decimal):
            # Use regex to check if expression can be converted into a decimal.
            if isinstance(a, str) and re.match(regex, args[0].strip()):
                a = Decimal(a.strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        val_sum += a

    return Decimal(val_sum)


def math_avg(args: List[Any]) -> Any:
    """Returns the average of all inputs. Arguments may include cell-range
    references as well as normal expressions; values from the cell-range are
    added into the sum. All non-empty inputs are converted to numbers; if an
    input cannot be converted to a number then the function returns a
    TYPE_ERROR. If the function’s inputs only include empty cells then the
    function returns a DIVIDE_BY_ZERO error. This function requires at least 1
    argument.
    """

    if len(args) < 1:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    # Flatten args list first (could have cell ranges)
    args_flat = flatten(args)

    res = num_nonempty = 0
    regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
    for a in args_flat:
        if a is None:
            continue
        num_nonempty += 1
        if not isinstance(a, Decimal):
            # Use regex to check if expression can be converted into a decimal.
            if isinstance(a, str) and re.match(regex, args[0].strip()):
                a = Decimal(a.strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        res += a

    if num_nonempty == 0:
        return CellError(CellErrorType.DIVIDE_BY_ZERO,
                         error_desc.get(CellErrorType.DIVIDE_BY_ZERO))
    return Decimal(res / num_nonempty)


def hlookup(args: List[Any]) -> Any:
    """Searches horizontally through a range of cells. The function searches
    through the first (i.e. topmost) row in range, looking for the first
    column that contains key in the search row (exact match, both type and
    value). If such a column is found, the cell in the index-th row of the
    found column. The index is 1-based; an index of 1 refers to the search row.
    If no column is found, the function returns a TYPE_ERROR.
    """

    if len(args) != 3:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    key = args[0]
    values = args[1]
    index = args[2]

    for row in range(len(values)):
        for col in range(len(values[0])):
            if type(values[row][col]) == type(key) and values[row][col] == key:
                return values[int(index) - 1][col]

    return CellError(CellErrorType.TYPE_ERROR,
                     error_desc.get(CellErrorType.TYPE_ERROR))


def vlookup(args: List[Any]) -> Any:
    """Searches vertically through a range of cells. The function searches
    through the first (i.e. leftmost) column in range, looking for the first
    row that contains key in the search column (exact match, both type and
    value). If such a column is found, the cell in the index-th column of the
    found row. The index is 1-based; an index of 1 refers to the search column.
    If no row is found, the function returns a TYPE_ERROR.
    """

    if len(args) != 3:
        return CellError(CellErrorType.TYPE_ERROR,
                         error_desc.get(CellErrorType.TYPE_ERROR))

    key = args[0]
    values = args[1]
    index = args[2]

    for col in range(len(values[0])):
        for row in range(len(values)):
            if type(values[row][col]) == type(key) and values[row][col] == key:
                return values[row][int(index) - 1]

    return CellError(CellErrorType.TYPE_ERROR,
                     error_desc.get(CellErrorType.TYPE_ERROR))


function_dict = {
    "AND": bool_and,
    "OR": bool_or,
    "NOT": bool_not,
    "XOR": bool_xor,
    "EXACT": string_match_exact,
    "IF": conditional_if,
    "IFERROR": conditional_iferror,
    "CHOOSE": conditional_choose,
    "ISBLANK": informational_isblank,
    "ISERROR": informational_iserror,
    "VERSION": informational_version,
    "INDIRECT": indirect,
    "MIN": math_min,
    "MAX": math_max,
    "SUM": math_sum,
    "AVERAGE": math_avg,
    "HLOOKUP": hlookup,
    "VLOOKUP": vlookup
}
