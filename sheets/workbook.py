from __future__ import annotations
from .error import CellError, CellErrorType, error_str, error_desc
from .sheet import Sheet
from .cell import Cell, CellType
from . import util
from . import functions
from typing import Optional, List, Dict, Tuple, Any, TextIO, Callable, \
                   Iterable, Set, Deque
import re
from lark import Transformer, Visitor, Lark, Token
from lark.reconstruct import Reconstructor
from decimal import Decimal
import json
from collections import deque
import traceback
from functools import cmp_to_key


class EvalExpressions(Transformer):
    """A Lark Transformer to handle formula parsing."""

    def __init__(self, wb):
        self.workbook = wb
        self._sheet_name = None
        self.invalid_sheet_refs = set()

    def set_sheet_name(self, sheet_name):
        self._sheet_name = sheet_name

    def expr(self, args):
        return eval(args[0])

    def number(self, args):
        return util.strip_trailing_zeroes(Decimal(args[0]))

    def parens(self, args):
        return args[0]

    def string(self, args):
        s = args[0][1:-1]
        if s in error_str:
            return CellError(error_str.get(s), error_desc.get(s))
        return s

    def bool(self, args):
        return args[0].lower() == "true"

    def error(self, args):
        return CellError(error_str.get(args[0]), error_desc.get(args[0]))

    def add_expr(self, args):
        if isinstance(args[0], CellError):
            return args[0]
        if isinstance(args[2], CellError):
            return args[2]

        args[0] = args[0] if args[0] else Decimal("0")
        args[2] = args[2] if args[2] else Decimal("0")

        # Use regex to check if the expression can be converted into a decimal.
        regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
        if type(args[0]) != Decimal:
            if type(args[0]) == str and re.match(regex, args[0].strip()):
                args[0] = Decimal(args[0].strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        # Same check for the second argument.
        if type(args[2]) != Decimal:
            if type(args[2]) == str and re.match(regex, args[2].strip()):
                args[2] = Decimal(args[2].strip())
            else:
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))

        if args[1] == "+":
            return args[0] + args[2]
        return args[0] - args[2]

    def mul_expr(self, args):
        if isinstance(args[0], CellError):
            return args[0]
        if isinstance(args[2], CellError):
            return args[2]

        args[0] = args[0] if args[0] else Decimal("0")
        args[2] = args[2] if args[2] else Decimal("0")

        # Use regex to check if the expression can be converted into a decimal.
        regex = r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$'
        if type(args[0]) != Decimal:
            if type(args[0]) == str and re.match(regex, args[0].strip()):
                args[0] = Decimal(args[0].strip())
            else:
                # Can't be parsed if it isn't a decimal and can't be converted.
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))
        # Same check for the second argument.
        if type(args[2]) != Decimal:
            if type(args[2]) == str and re.match(regex, args[2].strip()):
                args[2] = Decimal(args[2].strip())
            else:
                return CellError(CellErrorType.TYPE_ERROR,
                                 error_desc.get(CellErrorType.TYPE_ERROR))

        if args[1] == "*":
            return args[0] * args[2]
        elif args[2] == 0:
            return CellError(CellErrorType.DIVIDE_BY_ZERO,
                             error_desc.get(CellErrorType.DIVIDE_BY_ZERO))
        return args[0] / args[2]

    def unary_op(self, args):
        if isinstance(args[1], CellError):
            return args[1]

        if args[0] == "-":
            return -args[1]
        else:
            return args[1]

    def concat_expr(self, args):
        if isinstance(args[0], CellError):
            return args[0]
        if isinstance(args[1], CellError):
            return args[1]

        args[0] = str(args[0]) if args[0] else ""
        args[1] = str(args[1]) if args[1] else ""

        # Boolean to string conversion
        if isinstance(args[0], bool):
            args[0] = str(args[0]).upper()
        if isinstance(args[1], bool):
            args[1] = str(args[1]).upper()

        return args[0] + args[1]

    def compare_expr(self, args):
        if isinstance(args[0], CellError):
            return args[0]
        if isinstance(args[2], CellError):
            return args[2]

        if args[0] is None and args[2] is None:
            return args[1] in ["=", "==", ">=", "<="]
        elif args[0] is None:
            if isinstance(args[2], str):
                args[0] = ""
            elif isinstance(args[2], Decimal):
                args[0] = Decimal("0")
            elif isinstance(args[2], bool):
                args[0] = False
        elif args[2] is None:
            if isinstance(args[0], str):
                args[2] = ""
            elif isinstance(args[0], Decimal):
                args[2] = Decimal("0")
            elif isinstance(args[0], bool):
                args[2] = False

        if isinstance(args[0], str):
            args[0] = args[0].lower()
        if isinstance(args[2], str):
            args[2] = args[2].lower()

        if args[1] in ["=", "=="]:
            return args[0] == args[2] and isinstance(args[0], type(args[2]))
        elif args[1] in ["<>", "!="]:
            return args[0] != args[2] or not isinstance(args[0], type(args[2]))

        priorities = {bool: 3, str: 2, Decimal: 1}
        if not isinstance(args[0], type(args[2])):
            args[0] = priorities[type(args[0])]
            args[2] = priorities[type(args[2])]

        if args[1] == "<":
            return args[0] < args[2]
        elif args[1] == ">":
            return args[0] > args[2]
        elif args[1] == "<=":
            return args[0] <= args[2]
        elif args[1] == ">=":
            return args[0] >= args[2]

    def cell_range(self, args):
        sht_name = self._sheet_name
        start_location = end_location = None
        if len(args) == 2:
            start_location, end_location = args[0], args[1]
        else:
            sht_name, start_location, end_location = args[0], args[1], args[2]

        # Check if sheet_name exists. If so, check if the sheet name is valid
        if sht_name.lower() not in self.workbook.sheet_names_lower_to_orig:
            return CellError(CellErrorType.BAD_REFERENCE,
                             error_desc.get(CellErrorType.BAD_REFERENCE))
        sht_name = self.workbook.sheet_names_lower_to_orig[sht_name.lower()]

        # Compute area of cells to move from start_location and end_locations
        start_loc = util.quantify_cell_loc(start_location)
        end_loc = util.quantify_cell_loc(end_location)

        tl = (min(start_loc[0], end_loc[0]), min(start_loc[1], end_loc[1]))
        br = (max(start_loc[0], end_loc[0]), max(start_loc[1], end_loc[1]))

        # Return range of cells as a list of values
        vals = []
        for r in range(tl[1], br[1] + 1):
            row = []
            for c in range(tl[0], br[0] + 1):
                loc_str = util.stringify_cell_loc(c, r)
                row.append(self.workbook.get_cell_value(sht_name, loc_str))
            vals.append(row)
        return vals

    def func(self, args):
        func, inputs = args[0].upper(), args[1:]
        try:
            if func == "INDIRECT":
                return functions.function_dict[func](inputs, self._sheet_name,
                                                     self.workbook)
            else:
                return functions.function_dict[func](inputs)
        except KeyError as e:
            return CellError(CellErrorType.BAD_NAME,
                             error_desc.get(CellErrorType.BAD_NAME), e)

    def cell(self, args):
        try:
            if len(args) == 2:  # Cross-reference with another sheet
                return self.workbook.get_cell_value(args[0], args[1])
            return self.workbook.get_cell_value(self._sheet_name, args[0])
        except KeyError as e:
            self.invalid_sheet_refs.add(args[0])
            return CellError(CellErrorType.BAD_REFERENCE,
                             error_desc.get(CellErrorType.BAD_REFERENCE), e)
        except ValueError as e:
            return CellError(CellErrorType.BAD_REFERENCE,
                             error_desc.get(CellErrorType.BAD_REFERENCE), e)


class VisitRefs(Visitor):
    """A Lark Visitor that stores a list of CELLREFs found in the tree."""
    def __init__(self):
        self.refs = []
        self._sheet_name = None

    def reset_refs(self):
        self.refs = []

    def set_sheet_name(self, sheet_name):
        self._sheet_name = sheet_name

    def cell(self, tree):
        if len(tree.children) == 2:
            self.refs.append((tree.children[0].lower(),
                              tree.children[1].lower()))
        else:
            self.refs.append((self._sheet_name.lower(),
                              tree.children[0].lower()))


class RenameSheet(Visitor):
    def __init__(self):
        self._old_sheet_name = None
        self._new_sheet_name = None

    def set_old_sheet_name(self, sheet_name):
        self._old_sheet_name = sheet_name

    def set_new_sheet_name(self, sheet_name):
        self._new_sheet_name = sheet_name

    def cell(self, tree):
        sheet_match = (tree.children[0].type == "SHEET_NAME" and
                       tree.children[0] == self._old_sheet_name) or \
                      (tree.children[0].type == "QUOTED_SHEET_NAME" and
                       tree.children[0] == f"'{self._old_sheet_name}'")
        if len(tree.children) == 2 and sheet_match:
            # If a sheet’s name doesn’t start with an alphabetical character
            # or underscore, or if a sheet’s name contains spaces or any other
            # characters besides “A-Z”, “a-z”, “0-9” or the underscore “_”, it
            # must be quoted to parse correctly.
            if re.match(r"^([a-zA-Z_])([a-zA-Z0-9_]*)$", self._new_sheet_name):
                tree.children[0] = Token("SHEET_NAME", self._new_sheet_name)
            else:
                tree.children[0] = Token("QUOTED_SHEET_NAME",
                                         f"'{self._new_sheet_name}'")


class MoveFormula(Visitor):
    """A Lark Visitor that updates a cell's formula refs when it is moved."""
    def __init__(self):
        self._d_cols = 0
        self._d_rows = 0

    def set_deltas(self, d_cols, d_rows):
        self._d_cols = d_cols
        self._d_rows = d_rows

    def cell(self, tree):
        pattern = r"^(?P<abs_col>\$?)[A-Za-z]+(?P<abs_row>\$?)[1-9][0-9]*$"
        i = len(tree.children) - 1
        c, r = util.quantify_cell_loc(tree.children[i])
        new_c, new_r = c, r
        abs_col = abs_row = True

        groups = re.match(pattern, tree.children[i])
        if not groups.group("abs_col"):
            new_c += self._d_cols
            abs_col = False
        if not groups.group("abs_row"):
            new_r += self._d_rows
            abs_row = False

        if not 1 <= new_c <= util.MAX_COL or not 1 <= new_r <= util.MAX_ROW:
            tree.data = "error"
            tree.children[i] = Token("ERROR_VALUE", "#REF!")
        else:
            tree.children[i] = Token("CELLREF", util.stringify_cell_loc(
                new_c, new_r, abs_col, abs_row).lower())


class Workbook:
    """A workbook containing zero or more named spreadsheets.

    Any and all operations on a workbook that may affect calculated cell
    values should cause the workbook's contents to be updated properly.

    Attributes:
        sheet_arr (List[Sheet]): A list containing the sheet objects.
        sheet_name_to_idx (): A dictionary mapping from sheet name to their
        index in sheet_list.
    """

    def __init__(self):
        """Initialize a new empty workbook."""

        self.sheet_arr: List[Sheet] = []
        self.sheet_name_to_idx: Dict[str, int] = {}
        self.sheet_names_lower_to_orig: Dict[str, str] = {}
        self.parser = Lark.open("sheets/formulas.lark", start="formula",
                                maybe_placeholders=False)
        self.transformer = EvalExpressions(self)
        self.ref_visitor = VisitRefs()
        self.rename_visitor = RenameSheet()
        self.move_visitor = MoveFormula()
        self.orphans: Set[Cell] = set()
        self.cell_change_notif_funcs: List[Callable] = []

        self._sheet_being_deleted = None

    def num_sheets(self) -> int:
        """Return the number of spreadsheets in the workbook.

        Returns:
            The number of spreadsheets in the workbook.
        """

        return int(len(self.sheet_arr))

    def list_sheets(self) -> List[str]:
        """Return a list of the spreadsheet names in the workbook.

        The capitalization is specified at creation, and in the order that the
        sheets appear within the workbook.

        In this project, the sheet names appear in the order that the user
        created them; later, when the user is able to move and copy sheets,
        the ordering of the sheets in this function's result will also
        reflect such operations.

        A user should be able to mutate the return-value without affecting
        the workbook's internal state.

        Returns:
            The list of the spreadsheet names in the workbook.

        Raises:
            KeyError: Raises an exception.
        """

        return [sheet.name for sheet in self.sheet_arr]

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        """Add a new sheet to the workbook.

        If the sheet name is specified, it must be unique. If the sheet name is
        None, a unique sheet name is generated. "Uniqueness" is determined in a
        case-insensitive manner, but the case specified for the sheet name is
        preserved.

        The function returns a tuple with two elements:
        (0-based index of sheet in workbook, sheet name). This allows the
        function to report the sheet's name when it is auto-generated.

        If the spreadsheet name is an empty string (not None), or it is
        otherwise invalid, a ValueError is raised.

        Args:
            sheet_name: The name of the new sheet to be added.

        Returns:
            A tuple with two elements: (0-based index of sheet in workbook,
                                        sheet name)

        Raises:
            ValueError: Raises an exception.
        """

        assert (isinstance(sheet_name, str) or sheet_name is None), \
            "sheet_name arg must be of type str."

        if sheet_name is None:
            sheet_num = 1
            while ("Sheet" + str(sheet_num)).lower() in \
                    self.sheet_names_lower_to_orig:
                sheet_num += 1
            sheet_name = "Sheet" + str(sheet_num)

        else:
            valid_match = re.match(r'^(?!\s)[-.?!,:;!@#$%^&*()\w ]*(?<!\s)$',
                                   sheet_name)

            if not sheet_name or not valid_match or sheet_name.lower() \
                    in self.sheet_names_lower_to_orig:
                raise ValueError("Invalid spreadsheet name.")

        self.sheet_name_to_idx[sheet_name] = len(self.sheet_arr)
        self.sheet_names_lower_to_orig[sheet_name.lower()] = sheet_name
        self.sheet_arr.append(Sheet(sheet_name))
        self.update_orphans(sheet_name)
        return (len(self.sheet_arr) - 1, sheet_name)

    def del_sheet(self, sheet_name: str) -> None:
        """Delete the spreadsheet with the specified name.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        If the specified sheet name is not found, a KeyError is raised.

        Args:
            sheet_name: The name of the sheet to be deleted.

        Returns:
            None.

        Raises:
            KeyError: Raises an exception.
        """

        sheet_name_lower = sheet_name.lower()
        if sheet_name_lower not in self.sheet_names_lower_to_orig:
            raise KeyError("Sheet name not found.")

        sheet_name_orig = self.sheet_names_lower_to_orig[sheet_name_lower]
        idx_del = self.sheet_name_to_idx[sheet_name_orig]
        sheet = self.sheet_arr[idx_del]
        self._sheet_being_deleted = sheet_name_orig

        # For every cell in the sheet, remove references to it and re-evaluate
        # neighbors
        for loc in list(sheet.loc_to_cell.keys()):
            col, row = loc
            self.set_cell_contents(sheet_name_orig,
                                   util.stringify_cell_loc(col, row), "#REF!")
            cell = sheet.loc_to_cell.pop(loc)
            cell.make_empty()

        # Remove sheet from class attributes
        self.sheet_name_to_idx.pop(sheet_name_orig)
        self.sheet_names_lower_to_orig.pop(sheet_name_lower)
        self.sheet_arr.pop(idx_del)
        self._sheet_being_deleted = None

        # At the end, shift the indices of subsequent sheets left by 1
        for i in range(idx_del, len(self.sheet_arr)):
            self.sheet_name_to_idx[self.sheet_arr[i].name] -= 1

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:
        """Rename the specified sheet to the new sheet name.  Additionally, all
        cell formulas that referenced the original sheet name are updated to
        reference the new sheet name (using the same case as the new sheet
        name, and single-quotes iff [if and only if] necessary).

        The sheet_name match is case-insensitive; the text must match but the
        case does not have to.

        As with new_sheet(), the case of the new_sheet_name is preserved by
        the workbook.

        If the sheet_name is not found, a KeyError is raised.

        If the new_sheet_name is an empty string or is otherwise invalid, a
        ValueError is raised."""

        assert (isinstance(new_sheet_name, str) or new_sheet_name is None), \
            "sheet_name arg must be of type str."

        valid_match = re.match(r'^(?!\s)[-.?!,:;!@#$%^&*()\w ]*(?<!\s)$',
                               new_sheet_name)

        other_names = set(self.sheet_names_lower_to_orig.keys()) \
            - set([sheet_name.lower()])
        if not new_sheet_name or not valid_match or new_sheet_name.lower() \
                in other_names:
            raise ValueError("Invalid spreadsheet name.")
        elif sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError("Sheet name not found.")

        old_sheet_name = self.sheet_names_lower_to_orig[sheet_name.lower()]

        # Workbook and Sheet attribute changes
        idx = self.sheet_name_to_idx.pop(old_sheet_name)
        cur_sheet = self.sheet_arr[idx]
        cur_sheet.name = new_sheet_name
        self.sheet_name_to_idx[new_sheet_name] = idx
        self.sheet_names_lower_to_orig.pop(old_sheet_name.lower())
        self.sheet_names_lower_to_orig[new_sheet_name.lower()] = new_sheet_name

        self.rename_visitor.set_old_sheet_name(old_sheet_name)
        self.rename_visitor.set_new_sheet_name(new_sheet_name)

        for cell in cur_sheet.loc_to_cell.values():
            cell.sheet_name = new_sheet_name

            # Update children's formula of all cells using the parse tree
            if cell.pt:
                self.rename_visitor.visit(cell.pt)

                # Reconstruct formula with updated parse tree
                cell.contents = "=" + \
                    Reconstructor(self.parser).reconstruct(cell.pt)

            # Only update the contents and parse tree of affected cell's
            # immediate children
            for child in cell.children:
                if child.pt:
                    self.rename_visitor.visit(child.pt)
                    child.contents = "=" + \
                        Reconstructor(self.parser).reconstruct(child.pt)

        self.update_orphans(new_sheet_name)
        self.rename_visitor.set_old_sheet_name(None)
        self.rename_visitor.set_new_sheet_name(None)

    def move_sheet(self, sheet_name: str, index: int) -> None:
        """Move the specified sheet to the specified index in the workbook's
        ordered sequence of sheets.  The index can range from 0 to
        workbook.num_sheets() - 1.  The index is interpreted as if the
        specified sheet were removed from the list of sheets, and then
        re-inserted at the specified index.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        If the specified sheet name is not found, a KeyError is raised.

        If the index is outside the valid range, an IndexError is raised.

        Args:
            sheet_name: The name of the sheet.
            index: The index the sheet will be inserted into.
        """

        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError
        if index < 0 or index >= self.num_sheets():
            raise IndexError

        orig_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        oldindex = self.list_sheets().index(orig_name)
        self.sheet_arr.insert(index, self.sheet_arr.pop(oldindex))
        self.sheet_name_to_idx[sheet_name] = index

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        """Make a copy of the specified sheet, storing the copy at the end of
        the workbook's sequence of sheets.

        The copy's name is generated by appending "_1", "_2", ... to the
        original sheet's name (preserving the original sheet name's case),
        incrementing the number until a unique name is found. As usual,
        "uniqueness" is determined in a case-insensitive manner.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        The copy should be added to the end of the sequence of sheets in the
        workbook.  Like new_sheet(), this function returns a tuple with two
        elements:  (0-based index of copy in workbook, copy sheet name).  This
        allows the function to report the new sheet's name and index in the
        sequence of sheets.

        If the specified sheet name is not found, a KeyError is raised.

        Args:
            sheet_name: The name of the sheet.

        Returns:
            A tuple with two elements: (0-based index of sheet in workbook,
                                        copy sheet name)
        """

        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError

        # Name copy in O(n) time where n is the number of current copies.
        orig_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        copy_num = 1
        while (orig_name + "_" + str(copy_num)).lower() in \
                self.sheet_names_lower_to_orig:
            copy_num += 1

        orig_sheet = self.sheet_arr[self.sheet_name_to_idx[orig_name]]
        copy_name = orig_name + "_" + str(copy_num)

        # Create a new sheet and set the cell contents for that sheet.
        # Using set_cell_contents should handle all edge cases.
        _, _ = self.new_sheet(copy_name)
        for loc in list(orig_sheet.loc_to_cell.keys()):
            col, row = loc
            string_loc = util.stringify_cell_loc(col, row)

            orig_contents = self.get_cell_contents(sheet_name, string_loc)
            self.set_cell_contents(copy_name, string_loc, orig_contents)

        return (self.num_sheets() - 1, copy_name)

    def update_orphans(self, new_sheet_name: str):
        """Updates the cells that reference invalid sheet names. This function
        should be called whenever a sheet is created or renamed.
        """

        for cell in self.orphans:
            if new_sheet_name in cell.invalid_sheet_refs:
                col, row = cell.loc
                self.set_cell_contents(cell.sheet_name,
                                       util.stringify_cell_loc(col, row),
                                       cell.contents)

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        """Return a tuple (num-cols, num-rows) indicating the current extent of
        the specified spreadsheet.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        If the specified sheet name is not found, a KeyError is raised.

        Args:
            sheet_name: The name of the sheet.

        Returns:
            The extent of sheet_name.

        Raises:
            KeyError: Raises an exception.
        """

        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError("Sheet name not found.")
        sheet_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        cur_sheet = self.sheet_arr[self.sheet_name_to_idx[sheet_name]]
        return cur_sheet.get_extent()

    def reevaluate_refs(self, updated_cell: Cell):
        """Remove old parents and itself from their children and build
        new references.
        """

        while updated_cell.parents:
            p = updated_cell.parents.pop()
            p.children.remove(updated_cell)

        parent_refs = self.ref_visitor.refs
        self.ref_visitor.reset_refs()

        for parent_ref in parent_refs:
            try:
                # 1st wall of defense for self-reference
                c, r = updated_cell.loc
                if (updated_cell.sheet_name.lower(),
                        util.stringify_cell_loc(c, r)) in parent_refs:
                    raise RuntimeError

                cell = self.get_cell_instance(parent_ref[0], parent_ref[1])
                if not cell:
                    # If the cell reference is valid but doesn't exist yet
                    self.set_cell_contents(parent_ref[0], parent_ref[1], None)
                    cell = self.get_cell_instance(parent_ref[0], parent_ref[1])

                # 2nd wall of defense for self-reference
                if cell == updated_cell:
                    raise RuntimeError

                # Check for mypy type checking
                if cell:
                    updated_cell.parents.add(cell)
                    cell.children.add(updated_cell)
            except (KeyError, ValueError) as e:
                # Cell reference is an invalid sheet name or cell location
                updated_cell.val = CellError(
                        CellErrorType.BAD_REFERENCE,
                        error_desc[CellErrorType.BAD_REFERENCE],
                        e
                )
                updated_cell.val_type = CellType.ERROR
            except RuntimeError as e:
                # Cell references itself
                updated_cell.val = CellError(
                        CellErrorType.CIRCULAR_REFERENCE,
                        error_desc[CellErrorType.CIRCULAR_REFERENCE],
                        e
                )
                updated_cell.val_type = CellType.ERROR

    def evaluate_cell(self, updated_cell: Cell, initial_vals: Dict[Cell, Any]):
        """Evaluates the cell and its neighbors using the Lark transformer."""

        update_ordering_gen = util.topological_sort(updated_cell)
        for cell in update_ordering_gen:
            if cell not in initial_vals:
                initial_vals[cell] = cell.val

            if cell.contents and cell.contents[0] == "=":
                # Regardless of whether the transformer raises an exception
                # because of a bad reference, we should update the cell
                # updated_cell.parents.add()

                try:
                    self.transformer.set_sheet_name(cell.sheet_name)

                    # If parse tree is None, then there is nothing to
                    # evaluate or it is already an error
                    if cell.pt:
                        cell.val = self.transformer.transform(cell.pt)
                except KeyError as e:
                    cell.val = CellError(
                        CellErrorType.BAD_NAME,
                        error_desc[CellErrorType.BAD_NAME],
                        e
                    )
                except ValueError as e:
                    cell.val = CellError(
                        CellErrorType.BAD_REFERENCE,
                        error_desc[CellErrorType.BAD_REFERENCE],
                        e
                    )
                except TypeError as e:
                    print("ASDFE")
                    print(e)
                    cell.val = CellError(
                        CellErrorType.TYPE_ERROR,
                        error_desc[CellErrorType.TYPE_ERROR],
                        e
                    )
                except Exception as e:
                    print("ASDFR")
                    print(e)
                    traceback.print_stack()
                    print(cell.pt)
                    cell.val = CellError(
                        CellErrorType.TYPE_ERROR,
                        error_desc[CellErrorType.TYPE_ERROR],
                        e
                    )

                if isinstance(cell.val, CellError):
                    cell.val_type = CellType.ERROR

                # If cells reference invalid sheets, then add them to the
                # orphan list so we know which cells need to be updated
                # later when a sheet is added/renamed
                if self.transformer.invalid_sheet_refs:
                    self.orphans.add(cell)
                    for sheet_ref in self.transformer.invalid_sheet_refs:
                        cell.invalid_sheet_refs.add(sheet_ref)
                    self.transformer.invalid_sheet_refs.clear()

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        """Set the contents of the specified cell on the specified sheet.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to. Additionally, the cell location can be
        specified in any case.

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        A cell may be set to "empty" by specifying a contents of None.

        Leading and trailing whitespace are removed from the contents before
        storing them in the cell. Storing a zero-length string "" (or a
        string composed entirely of whitespace) is equivalent to setting the
        cell contents to None.

        If the cell contents appear to be a formula, and the formula is
        invalid for some reason, this method does not raise an exception;
        rather, the cell's value will be a CellError object indicating the
        naure of the issue.

        Args:
            sheet_name: The name of the sheet to be updated.
            location: The cell location to be updated.
            contents: The content to be inserted into the specified location.

        Returns:
            None.

        Raises:
            KeyError: Raises an exception.
        """

        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError("Sheet name not found.")
        sheet_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        cur_sheet = self.sheet_arr[self.sheet_name_to_idx[sheet_name]]

        # Get initial value before wiping it in update_cell for notifications
        cur_cell = cur_sheet.cell_from_loc(location)
        initial_vals = {cur_cell: cur_cell.val} if cur_cell else {}
        updated_cell = cur_sheet.update_cell(location, contents, self.parser)
        if not initial_vals:
            initial_vals[updated_cell] = None

        # Get list of references (parents) from updated_cell's parse tree
        self.ref_visitor.set_sheet_name(sheet_name)
        if updated_cell.pt:
            self.ref_visitor.visit(updated_cell.pt)

        # Regenerate new parents and children relationships
        self.reevaluate_refs(updated_cell)

        # For self-reference
        self.ref_visitor.reset_refs()

        # Clear invalid sheet refs
        updated_cell.invalid_sheet_refs.clear()

        # Cycle detection and finding neighbors
        cycle_exists, scc_arr = util.detect_cycle(updated_cell)

        # If a cycle exists, then propagate errors to all of its neighbors.
        # Otherwise, we evaluate cells in order of the topological sort
        if cycle_exists:
            for scc in scc_arr:
                for cell in scc:
                    if cell not in initial_vals:
                        initial_vals[cell] = cell.val

                    cell.val = CellError(
                        CellErrorType.CIRCULAR_REFERENCE,
                        error_desc[CellErrorType.CIRCULAR_REFERENCE]
                    )
                    cell.val_type = CellType.ERROR
        else:
            self.evaluate_cell(updated_cell, initial_vals)

        final_vals = {cell: cell.val for cell in initial_vals}
        changed_cells = [
            (cell.sheet_name,
             util.stringify_cell_loc(cell.loc[0], cell.loc[1]))
            for cell in final_vals if initial_vals[cell] != final_vals[cell]
            and cell.sheet_name != self._sheet_being_deleted
        ]
        self.call_notify_functions(changed_cells)

    def move_copy_cells_helper(self, sheet_name: str, start_location: str,
                               end_location: str, to_location: str,
                               copying: bool, to_sheet: Optional[str] = None) \
            -> None:
        """Helper to move or copy cells."""

        # Check if sheet_name exists. If so, check if the sheet name is valid
        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError
        sheet_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        src_sheet_idx = self.sheet_name_to_idx[sheet_name]
        src_sheet = self.sheet_arr[src_sheet_idx]

        # Check if to_sheet exists. If so, check if the sheet name is valid
        # If not, we are using current sheet
        if to_sheet is not None:
            if to_sheet.lower() not in self.sheet_names_lower_to_orig:
                raise KeyError
            tgt_sheet_name = self.sheet_names_lower_to_orig[to_sheet.lower()]
        else:
            tgt_sheet_name = sheet_name

        # Compute area of cells to move from start_location and end_locations
        start_loc = util.quantify_cell_loc(start_location)
        end_loc = util.quantify_cell_loc(end_location)

        src_tl = (min(start_loc[0], end_loc[0]), min(start_loc[1], end_loc[1]))
        src_br = (max(start_loc[0], end_loc[0]), max(start_loc[1], end_loc[1]))
        w = src_br[0] - src_tl[0]
        h = src_br[1] - src_tl[1]

        # Check if target area extends outside [0, ZZZZ9999]
        to_tl = util.quantify_cell_loc(to_location)
        to_br = util.stringify_cell_loc(to_tl[0] + w, to_tl[1] + h)
        # Call quantify_cell_loc just to check
        util.quantify_cell_loc(to_br)

        d_cols = to_tl[0] - src_tl[0]
        d_rows = to_tl[1] - src_tl[1]

        # Store source area cells (dict?), changing their info
        # change sheet name if moving to another sheet
        cells: Deque[Tuple] = deque()
        for c in range(src_tl[0], src_br[0] + 1):
            for r in range(src_tl[1], src_br[1] + 1):
                loc_str = util.stringify_cell_loc(c, r)
                cur_cell = self.get_cell_instance(sheet_name, loc_str)
                if cur_cell:
                    cells.append((c, r, cur_cell.contents, cur_cell.pt))
                    if not copying:
                        # Make cells in source area empty
                        self.set_cell_contents(sheet_name, loc_str, None)

                        # Delete cell if it has no children
                        if not cur_cell.children:
                            src_sheet.loc_to_cell.pop((c, r), None)

        # Iterate over area starting at to-location locations and fill in
        while cells:
            old_c, old_r, contents, pt = cells.popleft()
            new_loc = util.stringify_cell_loc(old_c + d_cols, old_r + d_rows)
            self.move_visitor.set_deltas(d_cols, d_rows)

            # Update contents formula based on how far the move is
            if pt:
                self.move_visitor.visit(pt)
                # Reconstruct formula with updated parse tree
                contents = "=" + \
                    Reconstructor(self.parser).reconstruct(pt)

            self.set_cell_contents(tgt_sheet_name, new_loc, contents)

    def move_cells(self, sheet_name: str, start_location: str,
                   end_location: str, to_location: str,
                   to_sheet: Optional[str] = None) -> None:
        """Move cells from one location to another, possibly moving them to
        another sheet.

        All formulas in the area being moved will also have
        all relative and mixed cell-references updated by the relative
        distance each formula is being copied.

        Cells in the source area (that are not also in the target area) will
        become empty due to the move operation.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be moved.  The to_location specifies the
        top-left corner of the target area to move the cells to.

        Both corners are included in the area being moved; for example,
        copying cells A1-A3 to B1 would be done by passing
        start_location="A1", end_location="A3", and to_location="B1".

        The start_location value does not necessarily have to be the top left
        corner of the area to move, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to move.

        This function works correctly even when the destination area overlaps
        the source area.

        The sheet name matches are case-insensitive; the text must match but
        the case does not have to.

        If to_sheet is None then the cells are being moved to another
        location within the source sheet.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.

        If the target area would extend outside the valid area of the
        spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        no changes are made to the spreadsheet.

        If a formula being moved contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal
        in the formula."""

        self.move_copy_cells_helper(sheet_name, start_location, end_location,
                                    to_location, False, to_sheet)

    def copy_cells(self, sheet_name: str, start_location: str,
                   end_location: str, to_location: str,
                   to_sheet: Optional[str] = None) -> None:
        """Copy cells from one location to another, possibly copying them to
        another sheet.

        All formulas in the area being copied will also have
        all relative and mixed cell-references updated by the relative
        distance each formula is being copied.

        Cells in the source area (that are not also in the target area) are
        left unchanged by the copy operation.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be copied.  The to_location specifies the
        top-left corner of the target area to copy the cells to.

        Both corners are included in the area being copied; for example,
        copying cells A1-A3 to B1 would be done by passing
        start_location="A1", end_location="A3", and to_location="B1".

        The start_location value does not necessarily have to be the top left
        corner of the area to copy, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to copy.

        This function works correctly even when the destination area overlaps
        the source area.

        The sheet name matches are case-insensitive; the text must match but
        the case does not have to.

        If to_sheet is None then the cells are being copied to another
        location within the source sheet.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.

        If the target area would extend outside the valid area of the
        spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        no changes are made to the spreadsheet.

        If a formula being copied contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal in the formula"""

        self.move_copy_cells_helper(sheet_name, start_location, end_location,
                                    to_location, True, to_sheet)

    def compare(self, vals: List[Any], c: int) -> None:
        """A custom sort function for sorting regions."""

        rankings = {
            type(None): 1,
            type(CellError(CellErrorType.TYPE_ERROR, None)): 2
        }

        def cmp(x, y):
            k = abs(c)
            if isinstance(x[k], type(y[k])):
                if x[k] == y[k]:
                    return 0
                elif x[k] < y[k]:
                    return -1
                return 1
            else:
                xrank = rankings.get(type(x[k]), 3)
                yrank = rankings.get(type(y[k]), 3)
                if xrank == yrank:
                    return 0
                elif xrank < yrank:
                    return -1
                return 1

        vals.sort(key=cmp_to_key(cmp), reverse=c < 0)

    def sort_region(self, sheet_name: str, start_location: str,
                    end_location: str, sort_cols: List[int]) -> None:
        """Sort the specified region of a spreadsheet with a stable sort, using
        the specified columns for the comparison.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be sorted.  Both corners are included in the
        area being sorted; for example, sorting the region including cells B3
        to J12 would be done by specifying start_location="B3" and
        end_location="J12".

        The start_location value does not necessarily have to be the top left
        corner of the area to sort, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to sort.

        The sort_cols argument specifies one or more columns to sort on.  Each
        element in the list is the one-based index of a column in the region,
        with 1 being the leftmost column in the region.  A column's index in
        this list may be positive to sort in ascending order, or negative to
        sort in descending order.  For example, to sort the region B3..J12 on
        the first two columns, but with the second column in descending order,
        one would specify sort_cols=[1, -2].

        The sorting implementation is a stable sort:  if two rows compare as
        "equal" based on the sorting columns, then they will appear in the
        final result in the same order as they are at the start.

        If multiple columns are specified, the behavior is as one would
        expect:  the rows are ordered on the first column indicated in
        sort_cols; when multiple rows have the same value for the first
        column, they are then ordered on the second column indicated in
        sort_cols; and so forth.

        No column may be specified twice in sort_cols; e.g. [1, 2, 1] or
        [2, -2] are both invalid specifications.

        The sort_cols list may not be empty.  No index may be 0, or refer
        beyond the right side of the region to be sorted.

        If the specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.
        If the sort_cols list is invalid in any way, a ValueError is raised."""

        # Check if sheet_name exists. If so, check if the sheet name is valid
        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError
        sht_name = self.sheet_names_lower_to_orig[sheet_name.lower()]

        # Compute area of cells to move from start_location and end_locations
        start_loc = util.quantify_cell_loc(start_location)
        end_loc = util.quantify_cell_loc(end_location)

        tl = (min(start_loc[0], end_loc[0]), min(start_loc[1], end_loc[1]))
        br = (max(start_loc[0], end_loc[0]), max(start_loc[1], end_loc[1]))

        # Check for invalid sort_cols list
        sort_cols_seen = set()
        for c in sort_cols:
            if c == 0 or abs(c) in sort_cols_seen or \
                    abs(c) > br[0] - tl[0] + 1:
                raise ValueError
            sort_cols_seen.add(abs(c))

        sort_cols_norm = [tl[0] + abs(x) - 1 for x in sort_cols]
        sort_list = []
        for r in range(tl[1], br[1] + 1):
            sort_row = [r]
            for c in sort_cols_norm:
                loc_str = util.stringify_cell_loc(c, r)
                sort_row.append(self.get_cell_value(sht_name, loc_str))
            sort_list.append(tuple(sort_row))

        for c in range(len(sort_cols), 0, -1):
            # sort_list.sort(key=itemgetter(abs(c)), reverse=c<0)
            # sort_list.sort(key=lambda x: x[abs(c)], reverse=c<0)
            # sort_list.sort(key=lambda x: self.compare(x, c), reverse=c<0)
            self.compare(sort_list, c * (1 if sort_cols[c-1] > 0 else -1))

        # Move rows to where they should belong
        row_to_contents: Dict[int, List[Tuple[Any, Any]]] = {}
        for i in range(len(sort_list)):
            cur_row = sort_list[i][0]
            tgt_row = tl[1] + i

            # If the row is already in the correct spot, skip below
            if cur_row == tgt_row:
                continue

            # Store the row we're about to replace
            if tgt_row not in row_to_contents:
                row_to_contents[tgt_row] = []
                for c in range(tl[0], br[0] + 1):
                    loc = util.stringify_cell_loc(c, tgt_row)
                    cell = self.get_cell_instance(sht_name, loc)
                    row_to_contents[tgt_row].append((
                        cell.contents if cell else None,
                        cell.pt if cell else None
                    ))

            if cur_row in row_to_contents:
                contents = row_to_contents[cur_row]
                for c in range(tl[0], br[0] + 1):
                    loc = util.stringify_cell_loc(c, tgt_row)
                    self.move_visitor.set_deltas(0, tgt_row - cur_row)

                    # Update contents formula based on how far the move is
                    pt = contents[c-tl[0]][1]
                    new_contents = contents[c-tl[0]][0]
                    if pt:
                        self.move_visitor.visit(pt)
                        new_contents = "=" + \
                            Reconstructor(self.parser).reconstruct(pt)

                    self.set_cell_contents(sht_name, loc, new_contents)
            else:
                cur_tl = util.stringify_cell_loc(tl[0], cur_row)
                cur_br = util.stringify_cell_loc(br[0], cur_row)
                tgt_loc = util.stringify_cell_loc(tl[0], tgt_row)
                self.move_cells(sht_name, cur_tl, cur_br, tgt_loc)

    def get_cell_instance(self, sheet_name: str, location: str) \
            -> Optional[Cell]:
        """Return the cell instance from the specified sheet.

        Args:
            sheet_name: The name of the sheet.
            location: The cell location.

        Returns:
            Cell instance at location.

        Raises:
            KeyError: Raises an exception.
        """

        sheet_name = sheet_name.translate({39: None})
        if sheet_name.lower() not in self.sheet_names_lower_to_orig:
            raise KeyError("Sheet name not found.")
        sheet_name = self.sheet_names_lower_to_orig[sheet_name.lower()]
        cur_sheet = self.sheet_arr[self.sheet_name_to_idx[sheet_name]]
        return cur_sheet.cell_from_loc(location)

    def get_cell_contents(self, sheet_name: str, location: str) \
            -> Optional[str]:
        """Return the contents of the specified cell on the specified sheet.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.  Additionally, the cell location can be
        specified in any case.

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        Any string returned by this function will not have leading or
        trailing whitespace, as this whitespace will have been stripped off
        by the set_cell_contents() function.

        This method will never return a zero-length string; instead, empty
        cells are indicated by a value of None.

        Args:
            sheet_name: The name of the sheet.
            location: The cell location.

        Returns:
            Contents at location.

        Raises:
            KeyError: Raises an exception.
        """

        cell = self.get_cell_instance(sheet_name, location)
        return cell.contents if cell else None

    def get_cell_value(self, sheet_name: str, location: str) -> Any:
        """Return the evaluated value of the specified cell on the specified
        sheet.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.  Additionally, the cell location can be
        specified in any case.

        If the specified sheet name is not found, a KeyError is raised.
        If the cell location is invalid, a ValueError is raised.

        The value of empty cells is None.  Non-empty cells may contain a
        value of str, decimal.Decimal, or CellError.

        Decimal values will not have trailing zeros to the right of any
        decimal place, and will not include a decimal place if the value is a
        whole number.  For example, this function would not return
        Decimal('1.000'); rather it would return Decimal('1').

        Args:
            sheet_name: The name of the sheet.
            location: The cell location.

        Returns:
            The value at location.

        Raises
            KeyError: Raises an exception.
        """

        cell = self.get_cell_instance(sheet_name, location)
        return cell.val if cell else None

    @staticmethod
    def load_workbook(fp: TextIO) -> Workbook:
        """This is a static method (not an instance method) to load a workbook
        from a text file or file-like object in JSON format, and return the
        new Workbook instance.

        Note that the _caller_ of this function is expected to have opened the
        file; this function merely reads the file.

        If the contents of the input cannot be parsed by the Python json
        module then a json.JSONDecodeError should be raised by the method.
        (Just let the json module's exceptions propagate through.) Similarly,
        if an IO read error occurs (unlikely but possible), let any raised
        exception propagate through.

        If any expected value in the input JSON is missing (e.g. a sheet
        object doesn't have the "cell-contents" key), raise a KeyError with
        a suitably descriptive message.

        If any expected value in the input JSON is not of the proper type
        (e.g. an object instead of a list, or a number instead of a string),
        raise a TypeError with a suitably descriptive message.

        Args:
            fp: An opened file object.

        Returns:
            A Workbook object with the name and cell contents specified by the
            json object referenced by fp.
        """

        data = json.load(fp)
        wb = Workbook()
        for sheet in data["sheets"]:
            sheet_name = sheet["name"]
            wb.new_sheet(sheet_name)
            for loc in sheet["cell-contents"]:
                contents = sheet["cell-contents"][loc]
                wb.set_cell_contents(sheet_name, loc, contents)
        return wb

    def save_workbook(self, fp: TextIO) -> None:
        """Instance method (not a static/class method) to save a workbook to a
        text file or file-like object in JSON format.  Note that the _caller_
        of this function is expected to have opened the file; this function
        merely writes the file.

        If an IO write error occurs (unlikely but possible), let any raised
        exception propagate through.

        Args:
            fp: An opened file object.
        """

        data: Dict[str, List[Dict[str, Any]]] = {}
        data["sheets"] = []
        for sheet in self.sheet_arr:
            sheet_dict: Dict[str, Any] = {}
            sheet_dict["name"] = sheet.name
            contents_dict = {}
            for loc in list(sheet.loc_to_cell.keys()):
                col, row = loc
                string_loc = util.stringify_cell_loc(col, row)
                contents = self.get_cell_contents(sheet.name, string_loc)
                contents_dict[string_loc.upper()] = contents
            sheet_dict["cell-contents"] = contents_dict
            data["sheets"].append(sheet_dict)

        json.dump(data, fp, indent=4)

    def notify_cells_changed(self,
                             notify_function: Callable[[Workbook,
                                                        Iterable[Tuple[str,
                                                                       str]]],
                                                       None]) -> None:
        """Request that all changes to cell values in the workbook are reported
        to the specified notify_function.  The values passed to the notify
        function are the workbook, and an iterable of 2-tuples of strings,
        of the form ([sheet name], [cell location]).  The notify_function is
        expected not to return any value; any return-value will be ignored.

        Multiple notification functions may be registered on the workbook;
        functions will be called in the order that they are registered.

        A given notification function may be registered more than once; it
        will receive each notification as many times as it was registered.

        If the notify_function raises an exception while handling a
        notification, this will not affect workbook calculation updates or
        calls to other notification functions.

        A notification function is expected to not mutate the workbook or
        iterable that it is passed to it.  If a notification function violates
        this requirement, the behavior is undefined."""

        self.cell_change_notif_funcs.append(notify_function)

    def call_notify_functions(self, changed_cells: Iterable[Tuple[str, str]]):
        """Calls each registered notification function upon changed cells.
        Catches and ignores any exception that may be raised."""

        for func in self.cell_change_notif_funcs:
            try:
                func(self, changed_cells)
            except Exception as e:
                print(e)


class ProxyRow:
    def __init__(self, row, sort_col):
        self.orig_row = row
        self.sort_col = sort_col
