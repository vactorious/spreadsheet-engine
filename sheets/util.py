from .error import CellError, CellErrorType, error_str
from .cell import Cell, CellType
import re
import decimal
from functools import reduce
from collections import deque, defaultdict
from typing import Tuple, Any, Optional, List, Set, Dict, Generator


MAX_COL = 475254
MAX_ROW = 9999


def strip_trailing_zeroes(d: decimal.Decimal) -> decimal.Decimal:
    """Given a decimal, returns a decimal without trailing zeroes to the right
    of the decimal point.
    """

    normalized = d.normalize()
    _, _, exponent = normalized.as_tuple()
    if exponent <= 0:
        return normalized
    return normalized.quantize(1)


def quantify_cell_loc(location: str) -> Tuple[int, int]:
    """Returns a tuple representing the cell location with number of columns
    and number of rows.
    """

    location = location.lower()
    location = location.replace("$", "")
    spl = re.findall(r"[^\W\d_]+|\d+", location)
    if len(spl) != 2 or not spl[0].isalpha() or not spl[1].isnumeric():
        raise ValueError("Invalid cell location.")
    col_str, row_str = spl
    col = reduce(lambda a, x: a * 26 + ord(x) - 64, col_str.upper(), 0)
    row = int(row_str)
    if col > MAX_COL or row > MAX_ROW or col < 1 or row < 1:
        raise ValueError("Invalid cell location.")
    return (col, row)


def stringify_cell_loc(col: int, row: int,
                       abs_col: bool = False, abs_row: bool = False) -> str:
    """Returns the cell location as a string ([Letter][Row]) such as "A1"."""

    if col > MAX_COL or row > MAX_ROW or col < 1 or row < 1:
        raise ValueError("Invalid cell location.")
    res = ""
    if abs_col:
        res += "$"
    col_str = ""
    while col > 0:
        col, remainder = divmod(col - 1, 26)
        col_str = chr(65 + remainder) + col_str
    res += col_str.lower()
    if abs_row:
        res += "$"
    res += str(row)
    return res


def parse_contents(contents: Optional[str], parser) \
        -> Tuple[Optional[str], Any, CellType, Optional[Any]]:
    """Returns the cleaned contents, value, its type, and the parse tree (if
    applicable) as a tuple given a cell's contents. If the contents cannot be
    parsed, the appropriate cell error type is stored as the value.
    """

    if not isinstance(contents, (str, type(None))):
        raise TypeError
    clean_contents, val_type, pt = contents, CellType.NONE, None
    val: Any = None
    if not contents:
        clean_contents = None
    else:
        # Strip leading whitespace
        clean_contents = contents.lstrip()
        if len(clean_contents) == 0:
            clean_contents = None
        else:
            if clean_contents[0] == "'":
                val_type = CellType.STRING
                val = clean_contents[1:]
            elif clean_contents[0] == "=":
                if len(clean_contents[1:]) == 0:
                    # Nothing after equal sign
                    val_type = CellType.ERROR
                    val = CellError(CellErrorType.PARSE_ERROR,
                                    "Invalid formula.", None)
                else:
                    try:
                        pt = parser.parse(clean_contents)
                        val_type = CellType.FORMULA
                    except Exception as e:
                        val_type = CellType.ERROR
                        val = CellError(CellErrorType.PARSE_ERROR,
                                        "Invalid formula.", e)
            else:
                clean_contents = contents.strip()
                if len(clean_contents) == 0:
                    clean_contents = None
                elif clean_contents.upper() in error_str:
                    clean_contents = clean_contents.upper()
                    val_type = CellType.ERROR
                    val = CellError(error_str[clean_contents], "Error", None)
                else:
                    is_dec = re.match(r'^[+-]?((\d+(\.\d*)?)|(\.\d+))$',
                                      clean_contents)
                    if is_dec:
                        val_type = CellType.DECIMAL
                        val = strip_trailing_zeroes(decimal.Decimal(
                                clean_contents))
                    elif clean_contents.lower() in ["true", "false"]:
                        val_type = CellType.BOOLEAN
                        val = clean_contents.lower() == "true"
                    else:
                        val_type = CellType.STRING
                        val = clean_contents

    return (clean_contents, val, val_type, pt)


def detect_cycle(cell: Cell) -> Tuple[bool, List[List[Cell]]]:
    """Returns whether the cell is actively or indirectly in a cycle, along
    with the strongly connected components in the cell dependency graph
    grouped together using Kosaraju's algorithm. Only returns those that
    are relevant to the updated cell (CIRCREF and REF errors).
    """

    # Build the inverse adjacency graph
    inverse_graph = {}
    queue = deque([cell])
    while queue:
        node = queue.popleft()
        if node in inverse_graph:
            continue
        inverse_graph[node] = node.children
        queue.extend(node.children)

    # Build regular DAG from inverse (a -> b = a depends on b)
    graph: defaultdict = defaultdict(list,
                                     {x: [] for x in inverse_graph.keys()})
    for i in inverse_graph:
        for j in inverse_graph[i]:
            graph[j].append(i)

    def get_scc(start_node: Cell, visited: Set[Cell]):
        """Traverses through the strongly connected component the start node
        is in and returns it as a list.
        """
        nonlocal inverse_graph
        scc, stack = [], [start_node]
        visited.add(start_node)
        while stack:
            node = stack.pop()
            scc.append(node)
            visited.add(node)
            for neighbor in inverse_graph[node]:
                if neighbor not in visited:
                    stack.append(neighbor)
        return scc

    # 1st passthrough: compute the order of cells by finishing time
    res: List[Cell] = []
    visited_status: Dict[Cell, int] = defaultdict(int)
    for start_node in graph:
        stack = [start_node]
        while stack:
            node = stack.pop()
            if node not in visited_status:
                stack.append(node)
                visited_status[node] = -1
                for neighbor in graph[node]:
                    if neighbor not in visited_status and neighbor in graph:
                        stack.append(neighbor)
            else:
                if visited_status[node] == -1:
                    visited_status[node] = 1
                    res.append(node)

    # 2nd passthrough to traverse strongly connected components
    scc_arr = []
    visited: Set[Cell] = set()
    while res:
        start_node = res.pop()
        if start_node not in visited:
            scc_arr.append(get_scc(start_node, visited))

    return any(map(lambda x: len(x) > 1, scc_arr)), scc_arr


def topological_sort(updated_cell: Cell) -> Generator[Cell, None, None]:
    """Returns the topological ordering of cells that need to be updated given
    an initial cell that is changed.
    """

    visited = set()
    stack: List[Cell] = []
    order, recursion = [], [updated_cell]

    while recursion:
        cur_cell = recursion.pop()
        if cur_cell not in visited:
            visited.add(cur_cell)
            recursion += list(cur_cell.children)
            while stack and cur_cell not in stack[-1].children:
                order.append(stack.pop())
            stack.append(cur_cell)

    for cell in stack:
        yield cell
    for cell in reversed(order):
        yield cell
