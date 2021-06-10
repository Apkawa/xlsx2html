import re

from openpyxl.cell import Cell
from openpyxl.utils import get_column_letter, column_index_from_string

CELL_LOCATION_RE = re.compile(
    r"""
    ^
        \#
        (?:
            (?P<sheet_name>\w+)[.!]
            |
        )
        (?P<coord>[A-Za-z]+[\d]+)
    $
    """,
    re.VERBOSE,
)


def parse_cell_location(cell_location: str):
    """
    >>> parse_cell_location("#Sheet1.C1")
    {'sheet_name': 'Sheet1', 'coord': 'C1'}
    >>> parse_cell_location("#Sheet1!C1")
    {'sheet_name': 'Sheet1', 'coord': 'C1'}
    >>> parse_cell_location("#C1")
    {'sheet_name': None, 'coord': 'C1'}
    """

    m = CELL_LOCATION_RE.match(cell_location)
    if not m:
        return None

    return m.groupdict()


def col_index_to_letter(col_index: int) -> str:
    return get_column_letter(col_index)


def letter_to_col_index(col_letter: str) -> int:
    return column_index_from_string(col_letter)


def get_cell_id(cell: Cell) -> str:
    return "{}!{}".format(cell.parent.title, cell.coordinate)
