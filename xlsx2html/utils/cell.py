import re

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


def parse_cell_location(cell_location):
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
