from typing import Union

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

SheetNameType = Union[int, str, None]


def get_sheet(wb: Workbook, sheet: SheetNameType = None) -> Worksheet:
    ws = wb.active
    if sheet is not None:
        try:
            ws = wb.get_sheet_by_name(sheet)
        except KeyError:
            ws = wb.worksheets[sheet]
    return ws
