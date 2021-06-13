from typing import List, Iterable, Any, Type, Optional

from openpyxl.cell import Cell
from openpyxl.formatting.formatting import ConditionalFormatting as _OPYXL_ConditionalFormatting
from openpyxl.worksheet.cell_range import MultiCellRange
from openpyxl.worksheet.worksheet import Worksheet

from xlsx2html.parser.cell import CellInfo


class ConditionalFormatting:
    _cf: _OPYXL_ConditionalFormatting
    _ws: Worksheet

    def __init__(self, cf: _OPYXL_ConditionalFormatting, ws: Worksheet):
        self._cf = cf
        self._ws = ws

    @property
    def cell_range(self) -> MultiCellRange:
        return self._cf.cells

    def get_cells(self) -> Iterable[Cell]:
        for cell_range in self._cf.cells:
            for row, col in cell_range.cells:
                yield self._ws.cell(row, col)

    def get_dataset(self) -> Iterable[Any]:
        for c in self.get_cells():
            yield c.value

    def get_rules(self):
        from xlsx2html.parser.conditional_format.rules import get_rule_class_by_type

        rules = []
        for r in self._cf.rules:
            rule_cls = get_rule_class_by_type(r.type)
            if rule_cls:
                rules.append(rule_cls(rule=r, parent=self))
        return rules

    def apply(self, cell: CellInfo) -> bool:
        pass


class ConditionalFormattingParser:
    _cf_list: List[ConditionalFormatting]
    _ws = Worksheet

    def __init__(self, ws: Worksheet):
        self._ws = ws
        self._cf_list = [ConditionalFormatting(cf, ws) for cf in ws.conditional_formatting]

    def get_formatters(self, coord: str) -> List[ConditionalFormatting]:
        return [cf for cf in self._cf_list if coord in cf.cell_range]

    def apply(self, cell: CellInfo) -> bool:
        cf_list = self.get_formatters(cell.coordinate)
        if not cf_list:
            return False
