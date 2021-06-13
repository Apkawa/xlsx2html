from typing import List, Union, Any

from openpyxl.formatting.rule import Rule as _OPYXL_Rule

from xlsx2html.parser.cell import CellInfo
from xlsx2html.parser.conditional_format.conditional import ConditionalFormatting


class BaseRule:
    type: Union[List[str]] = ""

    _rule: _OPYXL_Rule
    _parent: ConditionalFormatting

    def __init__(self, rule: _OPYXL_Rule, parent: ConditionalFormatting):
        self._rule = rule
        self._parent = parent

    def check_condition(self, value: Any):
        raise NotImplementedError()

    def apply(self, cell: CellInfo) -> bool:
        raise NotImplementedError()
