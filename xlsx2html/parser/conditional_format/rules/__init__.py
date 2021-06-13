from typing import Optional, Type
from .base import BaseRule
from .cell_is import CellIsRule


def get_rule_class_by_type(rule_type: str) -> Optional[Type[BaseRule]]:

    for c in BaseRule.__subclasses__():
        _t = c.type
        if isinstance(_t, str):
            _t = [_t]
        if rule_type in _t:
            return c
    return None
