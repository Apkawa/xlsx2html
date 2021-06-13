# https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_ST_ConditionalFormat_topic_ID0ETVFFB.html#topic_ID0ETVFFB
from xlsx2html.parser.conditional_format.rules.base import BaseRule

OPERATOR_MAP = {
    "lessThan": lambda v, t: v < t,
    "lessThanOrEqual": lambda v, t: v <= t,
    "equal": lambda v, t: v == t,
    "notEqual": lambda v, t: v != t,
    "greaterThanOrEqual": lambda v, t: v >= t,
    "greaterThan": lambda v, t: v > t,
    "between": lambda v, f, t: f >= v <= t,
    "notBetween": lambda v, f, t: not (f >= v <= t),
    "containsText": lambda v, t: t in v,
    "notContains": lambda v, t: t not in v,
    "beginsWith": lambda v, t: v.startswith(t),
    "endsWith": lambda v, t: v.endswith(t),
}


class CellIsRule(BaseRule):
    type = "cellIs"
