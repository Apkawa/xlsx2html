import openpyxl
import pytest

from xlsx2html.parser.conditional_format.conditional import ConditionalFormattingParser
from xlsx2html.parser.utils import get_sheet


@pytest.fixture
def fixture_ws(fixture_file):
    def _get_ws(filename, sheet=None):
        filepath = fixture_file(filename)
        wb = openpyxl.load_workbook(filepath, data_only=True)
        ws = get_sheet(wb, sheet)
        return ws

    return _get_ws


def test_conditional_parser(fixture_ws):
    ws = fixture_ws("conditional.xlsx")
    cfp = ConditionalFormattingParser(ws)
    cf_list = cfp.get_formatters("B15")
    assert len(cf_list) == 1
    cf = cf_list[0]
    assert list(cf.get_dataset()) == [9, 1, 6, 2, 10, 5, 7, 2, 1, 3, 2]
    rules = cf.get_rules()
    assert len(rules) == 1
    rule = rules[0]
    rule.check_condition()


def test_cellIs(fixture_ws):
    ws = fixture_ws("conditional.xlsx")
    conds = list(ws.conditional_formatting)
