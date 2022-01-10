import openpyxl

from xlsx2html.parser.parser import XLSXParser


def test_generic(fixture_file):
    p = XLSXParser(wb=openpyxl.load_workbook(fixture_file("example.xlsx"), data_only=True))
    result = p.get_sheet()
    assert result
