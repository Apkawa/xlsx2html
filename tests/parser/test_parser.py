from xlsx2html.parser.parser import XLSXParser


def test_generic(fixture_file):
    p = XLSXParser(filepath=fixture_file("example.xlsx"))
    result = p.get_sheet()
    assert result
