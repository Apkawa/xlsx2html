from xlsx2html.parser.parser import WBParser


def test_generic(fixture_file):
    p = WBParser(filepath=fixture_file("example.xlsx"))
    result = p.get_sheet()
    assert result
