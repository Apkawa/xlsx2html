import io

from xlsx2html.parser.parser import WBParser
from xlsx2html.render.html import HtmlRenderer


def xlsx2html(
    filepath,
    output=None,
    locale="en",
    sheet=None,
    parse_formula=False,
    # DEPRECATED
    append_headers=None,
    append_lineno=None,
    default_cell_border="none",
):
    parser = WBParser(filepath=filepath, parse_formula=parse_formula, locale=locale)
    result = parser.get_sheet(sheet)

    render = HtmlRenderer(default_border_style={"style": default_cell_border})
    html = render.render(result)

    if not output:
        output = io.StringIO()
    if isinstance(output, str):
        output = open(output, "w")
    output.write(html)
    output.flush()
    return output
