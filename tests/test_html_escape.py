import datetime
import random
import re
from html import escape as html_escape
from html.parser import HTMLParser
from io import BytesIO, StringIO

import openpyxl
import pytest
from openpyxl.utils.escape import unescape

from xlsx2html import xlsx2html
from xlsx2html.core import render_data_to_html, worksheet_to_data


class ParsedHTML(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tags = []
        self.text = []

    def handle_starttag(self, tag, attrs):
        self.tags.append((tag, dict(attrs)))

    def handle_startendtag(self, tag, attrs):
        self.tags.append((tag, dict(attrs)))

    def handle_data(self, data):
        self.text.append(data)

    def elements(self, tag):
        return [attrs for element, attrs in self.tags if element == tag]


def render_worksheet(ws, formula_ws=None, default_cell_border="none"):
    data = worksheet_to_data(ws, fs=formula_ws, default_cell_border=default_cell_border)
    return render_data_to_html(data, lambda _data, _html: None, lambda _row, _i: None)


def parse(html):
    parser = ParsedHTML()
    parser.feed(html)
    return parser


def render_workbook(wb):
    source = BytesIO()
    wb.save(source)
    source_bytes = source.getvalue()
    output = StringIO()
    xlsx2html(BytesIO(source_bytes), output)
    return source_bytes, output.getvalue()


def cell_fragments(html):
    return re.findall(r"<td [^>]*>(.*?)</td>", html, flags=re.DOTALL)


def test_cell_text_is_escaped_once():
    value = "<b>& \" '</b> &lt;script&gt;"
    ws = openpyxl.Workbook().active
    ws["A1"] = value

    html = render_worksheet(ws)
    parsed = parse(html)

    assert not parsed.elements("b")
    assert value in "".join(parsed.text)
    assert "&amp;lt;script&amp;gt;" in html


def test_xlsx2html_escapes_workbook_inputs():
    title = 'x" onmouseover="alert(1)'
    target = 'https://example.test/" onmouseover="alert(1)'
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = title
    ws["A1"] = "<img src=x onerror=alert(1)>"
    ws["A2"] = 42
    ws["A2"].number_format = '0"<script>alert(1)</script>"'
    ws["A3"] = "click"
    ws["A3"].hyperlink = target

    _, html = render_workbook(wb)
    parsed = parse(html)
    cells = parsed.elements("td")
    assert not parsed.elements("img")
    assert not parsed.elements("script")
    assert cells[0]["id"] == f"{title}!A1"
    assert set(cells[0]) == {"id", "style"}
    assert parsed.elements("a") == [{"href": target}]


def test_text_cells_are_escaped_once():
    targeted = [
        "&",
        "&amp;",
        "&amp;amp;",
        "&lt;script&gt;",
        "&copy &copy; &apos; &nbsp;",
        "&#60;script&#62;",
        "&#x3c;script&#x3e;",
        "&#0; &#x0; &#99999999;",
        "&quot; &#39; &#x27;",
        "<>&\"'",
        "Fish & chips",
        "&&amp;;",
        "&notanentity;",
        "</td><script>alert(1)</script>",
        '<img src=x onerror="alert(1)">',
        '" onmouseover="alert(1)',
        "_x003C_script_x003E_",
        "_x0026_amp_x003B_",
        "_x005F_x003C_ _x000D_ _x000A_ _x0022_ _x0027_ _xZZZZ_",
        "line1\n&amp;\nline2",
        "é &eacute; Ω &#937; 😀 &#x1F600;",
        "&Tab;&NewLine;",
        "&amp;<tag>\"'" * 20,
        "&" * 100,
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    for row, value in enumerate(targeted, 1):
        ws.cell(row, 1, value)

    source, rendered_html = render_workbook(wb)
    loaded_ws = openpyxl.load_workbook(BytesIO(source), data_only=True).active
    expected = [
        html_escape(unescape(loaded_ws.cell(row, 1).value), quote=True)
        for row in range(1, loaded_ws.max_row + 1)
    ]

    assert cell_fragments(rendered_html) == expected


def test_hyperlink_fragments_escaped_once():
    payloads = [
        "&",
        "&amp;",
        "&amp;amp;",
        "&quot;",
        "&#34;",
        "&#x22;",
        '" onmouseover="alert(1)',
        "<script>alert(1)</script>",
        "é &eacute; Ω &#937; 😀",
    ]
    wb = openpyxl.Workbook()
    ws = wb.active
    expected = []
    for row, payload in enumerate(payloads, 1):
        label = "label:" + payload
        target = "https://example.test/path?q=" + payload
        ws.cell(row, 1, label)
        ws.cell(row, 1).hyperlink = target
        expected.append(
            '<a href="{}">{}</a>'.format(
                html_escape(target, quote=True), html_escape(label, quote=True)
            )
        )

    _, rendered_html = render_workbook(wb)

    assert cell_fragments(rendered_html) == expected


def test_nested_safe_markup_is_not_escaped_again():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = -42
    ws["A1"].number_format = '[RED]0"<img src=x onerror=alert(1)>"'
    ws["A1"].hyperlink = "https://example.test/?a=1&amp;b=2"

    _, rendered_html = render_workbook(wb)

    assert cell_fragments(rendered_html) == [(
        '<a href="https://example.test/?a=1&amp;amp;b=2">'
        '<span style="color: RED">'
        "-42&quot;&lt;img src=x onerror=alert(1)&gt;&quot;"
        "</span></a>"
    )]


def test_rich_text_runs_escaped_once():
    rich_text = pytest.importorskip("openpyxl.cell.rich_text")
    InlineFont = pytest.importorskip("openpyxl.cell.text").InlineFont
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = rich_text.CellRichText(
        "prefix &amp; ",
        rich_text.TextBlock(InlineFont(b=True), "<script>alert(1)</script>"),
        " &lt;suffix&gt;",
    )

    _, rendered_html = render_workbook(wb)

    assert cell_fragments(rendered_html) == [
        "prefix &amp;amp; &lt;script&gt;alert(1)&lt;/script&gt; &amp;lt;suffix&amp;gt;"
    ]


def test_empty_cell_uses_non_breaking_space():
    ws = openpyxl.Workbook().active
    ws["A1"] = ""

    html = render_worksheet(ws)

    assert "&nbsp;" in html
    assert "&amp;amp;" not in html, "no double escaping happened"


@pytest.mark.parametrize(
    "value, number_format",
    [
        (42, '0"<img src=x onerror=alert(1)>"'),
        (datetime.date(2020, 1, 2), 'yyyy"<img src=x onerror=alert(1)>"'),
        (datetime.datetime(2020, 1, 2, 3, 4), 'yyyy-mm-dd"<img src=x onerror=alert(1)>"'),
        (datetime.time(3, 4), 'hh:mm"<img src=x onerror=alert(1)>"'),
        (datetime.timedelta(hours=3, minutes=4), '[h]:mm"<img src=x onerror=alert(1)>"'),
    ],
)
def test_custom_number_format_is_escaped(value, number_format):
    ws = openpyxl.Workbook().active
    ws["A1"] = value
    ws["A1"].number_format = number_format

    parsed = parse(render_worksheet(ws))

    assert not parsed.elements("img")
    assert "<img src=x onerror=alert(1)>" in "".join(parsed.text)


def test_colored_number_escapes_content_but_preserves_span():
    ws = openpyxl.Workbook().active
    ws["A1"] = -42
    ws["A1"].number_format = '[RED]0"<img src=x onerror=alert(1)>"'

    parsed = parse(render_worksheet(ws))

    assert not parsed.elements("img")
    assert parsed.elements("span") == [{"style": "color: RED"}]
    assert "<img src=x onerror=alert(1)>" in "".join(parsed.text)


def test_sheet_title_escaped():
    title = 'x" onmouseover="alert(1)'
    ws = openpyxl.Workbook().active
    ws.title = title
    ws["A1"] = "value"

    parsed = parse(render_worksheet(ws))
    attrs = parsed.elements("td")[0]

    assert attrs["id"] == f"{title}!A1"
    assert set(attrs) == {"id", "style"}


def test_hyperlink_target_escaped():
    target = 'https://example.test/" onmouseover="alert(1)'
    ws = openpyxl.Workbook().active
    ws["A1"] = "click"
    ws["A1"].hyperlink = target

    parsed = parse(render_worksheet(ws))

    assert parsed.elements("a") == [{"href": target}]


def test_hyperlink_label_escaped():
    ws = openpyxl.Workbook().active
    ws["A1"] = "<img src=x onerror=alert(1)>"
    ws["A1"].hyperlink = "https://example.test"

    parsed = parse(render_worksheet(ws))

    assert not parsed.elements("img")
    assert parsed.elements("a") == [{"href": "https://example.test"}]
    assert "<img src=x onerror=alert(1)>" in "".join(parsed.text)


def test_hyperlink_preserves_safe_colored_number_markup():
    ws = openpyxl.Workbook().active
    ws["A1"] = -42
    ws["A1"].number_format = "[RED]0"
    ws["A1"].hyperlink = "https://example.test"

    parsed = parse(render_worksheet(ws))

    assert parsed.elements("a") == [{"href": "https://example.test"}]
    assert parsed.elements("span") == [{"style": "color: RED"}]


def test_internal_hyperlink_escapes_sheet_title():
    title = 'x" onmouseover="alert(1)'
    ws = openpyxl.Workbook().active
    ws.title = title
    ws["A1"] = "click"
    ws["A1"].hyperlink = "#A1"

    parsed = parse(render_worksheet(ws))

    assert parsed.elements("a") == [{"href": f"#{title}.A1"}]


@pytest.mark.parametrize(
    "target",
    [
        "javascript:alert(1)",
        " javascript:alert(1)",
        "\tjavascript:alert(1)",
        "JaVaScRiPt:alert(1)",
        "java\nscript:alert(1)",
        "data:text/html,<script>alert(1)</script>",
        "vbscript:msgbox(1)",
    ],
)
def test_unsafe_hyperlink_scheme_is_not_rendered(target):
    ws = openpyxl.Workbook().active
    ws["A1"] = "click"
    ws["A1"].hyperlink = target

    parsed = parse(render_worksheet(ws))

    assert not parsed.elements("a")
    assert "click" in parsed.text


@pytest.mark.parametrize(
    "target",
    [
        "https://example.test/path?a=1&b=2",
        "mailto:user@example.test",
        "tel:+123456",
        "ftp://example.test/file",
        "file:///tmp/file",
        "/relative/path",
    ],
)
def test_safe_hyperlink_scheme_is_preserved(target):
    ws = openpyxl.Workbook().active
    ws["A1"] = "click"
    ws["A1"].hyperlink = target

    parsed = parse(render_worksheet(ws))

    assert parsed.elements("a") == [{"href": target}]


def test_unsafe_formula_hyperlink_is_not_rendered():
    ws = openpyxl.Workbook().active
    ws["A1"] = "click"

    formula_wb = openpyxl.Workbook()
    formula_ws = formula_wb.active
    formula_ws["A1"] = '=HYPERLINK("javascript:alert(1)","click")'

    parsed = parse(render_worksheet(ws, formula_ws=formula_ws))

    assert not parsed.elements("a")
    assert "click" in parsed.text


def test_custom_style_attribute_escaped():
    value = 'none" onmouseover="alert(1)'
    ws = openpyxl.Workbook().active
    ws["A1"] = "value"

    parsed = parse(render_worksheet(ws, default_cell_border=value))
    attrs = parsed.elements("td")[0]

    assert set(attrs) == {"id", "style"}
    assert value in attrs["style"]


def test_renderer_escapes_untrusted_formatted_value():
    ws = openpyxl.Workbook().active
    ws["A1"] = "safe"
    data = worksheet_to_data(ws)
    data["rows"][0][0]["formatted_value"] = "<script>alert(1)</script>"

    html = render_data_to_html(data, lambda _data, _html: None, lambda _row, _i: None)
    parsed = parse(html)

    assert not parsed.elements("script")
    assert "<script>alert(1)</script>" in "".join(parsed.text)
