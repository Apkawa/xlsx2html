from urllib.parse import urlsplit

from openpyxl.formula.tokenizer import Token, Tokenizer

from xlsx2html.utils.cell import parse_cell_location
from xlsx2html.utils.html import SafeHTML, escape_html

SAFE_URL_SCHEMES = {"file", "http", "https", "mailto", "tel", "sms", "geo"}


class HyperlinkType:
    __slots__ = ["location", "target", "title"]

    def __init__(self, location=None, target=None, title=None):
        self.location = location
        self.target = target
        self.title = title

    def __bool__(self):
        return bool(self.location or self.target)


def resolve_cell(worksheet, coord):
    if "!" in coord:
        sheet_name, coord = coord.split("!", 1)
        worksheet = worksheet.parent[sheet_name.lstrip("$")]
    return worksheet[coord]


def resolve_hyperlink_formula(cell, f_cell):
    if not f_cell or f_cell.data_type != "f" or not f_cell.value.startswith("="):
        return None
    tokens = Tokenizer(f_cell.value).items
    if not tokens:
        return None
    hyperlink = HyperlinkType(title=cell.value)
    func_token = tokens[0]
    if func_token.type == Token.FUNC and func_token.value == "HYPERLINK(":
        target_token = tokens[1]
        if target_token.type == Token.OPERAND:
            target = target_token.value
            if target_token.subtype == Token.TEXT:
                hyperlink.target = target[1:-1]
            elif target_token.subtype == Token.RANGE:
                hyperlink.target = resolve_cell(cell.parent, target).value

        if hyperlink:
            return hyperlink

    return None


def is_safe_url(href: str):
    # reject dangerous schemes like javascript:alert(1)
    try:
        url = urlsplit(href)
        # we allow relative URLs, which is a bit questionable in context of Excel files,
        # but it can't be used to do anything super nefarious
        return url.scheme == "" or url.scheme.lower() in SAFE_URL_SCHEMES
    except ValueError:
        return False


def format_hyperlink(value, cell, f_cell=None):
    hyperlink = HyperlinkType(title=value)

    if cell.hyperlink:
        hyperlink.location = cell.hyperlink.location
        hyperlink.target = cell.hyperlink.target

    # Parse function

    if not hyperlink:
        hyperlink = resolve_hyperlink_formula(cell, f_cell)
        if not hyperlink:
            return value

    if hyperlink.location is not None:
        href = "{}#{}".format(hyperlink.target or "", hyperlink.location)
    else:
        href = hyperlink.target

    if href is None:
        return value
    href = str(href)

    # Maybe link to cell
    if href.startswith("#"):
        location_info = parse_cell_location(href)
        if location_info:
            href = "#{}.{}".format(
                location_info["sheet_name"] or cell.parent.title, location_info["coord"]
            )

    if not is_safe_url(href):
        return value

    return SafeHTML(f'<a href="{escape_html(href)}">{escape_html(value)}</a>')
