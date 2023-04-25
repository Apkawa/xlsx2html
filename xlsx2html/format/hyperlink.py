from openpyxl.formula.tokenizer import Tokenizer, Token

from xlsx2html.utils.cell import parse_cell_location


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

    # Maybe link to cell
    if href and href.startswith("#"):
        location_info = parse_cell_location(href)
        if location_info:
            href = "#{}.{}".format(
                location_info["sheet_name"] or cell.parent.title, location_info["coord"]
            )

    return '<a href="{href}">{value}</a>'.format(href=href, value=value)
