from dataclasses import dataclass
from typing import Union, Optional

from openpyxl.cell import Cell
from openpyxl.formula.tokenizer import Tokenizer, Token

from openpyxl.worksheet.worksheet import Worksheet

from xlsx2html.utils.cell import parse_cell_location


@dataclass
class Hyperlink:
    title: str
    location: Optional[str] = None
    target: Optional[str] = None
    href: Optional[str] = None

    def __bool__(self) -> bool:
        return bool(self.location or self.target)


def resolve_cell(worksheet: Worksheet, coord: str) -> Cell:
    if "!" in coord:
        sheet_name, coord = coord.split("!", 1)
        worksheet = worksheet.parent[sheet_name.lstrip("$")]
    return worksheet[coord]


def resolve_hyperlink_formula(cell: Cell, f_cell: Optional[Cell] = None) -> Union[Hyperlink, None]:
    if not f_cell or f_cell.data_type != "f" or not f_cell.value.startswith("="):
        return None
    tokens = Tokenizer(f_cell.value).items
    if not tokens:
        return None
    hyperlink = Hyperlink(title=cell.value)
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


def get_hyperlink(value: str, cell: Cell, f_cell: Optional[Cell] = None) -> Union[Hyperlink, None]:
    hyperlink = Hyperlink(title=value)

    if cell.hyperlink:
        hyperlink.location = cell.hyperlink.location
        hyperlink.target = cell.hyperlink.target

    if not hyperlink:
        _h = resolve_hyperlink_formula(cell, f_cell)
        if _h:
            hyperlink = _h

    if not hyperlink:
        return None

    if hyperlink.location is not None:
        href = "{}#{}".format(hyperlink.target or "", hyperlink.location)
    else:
        href = hyperlink.target or ""

    # Maybe link to cell
    if href.startswith("#"):
        location_info = parse_cell_location(href)
        if location_info:
            href = "#{}.{}".format(
                location_info["sheet_name"] or cell.parent.title, location_info["coord"]
            )
    hyperlink.href = href
    return hyperlink


def format_hyperlink(value: str, cell: Cell, f_cell: Optional[Cell] = None) -> str:
    hyperlink = get_hyperlink(value, cell, f_cell)

    if not hyperlink:
        return value

    return '<a href="{href}">{value}</a>'.format(href=hyperlink.href, value=hyperlink.title)
