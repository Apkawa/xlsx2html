from dataclasses import dataclass
from typing import Optional, Union, Any

from openpyxl.cell import Cell

from xlsx2html.compat import OPENPYXL_24
from xlsx2html.utils.cell import get_cell_id
from xlsx2html.utils.color import normalize_color
from xlsx2html.format import format_cell


@dataclass
class Border:
    style: str
    color: Optional[str] = None


@dataclass
class Borders:
    top: Optional[Border] = None
    left: Optional[Border] = None
    right: Optional[Border] = None
    bottom: Optional[Border] = None

    diagonal_up: Optional[Border] = None
    diagonal_down: Optional[Border] = None


@dataclass
class Fill:
    pattern: str
    color: Optional[str]


@dataclass
class Font:
    size: int
    color: Optional[str] = None
    italic: bool = False
    underline: bool = False
    bold: bool = False
    strike: bool = False
    overline: bool = False
    outline: bool = False
    shadow: bool = False


@dataclass
class Alignment:
    horizontal: Optional[str] = None
    vertical: Optional[str] = None
    indent: Optional[float] = None
    text_rotation: Optional[int] = None


@dataclass
class CellInfo:
    id: str
    column: int
    column_letter: str
    row: int
    coordinate: str
    value: Any
    formatted_value: str

    alignment: Alignment
    # hyperlink: Optional[Hyperlink] = None
    # format: str = ''
    colspan: Optional[int] = None
    rowspan: Optional[int] = None

    height: int = 19
    border: Optional[Borders] = None
    fill: Optional[Fill] = None
    font: Optional[Font] = None

    @classmethod
    def from_cell(
        cls, cell: Cell, f_cell: Optional[Cell], _locale: Optional[str] = None
    ) -> "CellInfo":

        if OPENPYXL_24:
            col_idx = cell.col_idx
            col_letter = cell.column
        else:
            col_idx = cell.column
            col_letter = cell.column_letter

        cell_info = CellInfo(
            id=get_cell_id(cell),
            row=cell.row,
            column=col_idx,
            column_letter=col_letter,
            coordinate=cell.coordinate,
            value=cell.value,
            formatted_value=format_cell(cell, locale=_locale, f_cell=f_cell),
            alignment=Alignment(
                horizontal=cell.alignment.horizontal,
                vertical=cell.alignment.vertical,
                indent=cell.alignment.indent or None,
            ),
        )
        # https://docs.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.alignment
        text_rotation = cell.alignment.textRotation
        if text_rotation > 90:
            text_rotation = 90 - text_rotation
        # FIXME rotation angle 91-265 not work
        text_rotation *= -1
        cell_info.alignment.text_rotation = text_rotation or None

        if cell.fill:
            cell_info.fill = Fill(
                color=normalize_color(cell.fill.fgColor), pattern=cell.fill.patternType
            )
        if cell.font:
            font = Font(size=cell.font.sz)
            if cell.font.color:
                font.color = normalize_color(cell.font.color)
            font.bold = cell.font.b
            font.italic = cell.font.i
            font.underline = cell.font.u
            font.strike = cell.font.strike
            # no supporting
            # font.overline
            font.outline = cell.font.outline
            font.shadow = cell.font.shadow
            cell_info.font = font
        cell_info.border = cls.get_border(cell)
        # TODO merged border
        return cell_info

    @staticmethod
    def get_border(cell: Cell) -> Union[Borders, None]:
        border_info = Borders()
        is_none = True
        for b_dir in ["right", "left", "top", "bottom", "diagonal"]:
            b_s = getattr(cell.border, b_dir)
            if not b_s or not b_s.style:
                continue
            b = Border(style=b_s.style)
            if b_s.color:
                b.color = normalize_color(b_s.color)
            if b_dir == "diagonal":
                if cell.border.diagonalUp:
                    border_info.diagonal_up = b
                if cell.border.diagonalDown:
                    border_info.diagonal_down = b
            else:
                setattr(border_info, b_dir, b)

            is_none = False
        if is_none:
            return None
        return border_info
