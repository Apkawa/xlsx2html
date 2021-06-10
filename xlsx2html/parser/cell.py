from dataclasses import dataclass
from typing import Optional, Union, Any

from openpyxl.cell import Cell

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


@dataclass
class CellInfo:
    id: str
    column: int
    column_letter: str
    row: int
    coordinate: str
    value: Any
    formatted_value: str
    # hyperlink: Optional[Hyperlink] = None
    # format: str = ''
    colspan: Optional[int] = None
    rowspan: Optional[int] = None

    height: int = 19
    border: Optional[Union[Borders, Border]] = None
    textAlign: Optional[str] = None
    fill: Optional[Fill] = None
    font: Optional[Font] = None

    @classmethod
    def from_cell(
        cls, cell: Cell, f_cell: Optional[Cell], _locale: Optional[str] = None
    ) -> "CellInfo":
        cell_info = CellInfo(
            id=get_cell_id(cell),
            row=cell.row,
            column=cell.column,
            column_letter=cell.column_letter,
            coordinate=cell.coordinate,
            value=cell.value,
            formatted_value=format_cell(cell, locale=_locale, f_cell=f_cell),
        )

        if cell.alignment.horizontal:
            cell_info.textAlign = cell.alignment.horizontal

        if cell.fill:
            cell_info.fill = Fill(
                color=normalize_color(cell.fill.fgColor), pattern=cell.fill.patternType
            )
        if cell.font:
            font = Font(size=cell.font.sz)
            if cell.font.color:
                font.color = normalize_color(cell.font.color)
            font.bold = bool(cell.font.b)
            font.italic = bool(cell.font.i)
            font.underline = bool(cell.font.u)
            cell_info.font = font
        cell_info.border = cls.get_border(cell)
        # TODO merged border
        return cell_info

    @staticmethod
    def get_border(cell: Cell) -> Union[Borders, None]:
        border_info = Borders()
        is_none = True
        for b_dir in ["right", "left", "top", "bottom"]:
            b_s = getattr(cell.border, b_dir)
            if not b_s:
                continue
            b = Border(style=b_s.style)
            if b_s.color:
                b.color = normalize_color(b_s.color)

            setattr(border_info, b_dir, b)
            is_none = False
        if is_none:
            return None
        return border_info
