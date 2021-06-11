import io
from dataclasses import dataclass, InitVar
from typing import Optional, Union, TextIO, List, BinaryIO, cast

from xlsx2html.parser.parser import XLSXParser
from xlsx2html.parser.utils import SheetNameType
from xlsx2html.render.html import HtmlRenderer, BorderType, StyleType

OutputType = Optional[Union[TextIO, str]]
FilePathType = Union[BinaryIO, str]


@dataclass
class ConverterTableResult:
    """
    :param html: html of table
    :param css: css contents
        If :paramref:`XLSX2HTMLConverter.optimize_styles` set to `False` then css is empty
    """

    html: str
    css: str = ""


@dataclass
class XLSX2HTMLConverter:
    """
    :param filepath: xlsx file
    :type filepath: str | BinaryIO
    :param locale: ``en`` or ``zh_TW``. defaults to ``en``
    :type locale: str
    :param parse_formula: If `True` - enable parse formulas. defaults to `False`
    :param default_border_style: default border style. Can use short str like ``1px solid black``
        or dict like ``{'width': '1px', 'style': 'solid', 'color': 'black'}``
    :param optimize_styles: for split inline styles in cells and render separately
    :param display_grid:

        Show column letters and row numbers.
        If :paramref:`XLSX2HTMLConverter.default_border_style` is none - do enabled gray grid

    :type default_cell_border: str|dict, optional
    """

    filepath: FilePathType
    locale: str = "en"
    parse_formula: bool = False
    optimize_styles: bool = False
    display_grid: bool = False
    default_border_style: BorderType = None

    parser: InitVar[XLSXParser] = None
    renderer: InitVar[HtmlRenderer] = None

    def __post_init__(self, parser, renderer):
        self.parser: XLSXParser = parser or XLSXParser(
            filepath=self.filepath, parse_formula=self.parse_formula, locale=self.locale
        )
        self.renderer: HtmlRenderer = renderer or HtmlRenderer(
            default_border_style=self.default_border_style,
            optimize_styles=self.optimize_styles,
            display_grid=self.display_grid,
        )

    def _get_stream(self, output: OutputType) -> TextIO:
        if not output:
            output = io.StringIO()
        if isinstance(output, str):
            output = open(output, "w")
        return cast(TextIO, output)

    def get_table(
        self, sheet: SheetNameType = None, extra_attrs: Optional[StyleType] = None
    ) -> ConverterTableResult:
        """

        :param sheet: sheet name or idx, defaults to `None` what means get active sheet
        :param extra_attrs: additional attributes for `<table>` like class or id
        :return:
        """
        result = self.parser.get_sheet(sheet)
        self.renderer.build_style_cache(result.rows)
        return ConverterTableResult(
            html=self.renderer.render_table(result, attrs=extra_attrs),
            css=self.renderer.render_css(),
        )

    def get_tables(
        self,
        sheets: Optional[List[SheetNameType]] = None,
        extra_attrs: Optional[StyleType] = None,
    ) -> List[ConverterTableResult]:
        """

        :param sheets: list of sheet name or idx. By defaults get all sheets
        :param extra_attrs: additional attributes for `<table>` like class or id
        :return:
        """
        if sheets is None:
            sheets = cast(List[SheetNameType], self.parser.get_sheet_names())
        return [
            self.get_table(sheet, extra_attrs=extra_attrs) for sheet in sheets or []
        ]

    def get_html(self, sheet: SheetNameType = None) -> str:
        """
        Get full html with table
        :param sheet: sheet name or idx, defaults to `None` what means get active sheet
        :return: full html as string
        """
        result = self.parser.get_sheet(sheet)
        return self.renderer.render(result)

    def get_html_stream(
        self, output: OutputType, sheet: SheetNameType = None
    ) -> TextIO:
        """
        :param output: to path or file like, defaults to `None`
        :param sheet: sheet name or idx, defaults to `None` what means get active sheet
        :return: File like object
        """
        html = self.get_html(sheet)
        stream = self._get_stream(output)
        stream.write(html)
        stream.flush()
        return stream


def xlsx2html(
    filepath: FilePathType,
    output: OutputType = None,
    locale: str = "en",
    sheet: SheetNameType = None,
    parse_formula: bool = False,
    default_cell_border: BorderType = None,
) -> TextIO:
    """

    :param filepath: xlsx file
    :param output: to path or file like, defaults to `None`
    :param locale: ``en`` or ``zh_TW``. defaults to ``en``
    :param sheet: sheet name or idx, defaults to `None` what means get active sheet
    :param parse_formula: If `True` - enable parse formulas. defaults to `False`
    :param default_cell_border:
        default border style. Can use short str like ``1px solid black``
        or dict like ``{'width': '1px', 'style': 'solid', 'color': 'black'}``

    :return: File like object
    """
    converter = XLSX2HTMLConverter(
        filepath=filepath,
        locale=locale,
        parse_formula=parse_formula,
        default_border_style=default_cell_border,
    )
    output = converter.get_html_stream(output, sheet)
    return output
