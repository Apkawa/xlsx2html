from typing import List, Dict, Optional, Union, cast

from xlsx2html.constants.border import DEFAULT_BORDER_STYLE, BORDER_STYLES
from xlsx2html.utils.render import render_attrs, render_inline_styles
from xlsx2html.parser.cell import CellInfo, Border
from xlsx2html.parser.image import ImageInfo
from xlsx2html.parser.parser import Column, ParserResult

StyleType = Dict[str, Union[str, None]]


class HtmlRenderer:
    def __init__(
        self,
        display_grid: bool = False,
        default_border_style: Optional[StyleType] = None,
    ):
        self.default_border_style = default_border_style or {}
        self.display_grid = display_grid

    def render(self, result: ParserResult) -> str:
        html = """
        <!DOCTYPE html>
        <html lang="en">
        <head>
            <meta charset="UTF-8">
            <title>Title</title>
        </head>
        <body>
            %s
        </body>
        </html>
        """
        return html % self.render_table(result)

    def render_table(self, result: ParserResult) -> str:
        h = [
            "<table  "
            'style="border-collapse: collapse" '
            'border="0" '
            'cellspacing="0" '
            'cellpadding="0">',
            self.render_columns(result.cols),
        ]

        if self.display_grid:
            h.append(self.render_header(result.cols))

        for row in result.rows:
            trow = ["<tr>"]
            for i, cell in enumerate(row):
                if i == 0 and self.display_grid:
                    trow.append(self.render_lineno(cell.row))
                images = result.images.get((cell.column, cell.row)) or []
                trow.append(self.render_cell(cell, images))
            trow.append("</tr>")
            h.append("\n".join(trow))
        h.append("</table>")
        return "\n".join(h)

    def render_header(self, cols: List[Column]) -> str:
        h = ["<thead><tr><td></td>"]

        for col in cols:
            h.append(f"<th>{col.letter}<th>")
        h.append("</tr></thead>")
        return "\n".join(h)

    def render_lineno(self, lineno: int) -> str:
        return f"<th>{lineno}</th>"

    def render_columns(self, cols: List[Column]) -> str:
        h = ["<colgroup>"]

        for c in cols:
            h.append(self.render_column(c))
        h.append("</colgroup>")
        return "\n".join(h)

    def render_column(self, col: Column) -> str:
        return f'<col style="width: {col.width}px"/>'

    def render_cell(self, cell: CellInfo, images: List[ImageInfo]):
        formatted_images = "\n".join([self.render_image(img) for img in images])

        attrs = {"id": cell.id, "colspan": cell.colspan, "rowspan": cell.rowspan}

        styles = self.get_styles_from_cell(cell)

        return (
            '<td {attrs_str} style="{styles_str}">'
            "{formatted_images}"
            "{formatted_value}"
            "</td>"
        ).format(
            attrs_str=render_attrs(attrs),
            styles_str=render_inline_styles(styles),
            formatted_images=formatted_images,
            formatted_value=cell.formatted_value,
        )

    def get_border_style_from_cell(self, cell: CellInfo) -> StyleType:
        style: StyleType = {}
        if not cell.border:
            return style

        def _get_border_style(b: Border, prefix="border") -> StyleType:
            border_style: StyleType = cast(StyleType, BORDER_STYLES.get(b.style) or {})
            _style: StyleType = {}
            if not border_style and b.style:
                border_style = cast(StyleType, DEFAULT_BORDER_STYLE)
            elif not b.style:
                border_style = self.default_border_style

            for k, v in border_style.items():
                _style[f"{prefix}-{k}"] = v
            if b.color:
                _style[f"{prefix}-color"] = b.color
            return _style

        if isinstance(cell.border, Border):
            style.update(_get_border_style(cell.border))
        else:
            for b_dir in ["right", "left", "top", "bottom"]:
                b_s = getattr(cell.border, b_dir)
                if not b_s:
                    continue
                style.update(_get_border_style(b_s, f"border-{b_dir}"))
        return style

    def get_styles_from_cell(self, cell: CellInfo) -> StyleType:
        h_styles: StyleType = {
            "border-collapse": "collapse",
            "height": f"{cell.height}pt",
        }
        h_styles.update(self.get_border_style_from_cell(cell))

        if cell.textAlign:
            h_styles["text-align"] = cell.textAlign

        if cell.fill and cell.fill.pattern == "solid":
            # TODO patternType != 'solid'
            h_styles["background-color"] = cell.fill.color

        if cell.font:
            h_styles["font-size"] = "%spx" % cell.font.size
            if cell.font.color:
                h_styles["color"] = cell.font.color
            if cell.font.bold:
                h_styles["font-weight"] = "bold"
            if cell.font.italic:
                h_styles["font-style"] = "italic"
            if cell.font.underline:
                h_styles["font-decoration"] = "underline"
        return h_styles

    def render_image(self, image: ImageInfo) -> str:
        styles = render_inline_styles(
            {
                "margin-left": f"{image.offset.x}px",
                "margin-top": f"{image.offset.y}px",
                "position": "absolute",
            }
        )
        return (
            f'<img width="{image.width}" height="{image.height}"'
            f'style="{styles}"'
            f'src="{image.src}"'
            "/>"
        )
