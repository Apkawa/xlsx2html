from typing import List, Dict, Optional, cast

import csscompressor

from xlsx2html.constants.border import DEFAULT_BORDER_STYLE, BORDER_STYLES
from xlsx2html.parser.cell import CellInfo, Border
from xlsx2html.parser.image import ImageInfo
from xlsx2html.parser.parser import Column, ParserResult
from xlsx2html.utils import hash_str
from xlsx2html.utils.render import render_attrs, render_inline_styles
from xlsx2html.utils.style import compress_style, StyleType, BorderType


class HtmlRenderer:
    def __init__(
        self,
        display_grid: bool = False,
        default_border_style: Optional[BorderType] = None,
        table_attrs: Optional[StyleType] = None,
        inline_styles: bool = False,
    ):
        self.default_border_style = default_border_style or {}
        self.display_grid = display_grid
        self.table_attrs = table_attrs or {}
        self.inline_styles = inline_styles

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

        self.build_style_cache(result.rows)
        h = [self.render_table(result)]
        if not self.inline_styles:
            css_tag = f'<style type="text/css">{self.render_css() or ""}</style>'
            h.append(css_tag)
        return html % "\n".join(h)

    def render_table(self, result: ParserResult, attrs: Optional[StyleType] = None) -> str:

        t_attrs: StyleType = dict(border="0", cellspacing="0", cellpadding="0")
        t_attrs.update(self.table_attrs)
        t_attrs.update(attrs or {})
        h = [
            "<table  " 'style="border-collapse: collapse" ' f"{render_attrs(t_attrs)}" ">",
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
            h.append(f"<th>{col.letter}</th>")
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

    def render_cell(
        self, cell: CellInfo, images: List[ImageInfo], attrs: Optional[StyleType] = None
    ) -> str:

        formatted_images = "\n".join([self.render_image(img) for img in images])

        c_attrs = {"id": cell.id, "colspan": cell.colspan, "rowspan": cell.rowspan}

        class_name = self._cell_style_map[cell.coordinate]
        if not self.inline_styles:
            c_attrs["class"] = class_name
        else:
            c_attrs["style"] = self._style_hash_map[class_name]

        c_attrs.update(attrs or {})

        return ("<td {attrs_str}>" "{formatted_images}" "{formatted_value}" "</td>").format(
            attrs_str=render_attrs(c_attrs),
            formatted_images=formatted_images,
            formatted_value=cell.formatted_value,
        )

    def get_border_style_from_cell(self, cell: CellInfo) -> StyleType:
        style: StyleType = {}
        if not cell.border:
            return style

        def _get_border_style(b: Border, prefix: str = "border") -> StyleType:
            border_style: StyleType = cast(StyleType, BORDER_STYLES.get(b.style) or {})
            _style: StyleType = {}
            if not border_style and b.style:
                border_style = cast(StyleType, DEFAULT_BORDER_STYLE)
            elif not b.style:
                if isinstance(self.default_border_style, str):
                    # Maybe shortland
                    return {prefix: self.default_border_style}
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

    def get_styles_from_cell(
        self, cell: CellInfo, extra_style: Optional[StyleType] = None
    ) -> StyleType:
        h_styles: StyleType = {"border-collapse": "collapse", "height": f"{cell.height}pt"}
        h_styles.update(self.get_border_style_from_cell(cell))
        h_styles.update(extra_style or {})

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

    def build_style_cache(self, rows: List[List[CellInfo]]) -> None:
        cell_style_map: Dict[str, str] = {}
        style_hash_map: Dict[str, str] = {}

        for row in rows:
            for cell in row:
                style = self.get_styles_from_cell(cell)
                style = compress_style(style)
                style_str = render_inline_styles(style)
                class_name = "td-" + hash_str(style_str)
                style_hash_map[class_name] = style_str
                cell_style_map[cell.coordinate] = class_name

        self._cell_style_map = cell_style_map
        self._style_hash_map = style_hash_map

    def render_css(self) -> str:
        if self.inline_styles:
            return ""

        css = []
        for c_name, style in self._style_hash_map.items():
            css.append(f"td.{c_name} {{ {style} }}")
        css_str = "\n".join(css)
        # TODO add compress css
        css_str = csscompressor.compress(css=css_str)
        return css_str
