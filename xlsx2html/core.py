# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import openpyxl
import six
from openpyxl.styles.colors import COLOR_INDEX, aRGB_REGEX

from xlsx2html.format import format_cell

DEFAULT_BORDER_STYLE = {
    'style': 'solid',
    'width': '1px',
}

BORDER_STYLES = {
    'dashDot': None,
    'dashDotDot': None,
    'dashed': {
        'style': 'dashed',
    },
    'dotted': {
        'style': 'dotted',
    },
    'double': {
        'style': 'double',
    },
    'hair': None,
    'medium': {
        'style': 'solid',
        'width': '2px',
    },
    'mediumDashDot': {
        'style': 'solid',
        'width': '2px',
    },
    'mediumDashDotDot': {
        'style': 'solid',
        'width': '2px',
    },
    'mediumDashed': {
        'width': '2px',
        'style': 'dashed',
    },
    'slantDashDot': None,
    'thick': {
        'style': 'solid',
        'width': '1px',
    },
    'thin': {
        'style': 'solid',
        'width': '1px',
    },
}


def render_attrs(attrs):
    return ' '.join(["%s=%s" % a for a in sorted(attrs.items(), key=lambda a: a[0])])


def render_inline_styles(styles):
    return ';'.join(["%s: %s" % a for a in sorted(styles.items(), key=lambda a: a[0]) if a[1] is not None])


def normalize_color(color):
    # TODO RGBA
    rgb = None
    if color.type == 'rgb':
        rgb = color.rgb
    if color.type == 'indexed':
        rgb = COLOR_INDEX[color.indexed]
        if not aRGB_REGEX.match(rgb):
            # TODO system fg or bg
            rgb = '00000000'
    if rgb:
        return '#' + rgb[2:]
    return None


def get_border_style_from_cell(cell):
    h_styles = {}
    for b_dir in ['right', 'left', 'top', 'bottom']:
        b_s = getattr(cell.border, b_dir)
        if not b_s:
            continue
        border_style = BORDER_STYLES.get(b_s.style)
        if border_style is None and b_s.style:
            border_style = DEFAULT_BORDER_STYLE

        if not border_style:
            continue

        for k, v in border_style.items():
            h_styles['border-%s-%s' % (b_dir, k)] = v
        if b_s.color:
            h_styles['border-%s-color' % (b_dir)] = normalize_color(b_s.color)

    return h_styles


def get_styles_from_cell(cell, merged_cell_map=None):
    merged_cell_map = merged_cell_map or {}

    h_styles = {
        'border-collapse': 'collapse'
    }
    b_styles = get_border_style_from_cell(cell)
    if merged_cell_map:
        # TODO edged_cells
        for m_cell in merged_cell_map['cells']:
            b_styles.update(get_border_style_from_cell(m_cell))

    for b_dir in ['border-right-style', 'border-left-style', 'border-top-style', 'border-bottom-style']:
        if b_dir not in b_styles:
            b_styles[b_dir] = 'none'
    h_styles.update(b_styles)

    if cell.alignment.horizontal:
        h_styles['text-align'] = cell.alignment.horizontal

    if cell.fill.patternType == 'solid':
        # TODO patternType != 'solid'
        h_styles['background-color'] = normalize_color(cell.fill.fgColor)
    if cell.font:
        h_styles['font-size'] = "%spx" % cell.font.sz
        if cell.font.color:
            h_styles['color'] = normalize_color(cell.font.color)
        if cell.font.b:
            h_styles['font-weight'] = 'bold'
        if cell.font.i:
            h_styles['font-style'] = 'italic'
        if cell.font.u:
            h_styles['font-decoration'] = 'underline'
    return h_styles


def worksheet_to_data(ws, locale=None):
    merged_cell_map = {}
    exclded_cells = set(ws.merged_cells)

    for cell_range in ws.merged_cell_ranges:
        cell_range_list = list(ws[cell_range])
        m_cell = cell_range_list[0][0]

        merged_cell_map[m_cell.coordinate] = {
            'attrs': {
                'colspan': len(cell_range_list[0]),
                'rowspan': len(cell_range_list),
            },
            'cells': [c for rows in cell_range_list for c in rows],
        }
        exclded_cells.remove(m_cell.coordinate)

    max_col_number = 0

    data_list = []
    for row_i, row in enumerate(ws.iter_rows()):
        data_row = []
        data_list.append(data_row)
        for col_i, cell in enumerate(row):
            col_dim = ws.column_dimensions[cell.column]
            row_dim = ws.row_dimensions[cell.row]

            width = 0.89
            if col_dim.customWidth:
                width = round(col_dim.width / 10., 2)

            col_width = 96 * width

            if cell.coordinate in exclded_cells or row_dim.hidden or col_dim.hidden:
                continue

            if col_i > max_col_number:
                max_col_number = col_i

            height = 19

            if row_dim.customHeight:
                height = round(row_dim.height, 2)

            cell_data = {
                'value': cell.value,
                'formatted_value': format_cell(cell, locale=locale),
                'attrs': {},
                'col-width': col_width,
                'style': {
                    "width": "{}in".format(width),
                    "height": "{}px".format(height),
                },
            }
            merged_cell_info = merged_cell_map.get(cell.coordinate, {})
            if merged_cell_info:
                cell_data['attrs'].update(merged_cell_info['attrs'])

            cell_data['style'].update(get_styles_from_cell(cell, merged_cell_info))
            data_row.append(cell_data)

    col_list = []
    max_col_number += 1

    column_dimensions = sorted(ws.column_dimensions.items(), key=lambda d: d[0])

    for col_i, col_dim in column_dimensions:
        if not all([col_dim.min, col_dim.max]):
            continue
        width = 0.89
        if col_dim.customWidth:
            width = round(col_dim.width / 10., 2)
        col_width = 96 * width

        for i in six.moves.range((col_dim.max - col_dim.min) + 1):
            max_col_number -= 1
            col_list.append({"col-width": col_width})
            if max_col_number < 0:
                break
    return {'rows': data_list, 'cols': col_list}


def render_table(data):
    html = [
        '<table  '
        'style="border-collapse: collapse" '
        'border="0" '
        'cellspacing="0" '
        'cellpadding="0">'
        '<colgroup>'
    ]

    for col in data['cols']:
        html.append('<col width="%s">' % int(col['col-width']))
    html.append('</colgroup>')

    for i, row in enumerate(data['rows']):
        trow = ['<tr>']
        for col in row:
            trow.append('<td {attrs_str} style="{styles_str}">{formatted_value}</td>'.format(
                attrs_str=render_attrs(col['attrs']),
                styles_str=render_inline_styles(col['style']),
                **col))

        trow.append('</tr>')
        html.append('\n'.join(trow))
    html.append('</table>')
    return '\n'.join(html)


def render_data_to_html(data):
    html = '''
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
    '''
    return html % render_table(data)


def xlsx2html(filepath, output):
    ws = openpyxl.load_workbook(filepath, data_only=True).active
    data = worksheet_to_data(ws, locale='ru')
    html = render_data_to_html(data)

    with open(output, 'wb') as f:
        f.write(six.binary_type(html.encode('utf-8')))
