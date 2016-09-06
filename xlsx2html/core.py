# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import sys

import openpyxl
import six

BORDER_STYLES = {
    'dashDot': {},
    'dashDotDot': {},
    'dashed': {},
    'dotted': {},
    'double': {},
    'hair': {},
    'medium': {
        'style': 'solid',
        'width': '2px',
    },
    'mediumDashDot': {},
    'mediumDashDotDot': {},
    'mediumDashed': {},
    'slantDashDot': {},
    'thick': {},
    'thin': {},
}


def render_attrs(attrs):
    return ' '.join(["%s=%s" % a for a in sorted(attrs.items(), key=lambda a: a[0])])


def render_inline_styles(styles):
    return ';'.join(["%s: %s" % a for a in sorted(styles.items(), key=lambda a: a[0])])


def get_border_style_from_cell(cell):
    h_styles = {}
    for b_dir in ['right', 'left', 'top', 'bottom']:
        b_s = getattr(cell.border, b_dir)
        if not b_s:
            continue
        for k, v in BORDER_STYLES.get(b_s.style, {}).items():
            h_styles['border-%s-%s' % (b_dir, k)] = v
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
    h_styles['font-size'] = "%spx" % cell.font.sz
    if cell.font.color.rgb:
        h_styles['color'] = "#" + cell.font.color.rgb[2:]
    if cell.fill.patternType == 'solid':
        #TODO patternType != 'solid'
        h_styles['background-color'] = '#' + cell.fill.fgColor.rgb[2:]

    if cell.font.b:
        h_styles['font-weight'] = 'bold'
    if cell.font.i:
        h_styles['font-style'] = 'italic'
    if cell.font.u:
        h_styles['font-decoration'] = 'underline'
    return h_styles


def worksheet_to_data(ws):
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

    data_list = []
    for row in ws.iter_rows():
        data_row = []
        data_list.append(data_row)
        for cell in row:
            col_dim = ws.column_dimensions[cell.column]
            row_dim = ws.row_dimensions[cell.row]
            if cell.coordinate in exclded_cells or row_dim.hidden or col_dim.hidden:
                continue

            width = 0.89
            height = 19

            if col_dim.customWidth:
                width = round(col_dim.width / 10., 2)
            if row_dim.customHeight:
                height = round(row_dim.height, 2)

            cell_data = {
                'value': cell.value or '&nbsp;',
                'attrs': {},
                'col-width': 96 * (width),
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
    return data_list


def render_table(data):
    html = [
        '<table  style="border-collapse: collapse" border="0" cellspacing="0" cellpadding="0">'
    ]
    for i, row in enumerate(data):
        if i == 0:
            html.append('<colgroup>')
            for col in row:
                html.append('<col width="%s">' % int(col['col-width']))
            html.append('</colgroup>')
        trow = ['<tr>']
        for col in row:
            trow.append('<td {attrs_str} style="{styles_str}">{value}</td>'.format(
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


def xls2html(filepath, output):
    ws = openpyxl.load_workbook(filepath, data_only=True).active
    data = worksheet_to_data(ws)
    html = render_data_to_html(data)

    with open(output, 'wb') as f:
        f.write(six.binary_type(html.encode('utf-8')))


if __name__ == '__main__':
    xls2html(sys.argv[1], sys.argv[2])