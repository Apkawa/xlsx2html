# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import sys

import openpyxl

BORDER_STYLES = {
    'medium': {
        'style': 'solid',
        'width': '2px',
    }
}


def render_attrs(attrs):
    return ' '.join(["%s=%s" % a for a in attrs.items()])


def render_inline_styles(styles):
    return ';'.join(["%s: %s" % a for a in styles.items()])


def get_styles_from_cell(cell):
    h_styles = {
        'border-collapse': 'collapse'
    }
    for b_dir in ['right', 'left', 'top', 'bottom']:
        b_s = getattr(cell.border, b_dir)
        if b_s:
            for k, v in BORDER_STYLES.get(b_s.style, {}).items():
                h_styles['border-%s-%s' % (b_dir, k)] = v
        else:
            h_styles['border-%s-style' % b_dir] = 'none'
    if cell.alignment:
        h_styles['text-align'] = cell.alignment.horizontal
    h_styles['font-size'] = "%spx" % cell.font.sz
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
            'colspan': len(cell_range_list[0]),
            'rowspan': len(cell_range_list),
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
                width = col_dim.width / 10
            if row_dim.customHeight:
                height = row_dim.height

            cell_data = {
                'value': cell.value or '&nbsp;',
                'attrs': {},
                'col-width': 96 * (width),
                'style': {
                    "width": "{}in".format(width or 0.89),
                    "height": "{}px".format(height or 19),
                },
            }
            cell_data['attrs'].update(merged_cell_map.get(cell.coordinate) or {})
            cell_data['style'].update(get_styles_from_cell(cell))
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
        f.write(str(html.encode('utf-8')))


if __name__ == '__main__':
    xls2html(sys.argv[1], sys.argv[2])
