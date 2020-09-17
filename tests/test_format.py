# coding: utf-8
from __future__ import unicode_literals

import datetime

import pytest

from xlsx2html.format import format_decimal, format_date, format_datetime, format_time

decimal_formats = [
    ([-1500, '[RED]0.00', 'ru'], '<span style="color: RED">-1500,00</span>'),
    ([-1500, '[RED]#,##0.00', 'ru'], '<span style="color: RED">-1\xa0500,00</span>'),
    ([-1500, '[Red]0.00', 'ru'], '<span style="color: Red">-1500,00</span>'),
]


@pytest.mark.parametrize('fmt_kw,expected', decimal_formats)
def test_decimal_format(fmt_kw, expected):
    assert format_decimal(*fmt_kw) == expected


currency_formats = [
    (
        [1000, '#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]', 'ru'],
        '1\xa0000,00 р.'
    ),
    (
        [-1000, '#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]', 'ru'],
        '-1\xa0000,00 р.'
    ),
    (
        [-1000, '#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]', 'en'],
        '-1,000.00 р.'
    ),
    (
        [-1000, '#,##0.00 [$USD];[RED]-#,##0.00 [$USD]', 'en'],
        '<span style="color: RED">-1,000.00 USD</span>'
    ),
    (
        [-1000, '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)', 'en'],
        ' $ (1,000.00)'
    ),
    (
        [1000, '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)', 'en'],
        ' $ 1,000.00 '
    ),
    (
        [0, '_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)', 'en'],
        ' $ - '
    ),
]


@pytest.mark.parametrize('fmt_kw,expected', currency_formats)
def test_currency_format(fmt_kw, expected):
    assert format_decimal(*fmt_kw) == expected


dt = datetime.datetime(2019, 12, 25, 6, 9, 5)


@pytest.mark.parametrize('fmt_kw,expected',
                         [
                             ([dt.date(), 'MM/DD/YY', 'ru'], '12/25/19'),
                             ([dt.date(), 'MMMM', 'en'], 'December'),
                             ([dt.date(), 'MMMM', 'ru'], 'декабря'),
                         ]
                         )
def test_format_date(fmt_kw, expected):
    assert format_date(*fmt_kw) == expected


@pytest.mark.parametrize('fmt_kw,expected',
                         [
                             ([dt, 'MM/DD/YY', 'ru'], '12/25/19'),
                             ([dt, 'MMMM', 'en'], 'December'),
                             ([dt, 'MMMM', 'ru'], 'декабря'),
                             ([dt, 'DD/MM/YYYY HH:MM:SS', 'ru'], '25/12/2019 06:09:05'),
                             ([dt, 'DD/MM/YYYY HH:MM AM/PM', 'en'], '25/12/2019 06:09 AM'),
                         ]
                         )
def test_format_datetime(fmt_kw, expected):
    assert format_datetime(*fmt_kw) == expected


@pytest.mark.parametrize('fmt_kw,expected',
                         [
                             ([dt.time(), 'HH:MM:SS', 'ru'], '06:09:05'),
                             ([dt.time(), 'HH:MM AM/PM', 'en'], '06:09 AM'),
                             ([(dt + datetime.timedelta(hours=14)).time(), 'HH:MM:SS AM/PM', 'en'],
                              '08:09:05 PM'),
                             ([(dt + datetime.timedelta(hours=14)).time(), 'HH:MM:SS', 'en'],
                              '20:09:05'),
                         ]
                         )
def test_format_time(fmt_kw, expected):
    assert format_time(*fmt_kw) == expected
