# coding: utf-8
from __future__ import unicode_literals

from xlsx2html.format import format_decimal


def test_currency_format():
    fmt = u'#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]'
    assert format_decimal(1000, format=fmt, locale='ru') == '1\xa0000,00 р.'
    assert format_decimal(-1000, format=fmt, locale='ru') == '-1\xa0000,00 р.'
