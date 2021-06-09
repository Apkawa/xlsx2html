# coding: utf-8
from __future__ import unicode_literals

import datetime

import pytest

from xlsx2html.format import format_decimal
from xlsx2html.format import format_date, format_datetime, format_time, format_timedelta
from xlsx2html.format.dt import normalize_datetime_format

decimal_formats = [
    ([-1500, "[RED]0.00", "ru"], '<span style="color: RED">-1500,00</span>'),
    ([-1500, "[RED]#,##0.00", "ru"], '<span style="color: RED">-1\xa0500,00</span>'),
    ([-1500, "[Red]0.00", "ru"], '<span style="color: Red">-1500,00</span>'),
]


@pytest.mark.parametrize("fmt_kw,expected", decimal_formats)
def test_decimal_format(fmt_kw, expected):
    assert format_decimal(*fmt_kw) == expected


currency_formats = [
    (
        [1000, "#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]", "ru"],
        "1\xa0000,00 р.",
    ),
    (
        [-1000, "#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]", "ru"],
        "-1\xa0000,00 р.",
    ),
    (
        [-1000, "#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]", "en"],
        "-1,000.00 р.",
    ),
    (
        [-1000, "#,##0.00 [$USD];[RED]-#,##0.00 [$USD]", "en"],
        '<span style="color: RED">-1,000.00 USD</span>',
    ),
    (
        [-1000, "_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)", "en"],
        " $ (1,000.00)",
    ),
    (
        [1000, "_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)", "en"],
        " $ 1,000.00 ",
    ),
    (
        [0, "_(\\$* #,##0.00_);_(\\$* \\(#,##0.00\\);_(\\$* \\-??_);_(@_)", "en"],
        " $ - ",
    ),
]


@pytest.mark.parametrize("fmt_kw,expected", currency_formats)
def test_currency_format(fmt_kw, expected):
    assert format_decimal(*fmt_kw) == expected


test_datetime_formats = {
    "mm-dd-yy": "MM-dd-yy",
    "d-mmm-yy": "d-MMM-yy",
    "d-mmm-yyyy": "d-MMM-yyyy",
    "d-mmm": "d-MMM",
    "mmm-yy": "MMM-yy",
    "h:mm AM/PM": "h:mm a",
    "h:mm:ss AM/PM": "h:mm:ss a",
    "h:mm": "H:mm",
    "h:mm:ss": "H:mm:ss",
    "m/d/yy h:mm": "M/d/yy H:mm",
    "mm:ss": "mm:ss",
    "mm:ss.0": "mm:ss.S",
    "yyyy-mm-dd hh:mm:ss.000": "yyyy-MM-dd HH:mm:ss.SSS",
    "h m m s": "H m m s",
    "h m m s m": "H m m s M",
    "h m s m": "H m s M",
    "m s m": "m s M",
    "s m": "s m",
    "h m d s m": "H m d s M",
    "y s d mmm m": "yy s d MMM m",
    "y s d m": "yy s d m",
    "m s y m": "m s yy M",
    "h mmm s m": "H MMM s m",
    "h mmm m s m": "H MMM m s M",
    "h m s am": "H m s 'a'M",
    'h m s"s"': "H m s's'",
    "h m s'": "H m s''",
    "h m s''": "H m s''''",
    "h m s\"''\"": "H m s''''",
    'h m s"a\'"': "H m s'a'''",
}


@pytest.mark.parametrize("xlfmt, bfmt", test_datetime_formats.items())
def test_normalize_format(xlfmt, bfmt):
    assert normalize_datetime_format(xlfmt) == bfmt


dt = datetime.datetime(2019, 12, 25, 6, 9, 5)


@pytest.mark.parametrize(
    "fmt_kw,expected",
    [
        ([dt.date(), "MM/DD/YY", "ru"], "12/25/19"),
        ([dt.date(), "MMMM", "en"], "December"),
        ([dt.date(), "MMMM", "ru"], "декабря"),
    ],
)
def test_format_date(fmt_kw, expected):
    assert format_date(*fmt_kw) == expected


@pytest.mark.parametrize(
    "fmt_kw,expected",
    [
        ([dt, "MM/DD/YY", "ru"], "12/25/19"),
        ([dt, "MMMM", "en"], "December"),
        ([dt, "MMMM", "ru"], "декабря"),
        ([dt, "DD/MM/YYYY HH:MM:SS", "ru"], "25/12/2019 06:09:05"),
        ([dt, "DD/MM/YYYY HH:MM AM/PM", "en"], "25/12/2019 06:09 AM"),
    ],
)
def test_format_datetime(fmt_kw, expected):
    assert format_datetime(*fmt_kw) == expected


@pytest.mark.parametrize(
    "fmt_kw,expected",
    [
        ([dt.time(), "HH:MM:SS", "ru"], "06:09:05"),
        ([dt.time(), "HH:MM AM/PM", "en"], "06:09 AM"),
        (
            [(dt + datetime.timedelta(hours=14)).time(), "HH:MM:SS AM/PM", "en"],
            "08:09:05 PM",
        ),
        ([(dt + datetime.timedelta(hours=14)).time(), "HH:MM:SS", "en"], "20:09:05"),
    ],
)
def test_format_time(fmt_kw, expected):
    assert format_time(*fmt_kw) == expected


td = datetime.timedelta(hours=2, minutes=3, seconds=4, milliseconds=5)
td2 = datetime.timedelta(hours=2, minutes=3, seconds=4, milliseconds=999)

timedelta_arge = [
    ([td, "[hhh]:mm:ss.000"], "002:03:04.005"),
    ([td, "[hhh]:mm:ss.00"], "002:03:04.01"),
    ([td2, "[hhh]:mm:ss.000"], "002:03:04.999"),
    ([td2, "[h]:mm:ss"], "2:03:05"),
    ([td2, "[mmmm]:ss"], "0123:05"),
    ([td2, "[s]"], "7385"),
    ([td, "[s].00"], "7384.01"),
    ([td, "[hhh] - mm - ss.000 x"], "002 - 03 - 04.005 x"),
]


@pytest.mark.parametrize("args,expected", timedelta_arge)
def test_format_timedelta(args, expected):
    assert format_timedelta(*args) == expected
