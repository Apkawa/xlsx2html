from __future__ import unicode_literals

import re

from babel import dates as babel_dates
from babel.dates import LC_TIME

DATE_REPLACES = {
    'DD': 'dd',
    'YYYY': 'yyyy',
    'YY': 'yy',
}
RE_TIME = re.compile(
    r"""
        \b
        (?P<hours>[H]{1,2})
        .?
        (?P<minutes>[M]{1,2})
        (?:
            .?
            (?P<seconds>[S]{1,2})
            |
        )
        (?:
            .+?
            (?P<h12>AM/PM)
            |
        )
        \b
        """,
    re.VERBOSE

)
FIX_TIME_REPLACES = {
    r'\b([H]{1,2}).?([M]{1,2}).?([S]{1,2}).+?(AM/PM)\b': r'hh\1m\2s\3a',
    r'\bHH(.?)MM(.?)SS\b': r'H\1m\2s',
    r'AM/PM': 'a'
}



def normalize_date_format(_format):
    for f, to in DATE_REPLACES.items():
        _format = _format.replace(f, to)
    return _format


def normalize_time_format(_format):
    def replace_time(m):
        groups = m.groupdict()
        text = m.group(0).lower()
        if groups.get('h12'):
            text = text.replace(groups['h12'].lower(), 'a')
        else:
            text = text.replace('h', 'H')
        return text

    _format = RE_TIME.sub(replace_time, _format)
    return _format.replace('\\', '')


def normalize_datetime_format(_format):
    for fn in [normalize_time_format, normalize_date_format]:
        _format = fn(_format)
    return _format.replace('\\', '')


def format_date(date, fmt, locale=LC_TIME):
    fmt = normalize_date_format(fmt)
    return babel_dates.format_date(date, fmt, locale)


def format_datetime(date, fmt, locale=LC_TIME, tzinfo=None):
    fmt = normalize_datetime_format(fmt)
    return babel_dates.format_datetime(date, fmt, locale=locale, tzinfo=tzinfo)


def format_time(date, fmt, locale=LC_TIME, tzinfo=None):
    fmt = normalize_time_format(fmt)
    return babel_dates.format_time(date, fmt, locale=locale, tzinfo=tzinfo)
