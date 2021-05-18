from __future__ import unicode_literals

import re

from babel import dates as babel_dates
from babel.dates import LC_TIME

RE_DATE_TOK = re.compile(r"(?:y+|m+|d+|h+|s+|\.0+|am/pm|a/p)", re.I)
MAYBE_MINUTE = ['m', 'mm']


def normalize_datetime_format(fmt):
    has_ap = False
    is_minute = set()
    must_minute = False
    found = [(m.group(0).lower(), *m.span()) for m in RE_DATE_TOK.finditer(fmt)]
    for i, (tok, start, end) in enumerate(found):
        if tok in ['am/pm', 'a/p']:
            has_ap = True
        elif tok[0] == 'h':
            # First m after h is always minute
            must_minute = True
        elif must_minute and tok in ['m', 'mm']:
            is_minute.add(i)
            must_minute = False
        elif tok[0] == 's':
            last_i = i - 1
            if last_i < 0:
                must_minute = True
            elif last_i not in is_minute:
                if found[last_i][0] in ['m', 'mm']:
                    # m right before s is always minute
                    is_minute.add(last_i)
                elif not len(is_minute):
                    # if no previous m, first m after s is always minute
                    must_minute = True

    parts = []
    pos = 0
    for i, (tok, start, end) in enumerate(found):
        parts.append(fmt[pos:start])
        if tok[0] == 'h':
            tok = tok[:2]
            if not has_ap:
                tok = tok.upper()
        elif tok[0] == 'm':
            if len(tok) > 5:
                tok = tok[:4]  # Defaults to MMMM
            if len(tok) > 2 or i not in is_minute:
                tok = tok.upper()
        elif tok[0] == 's':
            tok = tok[:2]
        elif tok[:2] == '.0':
            tok = tok.replace('0', 'S')
        elif tok == 'am/pm':
            tok = 'a'
        elif tok == 'a/p':
            # Fudging presentation to AM; maybe should be A in some cases?
            tok = 'a'
        elif tok[0] == 'y':
            if len(tok) > 2:
                tok = 'yyyy'
            else:
                tok = 'yy'
        elif tok[0] == 'd':
            if len(tok) == 3:
                tok = 'EEE'
            elif len(tok) >= 4:
                tok = 'EEEE'
        else:
            raise ValueError(f'Unhandled datetime token {tok}')
        parts.append(tok)
        pos = end
    parts.append(fmt[pos:])

    return ''.join(parts)


def format_date(date, fmt, locale=LC_TIME):
    fmt = normalize_datetime_format(fmt)
    return babel_dates.format_date(date, fmt, locale)


def format_datetime(date, fmt, locale=LC_TIME, tzinfo=None):
    fmt = normalize_datetime_format(fmt)
    return babel_dates.format_datetime(date, fmt, locale=locale, tzinfo=tzinfo)


def format_time(date, fmt, locale=LC_TIME, tzinfo=None):
    fmt = normalize_datetime_format(fmt)
    return babel_dates.format_time(date, fmt, locale=locale, tzinfo=tzinfo)
