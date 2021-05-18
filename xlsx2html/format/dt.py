from __future__ import unicode_literals

import re

from babel import dates as babel_dates
from babel.dates import LC_TIME

RE_DATE_TOK = re.compile(r'(?:\\[\\*_"]?|_.|\*.|y+|m+|d+|h+|s+|\.0+|am/pm|a/p|"[^"]*")', re.I)
MAYBE_MINUTE = ['m', 'mm']
DATE_PERIOD = ['am/pm', 'a/p']


def normalize_datetime_format(fmt):
    has_ap = False
    is_minute = set()
    must_minute = False
    found = [(m.group(0), *m.span()) for m in RE_DATE_TOK.finditer(fmt)]
    for i, (text, start, end) in enumerate(found):
        tok = text.lower()
        if tok in DATE_PERIOD:
            has_ap = True
        elif tok[0] == 'h':
            # First m after h is always minute
            must_minute = True
        elif must_minute and tok in MAYBE_MINUTE:
            is_minute.add(i)
            must_minute = False
        elif tok[0] == 's':
            last_i = i - 1
            if last_i < 0:
                must_minute = True
            elif last_i not in is_minute:
                if found[last_i][0] in MAYBE_MINUTE:
                    # m right before s is always minute
                    is_minute.add(last_i)
                elif not len(is_minute):
                    # if no previous m, first m after s is always minute
                    must_minute = True

    parts = []
    pos = 0
    plain = []

    def clean_plain():
        def s(m):
            g = m.group().replace("'", "''")
            if not re.fullmatch(r"'*", g):
                g = f"'{g}'"
            return g
        t = ''.join(plain)
        t = re.sub(r"[a-z']+", s, t)
        print(plain, repr(t))
        return t

    for i, (text, start, end) in enumerate(found):
        tok = text.lower()
        plain.append(fmt[pos:start])
        pos = end
        if tok[0] == '\\':
            # Escape sequence \*_"
            plain.append(text[1:])
            continue
        elif tok[0] == '_':
            # normally puts space at same size as next character; just treat as space
            plain.append(' ')
            continue
        elif tok[0] == '*':
            # Don't include repeating character
            continue
        elif tok[0] == '"':
            # Quoted string
            plain.append(text[1:-1])
            continue
        elif tok[0] == 'h':
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
        elif tok in DATE_PERIOD:
            # Fudging A/P to AM; maybe should be A in some cases?
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
        if len(plain):
            parts.append(clean_plain())
            plain = []
        parts.append(tok)
    plain.append(fmt[pos:])
    if len(plain):
        parts.append(clean_plain())
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
