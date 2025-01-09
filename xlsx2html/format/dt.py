from __future__ import unicode_literals

import datetime as dt
import re

from babel import dates as babel_dates
from babel.dates import LC_TIME

RE_DATE_TOK = re.compile(
    r'(?:\\[\\*_"]?|_.|\*.|y+|m+|d+|h+|s+|\.0+|am/pm|a/p|"[^"]*")', re.I
)
RE_TD_TOK = re.compile(
    r'(?:\\[\\*_"]?|_.|\*.|\[h+\]|\[m+\]|\[s+\](?:\.0+)?|m+|s+(?:\.0+)?|h+|y+|d+|"[^"]*")',
    re.I,
)
MAYBE_MINUTE = ["m", "mm"]
DATE_PERIOD = ["am/pm", "a/p"]


def normalize_datetime_format(fmt, fixed_for_time=False):
    has_ap = False
    is_minute = set()
    must_minute = False
    found = [(m.group(0), *m.span()) for m in RE_DATE_TOK.finditer(fmt)]
    for i, (text, start, end) in enumerate(found):
        tok = text.lower()
        tok_type = tok[0]
        if tok in DATE_PERIOD:
            has_ap = True
        elif tok_type == "h":
            # First m after h is always minute
            must_minute = True
        elif must_minute and tok in MAYBE_MINUTE:
            is_minute.add(i)
            must_minute = False
        elif tok_type == "s":
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

        t = "".join(plain)
        t = re.sub(r"[a-z']+", repl=s, string=t, flags=re.I)
        return t

    for i, (text, start, end) in enumerate(found):
        tok = text.lower()
        tok_type = tok[0]
        plain.append(fmt[pos:start])
        pos = end
        if tok_type == "\\":
            # Escape sequence \*_"
            plain.append(text[1:])
            continue
        elif tok_type == "_":
            # normally puts space at same size as next character; just treat as space
            plain.append(" ")
            continue
        elif tok_type == "*":
            # Don't include repeating character
            continue
        elif tok_type == '"':
            # Quoted string
            plain.append(text[1:-1])
            continue
        elif tok_type == "h":
            tok = tok[:2]
            if not has_ap:
                tok = tok.upper()
        elif tok_type == "m":
            if len(tok) > 5:
                tok = tok[:4]  # Defaults to MMMM
            if len(tok) > 2 or i not in is_minute:
                tok = tok.upper()
        elif tok_type == "s":
            tok = tok[:2]
        elif tok_type == ".":
            tok = tok.replace("0", "S")
        elif tok in DATE_PERIOD:
            # Fudging A/P to AM; maybe should be A in some cases?
            tok = "a"
        elif tok_type == "y":
            if len(tok) > 2:
                tok = "yyyy"
            else:
                tok = "yy"
        elif tok_type == "d":
            if len(tok) == 3:
                tok = "EEE"
            elif len(tok) >= 4:
                tok = "EEEE"
        else:
            raise ValueError(f"Unhandled datetime token {tok}")
        if fixed_for_time:
            if tok == "d":
                plain.append("0")
                continue
            elif tok == "dd":
                plain.append("00")
                continue
        if len(plain):
            parts.append(clean_plain())
            plain = []
        parts.append(tok)
    plain.append(fmt[pos:])
    if len(plain):
        parts.append(clean_plain())
    return "".join(parts)


def format_date(date, fmt, locale=LC_TIME):
    fmt = normalize_datetime_format(fmt)
    datetime = dt.datetime.combine(date, dt.time())
    return babel_dates.format_datetime(datetime, fmt, locale=locale)


def format_datetime(datetime, fmt, locale=LC_TIME, tzinfo=None):
    fmt = normalize_datetime_format(fmt)
    return babel_dates.format_datetime(datetime, fmt, locale=locale, tzinfo=tzinfo)


def format_time(time, fmt, locale=LC_TIME, tzinfo=None, date=None):
    # Excel times are treated as Saturday 1900-01-00, which doesn't exist.
    # So use Saturday 1900-01-06 and force day to 0 instead
    fixed_for_time = False
    if date is None:
        date = dt.date(1900, 1, 6)
        fixed_for_time = True
    datetime = dt.datetime.combine(date, time)
    fmt = normalize_datetime_format(fmt, fixed_for_time=fixed_for_time)
    return babel_dates.format_datetime(datetime, fmt, locale=locale, tzinfo=tzinfo)


def format_timedelta(timedelta, fmt):
    e_s = timedelta.total_seconds()
    e_m, s = divmod(e_s, 60)
    e_m = int(e_m)
    e_h, m = divmod(e_m, 60)

    plain = []
    pos = 0

    for match in RE_TD_TOK.finditer(fmt):
        text = match.group(0)
        start, end = match.span()
        tok = text.lower()
        plain.append(fmt[pos:start])
        pos = end
        tok_type = tok[0]
        if tok_type == "[":
            tok_type = tok[:2]
        if tok_type == "\\":
            # Escape sequence \*_"
            plain.append(text[1:])
        elif tok_type == "_":
            # normally puts space at same size as next character; just treat as space
            plain.append(" ")
        elif tok_type == "*":
            # Don't include repeating character
            pass
        elif tok_type == '"':
            # Quoted string
            plain.append(text[1:-1])
        elif tok_type == "[h":
            f = "{:0>" + str(len(tok) - 2) + "}"
            plain.append(f.format(e_h))
        elif tok_type == "[m":
            f = "{:0>" + str(len(tok) - 2) + "}"
            plain.append(f.format(e_m))
        elif tok_type == "[s":
            if "." in tok:
                mstok = tok.split(".")[1]
                f = "{:0" + str(len(tok) - 2) + "." + str(len(mstok)) + "f}"
            else:
                f = "{:0" + str(len(tok) - 2) + ".0f}"
            plain.append(f.format(e_s))
        elif tok == "m":
            plain.append(str(m))
        elif tok == "mm":
            plain.append(f"{m:0>2}")
        elif tok_type == "s":
            if "." in tok:
                mstok = tok.split(".")[1]
                f = "".join(
                    [
                        "{:0",
                        str(min(len(tok), 2) + len(mstok) + 1),
                        ".",
                        str(len(mstok)),
                        "f}",
                    ]
                )

            else:
                f = "{:0" + str(min(len(tok), 2)) + ".0f}"
            plain.append(f.format(s))
        else:
            raise ValueError(f"Unhandled datetime token {tok}")
    plain.append(fmt[pos:])
    return "".join(plain)
