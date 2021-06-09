import re

from babel import Locale, UnknownLocaleError

from xlsx2html.constants import LCID_HEX_MAP

LOCALE_FORMAT_RE = re.compile(
    r"""
    \[
    \$
        (?:.+|)
        -(?P<lcid>[0-9A-Fa-f]{3,4})
    \]
    """,
    re.VERBOSE,
)


def parse_locale_code(code):
    """
    >>> parse_locale_code('-404')
    'zh_Hant_TW'
    >>> parse_locale_code('404')
    'zh_Hant_TW'
    >>> parse_locale_code('0404')
    'zh_Hant_TW'
    >>> parse_locale_code('58050')
    """
    try:
        lcid = abs(int(code, 16))
        locale_code = LCID_HEX_MAP.get(lcid, "UNKNOWN")
        return str(Locale.parse(locale_code))
    except UnknownLocaleError:
        return None


def extract_locale_from_format(fmt):
    """
    >>> extract_locale_from_format('[$-404]e/m/d')
    ('zh_Hant_TW', 'e/m/d')
    >>> extract_locale_from_format('[$USD-404]e/m/d')
    ('zh_Hant_TW', 'e/m/d')
    >>> extract_locale_from_format('[$$-404]#.00')
    ('zh_Hant_TW', '#.00')
    >>> extract_locale_from_format('[RED]e/m/d')
    (None, '[RED]e/m/d')
    """
    locale = None
    m = LOCALE_FORMAT_RE.match(fmt)
    if not m:
        return locale, fmt

    win_locale = m.group()
    # todo keep currency
    new_locale = parse_locale_code(m.group(1))
    if new_locale:
        locale = new_locale
    fmt = fmt.replace(win_locale, "")
    return locale, fmt
