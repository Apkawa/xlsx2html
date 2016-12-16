# -*- coding: utf-8 -*-
from __future__ import unicode_literals

import datetime
from decimal import Decimal

import re
import six
from babel import Locale
from babel.dates import format_datetime, format_date, format_time, LC_TIME
from babel.numbers import (
    NumberPattern,
    number_re,
    parse_grouping,
    LC_NUMERIC
)

CLEAN_RE = re.compile(r'[*-]')
CLEAN_CURRENCY_RE = re.compile(r'\[\$(.+?)(:?-[\d]+|)\]')
COLOR_FORMAT = re.compile(r'\[([A-Z]+)\]')

TIME_REPLACES = {
    'DD': 'dd',
    'MM': 'm',
    'HH': 'H',
    'SS': 's',
    'AM/PM': 'a'
}

DATE_REPLACES = {
    'DD': 'dd',
    'YYYY': 'yyyy',
    'YY': 'yy',

}
# http://unicode.org/reports/tr35/tr35-dates.html#Date_Format_Patterns
FIX_BUILTIN_FORMATS = {
    14: 'MM-dd-yy',
    15: 'd-MMM-yy',
    16: 'd-MMM',
    17: 'MMM-yy',
    22: 'M/d/yy h:mm',
}


def normalize_date_format(_format):
    for f, to in DATE_REPLACES.items():
        _format = _format.replace(f, to)
    return _format


def normalize_time_format(_format):
    for f, to in TIME_REPLACES.items():
        _format = _format.replace(f, to)
    return _format.replace('\\', '')


def normalize_datetime_format(_format):
    parts = _format.split('\\')
    parts = [nf(p) for p, nf in zip(parts, [normalize_date_format, normalize_time_format])]
    return ''.join(parts)


class ColorNumberPattern(NumberPattern):
    def __init__(self, *args, **kwargs):
        super(ColorNumberPattern, self).__init__(*args, **kwargs)
        self.pos_color = self.neg_color = None
        try:
            self.pos_color = COLOR_FORMAT.findall(self.prefix[0])[0]
        except IndexError:
            pass
        try:
            self.neg_color = COLOR_FORMAT.findall(self.prefix[1])[0]
        except IndexError:
            pass

        self.prefix = [COLOR_FORMAT.sub('', p) for p in self.prefix]

    def apply(self, value, locale, currency=None, force_frac=None):
        formatted = super(ColorNumberPattern, self).apply(value, locale, currency, force_frac)
        return self.apply_color(value, formatted)

    def apply_color(self, value, formatted):
        if not isinstance(value, Decimal):
            value = Decimal(str(value))
        value = value.scaleb(self.scale)
        is_negative = int(value.is_signed())
        template = '<span style="color: {color}">{value}</span>'
        if is_negative and self.neg_color:
            return template.format(color=self.neg_color, value=formatted)
        if not is_negative and self.pos_color:
            return template.format(color=self.pos_color, value=formatted)
        return formatted


def parse_pattern(pattern):
    """Parse number format patterns"""
    if isinstance(pattern, NumberPattern):
        return pattern

    def _match_number(pattern):
        rv = number_re.search(pattern)
        if rv is None:
            raise ValueError('Invalid number pattern %r' % pattern)
        return rv.groups()

    # Do we have a negative subpattern?
    pattern = CLEAN_CURRENCY_RE.sub('\\1', pattern.replace('\\', ''))
    pattern = CLEAN_RE.sub('', pattern).strip().replace('_', ' ')
    if ';' in pattern:
        pattern, neg_pattern = pattern.split(';', 1)
        # maybe zero and text pattern
        neg_pattern = neg_pattern.split(';', 1)[0]
        pos_prefix, number, pos_suffix = _match_number(pattern)
        neg_prefix, _, neg_suffix = _match_number(neg_pattern)
        # TODO Do not remove from neg prefix
        neg_prefix = '-' + neg_prefix
    else:
        pos_prefix, number, pos_suffix = _match_number(pattern)
        neg_prefix = '-' + pos_prefix
        neg_suffix = pos_suffix

    if 'E' in number:
        number, exp = number.split('E', 1)
    else:
        exp = None
    if '@' in number:
        if '.' in number and '0' in number:
            raise ValueError('Significant digit patterns can not contain '
                             '"@" or "0"')
    if '.' in number:
        integer, fraction = number.rsplit('.', 1)
    else:
        integer = number
        fraction = ''

    def parse_precision(p):
        """Calculate the min and max allowed digits"""
        min = max = 0
        for c in p:
            if c in '@0':
                min += 1
                max += 1
            elif c == '#':
                max += 1
            elif c in ',. ':
                continue
            else:
                break
        return min, max

    int_prec = parse_precision(integer)
    frac_prec = parse_precision(fraction)
    if exp:
        frac_prec = parse_precision(integer + fraction)
        exp_plus = exp.startswith('+')
        exp = exp.lstrip('+')
        exp_prec = parse_precision(exp)
    else:
        exp_plus = None
        exp_prec = None
    grouping = parse_grouping(integer)
    number_pattern = ColorNumberPattern(pattern, (pos_prefix, neg_prefix),
        (pos_suffix, neg_suffix), grouping,
        int_prec, frac_prec,
        exp_prec, exp_plus)
    return number_pattern


def format_decimal(number, format=None, locale=LC_NUMERIC):
    u"""Return the given decimal number formatted for a specific locale.

    >>> format_decimal(1.2345, locale='en_US')
    u'1.234'
    >>> format_decimal(1.2346, locale='en_US')
    u'1.235'
    >>> format_decimal(-1.2346, locale='en_US')
    u'-1.235'
    >>> format_decimal(1.2345, locale='sv_SE')
    u'1,234'
    >>> format_decimal(1.2345, locale='de')
    u'1,234'

    The appropriate thousands grouping and the decimal separator are used for
    each locale:

    >>> format_decimal(12345.5, locale='en_US')
    u'12,345.5'

    :param number: the number to format
    :param format:
    :param locale: the `Locale` object or locale identifier
    """
    locale = Locale.parse(locale)
    if not format:
        format = locale.decimal_formats.get(format)
    pattern = parse_pattern(format)
    return pattern.apply(number, locale)


def format_cell(cell, locale=None):
    value = cell.value
    formatted_value = value or '&nbsp;'
    number_format = cell.number_format
    if not number_format:
        return formatted_value
    if isinstance(value, six.integer_types) or isinstance(value, float):
        if number_format.lower() != 'general':
            locale = locale or LC_NUMERIC
            formatted_value = format_decimal(value, number_format, locale=locale)

    locale = locale or LC_TIME

    # Possible problem with dd-mmm and more
    number_format = FIX_BUILTIN_FORMATS.get(cell._style.numFmtId, number_format)
    number_format = number_format.split(';')[0]

    if type(value) == datetime.date:
        number_format = normalize_date_format(number_format)

        formatted_value = format_date(value, number_format, locale=locale)
    elif type(value) == datetime.datetime:
        number_format = normalize_datetime_format(number_format)
        formatted_value = format_datetime(value, number_format, locale=locale)
    elif type(value) == datetime.time:
        number_format = normalize_time_format(number_format)
        formatted_value = format_time(value, number_format, locale=locale)
    return formatted_value
