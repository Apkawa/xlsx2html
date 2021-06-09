import re
from decimal import Decimal

from babel import Locale
from babel.numbers import NumberPattern, number_re, parse_grouping, LC_NUMERIC

ASTERISK_CLEAN_RE = re.compile(r"[*]")
QUESTION_MARK_RE = re.compile(r"\?")
UNDERSCORE_RE = re.compile(r"_.")
MINUS_CLEAN_RE = re.compile(r"[-]")
CLEAN_CURRENCY_RE = re.compile(r"\[\$(.+?)(:?-[\d]+|)\]")
COLOR_FORMAT = re.compile(r"\[([A-Z]+)\]", re.IGNORECASE)


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

        self.prefix = [COLOR_FORMAT.sub("", p) for p in self.prefix]

    def apply(self, value, locale, currency=None, force_frac=None):
        formatted = super(ColorNumberPattern, self).apply(
            value, locale, currency, force_frac
        )
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


class PatternParser:
    def __init__(self, pattern):
        self.general_pattern = None
        self.by_sign_pattern = None

        def _match_number(pattern):
            rv = number_re.search(pattern)
            if rv is None:
                raise ValueError("Invalid number pattern %r" % pattern)
            return rv.groups()

        def parse_precision(p):
            """Calculate the min and max allowed digits"""
            min = max = 0
            for c in p:
                if c in "@0":
                    min += 1
                    max += 1
                elif c == "#":
                    max += 1
                elif c in ",. ":
                    continue
                else:
                    break
            return min, max

        def generate_color_number_pattern(pattern, number, prefix, suffix):
            (grouping, int_prec, frac_prec, exp_prec, exp_plus) = handle_number(number)
            if number != "":
                return ColorNumberPattern(
                    pattern,
                    prefix,
                    suffix,
                    grouping,
                    int_prec,
                    frac_prec,
                    exp_prec,
                    exp_plus,
                )
            else:
                return pattern

        def handle_number(number):
            if "E" in number:
                number, exp = number.split("E", 1)
            else:
                exp = None
            if "@" in number:
                if "." in number and "0" in number:
                    raise ValueError(
                        "Significant digit patterns can not contain " '"@" or "0"'
                    )
            if "." in number:
                integer, fraction = number.rsplit(".", 1)
            else:
                integer = number
                fraction = ""

            int_prec = parse_precision(integer)
            frac_prec = parse_precision(fraction)
            if exp:
                frac_prec = parse_precision(integer + fraction)
                exp_plus = exp.startswith("+")
                exp = exp.lstrip("+")
                exp_prec = parse_precision(exp)
            else:
                exp_plus = None
                exp_prec = None
            grouping = parse_grouping(integer)
            return (grouping, int_prec, frac_prec, exp_prec, exp_plus)

        """Parse number format patterns"""
        if isinstance(pattern, NumberPattern):
            self.general_pattern = pattern
            return

        pattern = CLEAN_CURRENCY_RE.sub("\\1", pattern.replace("\\", ""))
        pattern = ASTERISK_CLEAN_RE.sub("", pattern).strip()
        pattern = UNDERSCORE_RE.sub(" ", pattern)
        pattern = QUESTION_MARK_RE.sub("", pattern)
        if ";" in pattern:
            pattern_parts = pattern.split(";")
            pos_pattern = pattern_parts[0]
            neg_pattern = pattern_parts[1]
            zero_pattern = pattern_parts[2] if len(pattern_parts) > 2 else pos_pattern
            # text_pattern = pattern_parts[3] if len(pattern_parts) > 3 else pos_pattern

            by_sign_pattern_list = []
            for _pattern in [pos_pattern, neg_pattern, zero_pattern]:
                prefix, number, suffix = _match_number(_pattern)
                color_number_pattern = generate_color_number_pattern(
                    _pattern, number, (prefix, prefix), (suffix, suffix)
                )
                by_sign_pattern_list.append(color_number_pattern)
            self.by_sign_pattern = tuple(by_sign_pattern_list)

        else:
            pattern = MINUS_CLEAN_RE.sub("", pattern)
            pos_prefix, number, pos_suffix = _match_number(pattern)
            neg_prefix = "-" + pos_prefix
            neg_suffix = pos_suffix

            self.general_pattern = generate_color_number_pattern(
                pattern, number, (pos_prefix, neg_prefix), (pos_suffix, neg_suffix)
            )

    def apply(self, number, locale):
        if self.general_pattern:
            pattern = self.general_pattern
        else:
            pos_pattern, neg_pattern, zero_pattern = self.by_sign_pattern

            pattern = zero_pattern
            if number > 0:
                pattern = pos_pattern
            elif number < 0:
                pattern = neg_pattern

        # in case there are no number included - pattern will be string
        if isinstance(pattern, str):
            return pattern
        else:
            return pattern.apply(number, locale)


def format_decimal(number, format=None, locale=LC_NUMERIC):
    """Return the given decimal number formatted for a specific locale.

    >>> format_decimal(1.2345, locale='en_US')
    '1.234'
    >>> format_decimal(1.2346, locale='en_US')
    '1.235'
    >>> format_decimal(-1.2346, locale='en_US')
    '-1.235'
    >>> format_decimal(1.2345, locale='sv_SE')
    '1,234'
    >>> format_decimal(1.2345, locale='de')
    '1,234'

    The appropriate thousands grouping and the decimal separator are used for
    each locale:

    >>> format_decimal(12345.5, locale='en_US')
    '12,345.5'

    :param number: the number to format
    :param format:
    :param locale: the `Locale` object or locale identifier
    """
    locale = Locale.parse(locale)
    if not format:
        format = locale.decimal_formats.get(format)
    pattern = PatternParser(format)
    return pattern.apply(number, locale)
