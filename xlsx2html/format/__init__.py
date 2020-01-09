import datetime

import six
from babel.dates import LC_TIME
from babel.numbers import (
    LC_NUMERIC
)

from xlsx2html.constants import BUILTIN_FORMATS
from .dt import format_time, format_datetime, format_date
from .locale import extract_locale_from_format
from .number import format_decimal


def format_cell(cell, locale=None, f_cell=None):
    value = cell.value
    formatted_value = value or '&nbsp;'
    cell_format = cell.number_format
    if not cell_format:
        return format_hyperlink(formatted_value, cell.hyperlink)

    if isinstance(value, six.integer_types) or isinstance(value, float):
        if cell_format.lower() != 'general':
            locale = locale or LC_NUMERIC
            formatted_value = format_decimal(value, cell_format, locale=locale)

    locale = locale or LC_TIME

    # Possible problem with dd-mmm and more
    cell_format = BUILTIN_FORMATS.get(cell._style.numFmtId, cell_format)
    cell_format = cell_format.split(';')[0]

    new_locale, cell_format = extract_locale_from_format(cell_format)
    if new_locale:
        locale = new_locale

    if type(value) == datetime.date:
        formatted_value = format_date(value, cell_format, locale=locale)
    elif type(value) == datetime.datetime:
        formatted_value = format_datetime(value, cell_format, locale=locale)
    elif type(value) == datetime.time:
        formatted_value = format_time(value, cell_format, locale=locale)
    if cell.hyperlink:
        return format_hyperlink(formatted_value, cell)
    return formatted_value


def format_hyperlink(value, cell):
    hyperlink = cell.hyperlink
    if hyperlink is None or hyperlink.target is None:
        return value

    if hyperlink.location is not None:
        href = "{}#{}".format(hyperlink.target, hyperlink.location)
    else:
        href = hyperlink.target

    return '<a href="{href}">{value}</a>'.format(href=href, value=value)
