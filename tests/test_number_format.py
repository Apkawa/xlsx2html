import pytest

from xlsx2html.format import format_decimal

test_formats = [
    ('00000', [
        (983, '00983'),
        (25983, '25983')
    ]),
    ('0.00', [
        (983, '983.00'),
        (5.5235, '5.52'),
    ]),
    ('0;-0;;@', [
        (98, '98'),
        (0, ''),
        (-9, '-9'),
        ('text', 'text')
    ]),
    (';;;@', [
        (90, ''),
        (0, ''),
        (-9, ''),
        ('test', 'test'),
    ]),
    # Lines numbers up with the decimal.
    ('#.???', [
        (3.256, '3.256'),
        (3.25, '3.25'),
        (3.256356, '3.256'),
    ]),
    # Displays numbers in thousands.
    ('#,', [
        (4058.50, '4'),
        (42058, '42'),
        (420, ''),
        (999550, '1000')
    ]),
    # Displays numbers in millions.
    ('#,###,, "M"', [
        (32654236, '33 M'),
        (1.2357E+12, '1,235,699 M')
    ]),
    # Represents numbers in millions.
    ('0.00,,', [
        (1000000, '1.00'),
        (12200000, '12.20'),
        (120000, '0.12'),
    ]),
    # Displays “Error!” in red for negative numbers
    ('0;[Red]”Error!”;0;[Red]”Error!”', [
        (10, '10'),
        ('hello', 'Error!'),
        (-10, 'Error!'),
        (0, '0')
    ]),
    # Displays the negative sign on the right side
    ('0.0_-;0.0-', [
        (-5, '5.0-')
    ])
]


@pytest.mark.parametrize('fmt,expects', test_formats)
def test_decimal_format(fmt, expects):
    for _input, _expect in expects:
        assert format_decimal(_input, fmt, locale='en') == _expect
