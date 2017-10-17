# coding: utf-8
from __future__ import unicode_literals

import os
from unittest import TestCase
from xlsx2html.format import format_decimal

FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), 'fixtures')


class FormatTestCase(TestCase):
    maxDiff = None

    def setUp(self):
        pass

    def test_currency_format(self):
        format = u'#,##0.00\\ [$\u0440.-419];\\-#,##0.00\\ [$\u0440.-419]'
        self.assertEqual(
            format_decimal(1000, format=format, locale='ru'),
            '1\xa0000,00 р.')
        self.assertEqual(
            format_decimal(-1000, format=format, locale='ru'),
            '-1\xa0000,00 р.')
