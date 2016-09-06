# coding: utf-8
from __future__ import unicode_literals

import tempfile

import os
from unittest import TestCase

FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), 'fixtures')

from xlsx2html.core import xls2html


class XLS2HTMLTestCase(TestCase):
    maxDiff = None
    def setUp(self):
        self.xlsx_file = os.path.join(FIXTURES_ROOT, 'example.xlsx')
        self.expect_result = open(os.path.join(FIXTURES_ROOT, 'example.html')).read()
        self.tmp_file = tempfile.mktemp(suffix='xls2html_')

    def test_xls2html(self):
        xls2html(self.xlsx_file, self.tmp_file)
        result_html = open(self.tmp_file).read()
        self.assertEqual(result_html, self.expect_result)


    def tearDown(self):
        if os.path.exists(self.tmp_file):
            os.unlink(self.tmp_file)
