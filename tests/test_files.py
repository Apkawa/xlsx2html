# coding: utf-8
from __future__ import unicode_literals

import os

import pytest

from xlsx2html.core import xlsx2html

FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), 'fixtures')

XLSX_FILE = os.path.join(FIXTURES_ROOT, 'example.xlsx')


def test_xlsx2html(temp_file):
    out_file = temp_file()
    xlsx2html(XLSX_FILE, out_file, locale='en')
    result_html = open(out_file).read()
    assert result_html


@pytest.mark.webtest()
def test_screenshot_diff(temp_file, browser, screenshot_match):
    out_file = temp_file()
    xlsx2html(XLSX_FILE, out_file, locale='en')
    browser.visit('file://' + out_file)
    screenshot_match()
