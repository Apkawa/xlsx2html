# coding: utf-8
import io
import os

import pytest
from openpyxl import load_workbook

from xlsx2html.core import xlsx2html, XLSX2HTMLConverter

FIXTURES_ROOT = os.path.join(os.path.dirname(__file__), "fixtures")


def get_fixture(name):
    return os.path.join(FIXTURES_ROOT, name)


XLSX_FILE = get_fixture("example.xlsx")


def test_xlsx2html(temp_file):
    out_file = temp_file()
    xlsx2html(XLSX_FILE, out_file, locale="en")
    result_html = open(out_file).read()
    assert result_html


def test_xlsx2html_benchmarked(benchmark):
    out_file = benchmark(xlsx2html, XLSX_FILE, locale="en")
    out_file.seek(0)
    result_html = out_file.read()
    assert result_html


def test_use_streams():
    xlsx_file = open(XLSX_FILE, "rb")
    out_file = io.StringIO()
    xlsx2html(xlsx_file, out_file, locale="en")
    assert out_file.tell() > 0
    out_file.seek(0)
    result_html = out_file.read()
    assert result_html


def test_default_output():
    xlsx_file = open(XLSX_FILE, "rb")
    out_file = xlsx2html(xlsx_file, locale="en")
    assert out_file.tell() > 0
    out_file.seek(0)
    result_html = out_file.read()
    assert result_html


def test_use_workbook_instance():
    xlsx_file = load_workbook(XLSX_FILE)
    out_file = xlsx2html(xlsx_file)
    assert out_file.tell() > 0
    out_file.seek(0)
    result_html = out_file.read()
    assert result_html


def test_use_workbook_instance_with_parse_formula():
    xlsx_file = load_workbook(XLSX_FILE)
    with pytest.raises(ValueError) as e:
        xlsx2html(xlsx_file, parse_formula=True)
    assert (
        str(e.value) == "for parse_formula must be set "
        "`formula_wb=openpyxl.load_workbook(filepath, data_only=False)`"
    )

    converter = XLSX2HTMLConverter(
        filepath=load_workbook(XLSX_FILE, data_only=True),
        parse_formula=True,
        formula_wb=load_workbook(XLSX_FILE, data_only=False),
    )
    out_file = converter.get_html_stream()
    assert out_file.tell() > 0
    out_file.seek(0)
    result_html = out_file.read()
    assert result_html


def test_hyperlink(temp_file):
    out_file = temp_file()
    xlsx2html(get_fixture("hyperlinks.xlsx"), out_file, locale="en", parse_formula=True)
    result_html = open(out_file).read()
    assert result_html


@pytest.mark.webtest()
def test_screenshot_diff(temp_file, browser, screenshot_regression, fixture_file):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()
    xlsx2html(XLSX_FILE, out_file, locale="en")

    browser.visit("file://" + out_file)
    screenshot_regression()
