import pytest

from xlsx2html.parser.parser import XLSXParser
from xlsx2html.render.html import HtmlRenderer


def test_generic(fixture_file):
    parser = XLSXParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer()

    html = render.render(result)
    assert html


@pytest.mark.webtest()
def test_render_example(fixture_file, temp_file, browser, screenshot_regression):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()

    parser = XLSXParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer(display_grid=True, inline_styles=True)

    html = render.render(result)
    assert html

    with open(out_file, "w") as f:
        f.write(html)

    browser.visit("file://" + out_file)
    screenshot_regression()


@pytest.mark.webtest()
def test_optimize_style(fixture_file, temp_file, browser, screenshot_regression):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()

    parser = XLSXParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer(
        display_grid=True, inline_styles=False, default_border_style="0.5px solid gray"
    )

    html = render.render(result)
    assert html

    with open(out_file, "w") as f:
        f.write(html)

    browser.visit("file://" + out_file)
    screenshot_regression()


@pytest.mark.webtest()
def test_render_sheet2(fixture_file, temp_file, browser, screenshot_regression):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()

    parser = XLSXParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet(1)

    render = HtmlRenderer(display_grid=True, inline_styles=False)

    html = render.render(result)
    assert html

    with open(out_file, "w") as f:
        f.write(html)

    browser.visit("file://" + out_file)
    screenshot_regression()


@pytest.mark.webtest()
def test_render_conditional(fixture_file, temp_file, browser, screenshot_regression):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()

    parser = XLSXParser(filepath=fixture_file("conditional.xlsx"), parse_formula=True)
    result = parser.get_sheet()

    render = HtmlRenderer(display_grid=True, inline_styles=False)

    html = render.render(result)
    assert html

    with open(out_file, "w") as f:
        f.write(html)

    browser.visit("file://" + out_file)
    screenshot_regression()
