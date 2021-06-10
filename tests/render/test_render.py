import pytest

from xlsx2html.parser.parser import WBParser
from xlsx2html.render.html import HtmlRenderer


def test_generic(fixture_file):
    parser = WBParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer()

    html = render.render(result)
    assert html


@pytest.mark.webtest()
def test_render_example(fixture_file, temp_file, browser, screenshot_regression):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()

    parser = WBParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer(display_grid=True)

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

    parser = WBParser(filepath=fixture_file("example.xlsx"))
    result = parser.get_sheet()

    render = HtmlRenderer(
        display_grid=True,
        optimize_styles=True,
        default_border_style={"style": "solid", "color": "gray", "width": "0.5px"},
    )

    html = render.render(result)
    assert html

    with open(out_file, "w") as f:
        f.write(html)

    browser.visit("file://" + out_file)
    screenshot_regression()
