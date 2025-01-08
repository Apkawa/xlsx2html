import pytest

from xlsx2html import xlsx2html
from xlsx2html.utils.image import bytes_to_datauri


def test_to_datauri(fixture_file):
    path = fixture_file("img.png")
    with open(path, "rb") as fp:
        res = bytes_to_datauri(fp, path)
    assert res == (
        "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABcAAAAXBAMAAAASBMmTAA"
        "AALVBMVEVHcEzycDbwcDfybTTzbjXybjXybjXybjT0bzbnajHzbjTybjXybjXybjTybj"
        "WllWobAAAADnRSTlMAUB8SjEP3mjAKalzJ2RSEbQUAAACsSURBVBjTY2BgDhQEg2AGIH"
        "B7BwUODAwsMPa7JwwMZlDmitZ3BgxxUI4A07sABjkERwDOcZSCcZ5BJMGcJwYmUM4L1X"
        "cFDAx1T1VBnMfc7xQYGPpegw14zPUM6A6+VxAOI9A2Br2HEA4zyFV1yRAOQ90TA/V3Dl"
        "DOHKA9byZAOZxu7555MkA5DMzSQJ+BOHEvGCCAFehqu7dBSmAQ+8wA1acMflD2M6BtyK"
        "EDAN2RhY1kOYaNAAAAAElFTkSuQmCC"
    )


@pytest.mark.webtest()
def test_image(temp_file, browser, screenshot_regression, fixture_file):
    browser.driver.set_window_size(1280, 1024)
    out_file = temp_file()
    xlsx2html(fixture_file("image.xlsx"), out_file, locale="en", parse_formula=True)
    browser.visit("file://" + out_file)
    print('Window size', browser.driver.get_window_size())
    screenshot_regression()
