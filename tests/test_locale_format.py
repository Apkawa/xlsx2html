from xlsx2html.format.locale import extract_locale_from_format


def test_extract_locale_format():
    fmt = "[$-404]e/m/d"

    locale, fmt = extract_locale_from_format(fmt)
    assert locale == "zh_Hant_TW"
    assert fmt == "e/m/d"
