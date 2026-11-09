from html import escape


class SafeHTML(str):
    """HTML generated in xlsx2html with already escaped user-supplied values. Works similarly to Django SafeString."""

    __slots__ = ()


def escape_html(value, accept_fragments: bool = True) -> SafeHTML:
    if isinstance(value, SafeHTML) and accept_fragments:
        return value
    return SafeHTML(escape(str(value), quote=True))
