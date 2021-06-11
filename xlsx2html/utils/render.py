from xlsx2html.utils.style import StyleType


def render_attrs(attrs) -> str:
    if not attrs:
        return ""
    return " ".join(
        [
            '%s="%s"' % a
            for a in sorted(attrs.items(), key=lambda a: a[0])
            if a[1] is not None
        ]
    )


def render_inline_styles(styles) -> str:
    if not styles:
        return ""
    return ";".join(
        [
            "%s: %s" % a
            for a in sorted(styles.items(), key=lambda a: a[0])
            if a[1] is not None
        ]
    )
