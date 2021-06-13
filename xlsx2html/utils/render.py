from typing import Dict, Any

from xlsx2html.utils.style import StyleType

AttrsType = Dict[str, Any]


def render_attrs(attrs: AttrsType) -> str:
    if not attrs:
        return ""
    return " ".join(['%s="%s"' % a for a in sorted(attrs.items(), key=lambda a: a[0]) if a[1]])


def render_inline_styles(styles: StyleType) -> str:
    if not styles:
        return ""
    return ";".join(
        ["%s: %s" % (k, v) for k, v in sorted(styles.items(), key=lambda a: a[0]) if v]
    )
