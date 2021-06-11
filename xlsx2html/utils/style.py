from collections import defaultdict
from typing import Dict, Union, Optional, List

BORDER_DIRECTIONS = ["right", "left", "top", "bottom"]
BORDER_PROPS = ["width", "style", "color"]

StyleType = Dict[str, Union[str, None]]
BorderType = Union[str, StyleType]


def border_make_shorthand(style: StyleType) -> StyleType:
    """
    >>> border_make_shorthand({"border-color": "#000", "border-width": "1px", "border-style": "solid"})
    {'border': '1px solid #000'}
    >>> border_make_shorthand({"border-width": "1px", "border-style": "solid"})
    {'border': '1px solid'}

    >>> border_make_shorthand({"border-color": "#000", "border-style": "solid"})
    {'border': 'solid #000'}


    >>> border_make_shorthand({ \
        'border-right-color': '#000', 'border-right-width': '1px', 'border-right-style': 'solid', \
        'border-left-color': '#000', 'border-left-width': '1px', 'border-left-style': 'solid', \
        })
    {'border-right': '1px solid #000', 'border-left': '1px solid #000'}

    >>> border_make_shorthand({ \
        'border-right-color': '#000', 'border-right-width': '1px', 'border-right-style': 'solid', \
        'border-left-color': '#000', 'border-left-width': '1px', 'border-left-style': 'solid', \
        'border-top-color': '#000', 'border-top-width': '1px', 'border-top-style': 'solid', \
        'border-bottom-color': '#000', 'border-bottom-width': '1px', 'border-bottom-style': 'solid', \
        })
    {'border': '1px solid #000'}

    >>> border_make_shorthand({ \
        'border-right-color': '#001', 'border-right-width': '1px', 'border-right-style': 'solid', \
        'border-left-color': '#000', 'border-left-width': '1px', 'border-left-style': 'solid', \
        'border-top-color': '#000', 'border-top-width': '1px', 'border-top-style': 'solid', \
        'border-bottom-color': '#000', 'border-bottom-width': '1px', 'border-bottom-style': 'solid', \
        })
    {'border-right': '1px solid #001', 'border-left': '1px solid #000', 'border-top': '1px solid #000', 'border-bottom': '1px solid #000'}

    >>> border_make_shorthand({'border-right': '1px solid #000', 'border-left': '1px solid #000', \
        'border-top': '1px solid #000', 'border-bottom': '1px solid #000'})
    {'border': '1px solid #000'}


    >>> border_make_shorthand({"border-right-color": "red"})
    {'border-right-color': 'red'}

    >>> border_make_shorthand({"border-right-style": "solid", \
        'foo-bar': '1', 'border-radius': '2px'})
    {'foo-bar': '1', 'border-radius': '2px', 'border-right': 'solid'}

    :param style:
    :return:
    """
    new_style = dict(style)
    border_styles: Dict[str, List[str]] = defaultdict(list)
    border_names = set()

    for d in [""] + BORDER_DIRECTIONS:
        for bp in [""] + BORDER_PROPS:
            _bs_name = "-".join(filter(None, ["border", d, bp]))
            _bs = new_style.pop(_bs_name, None)
            if _bs:
                border_names.add(_bs_name)
                border_styles[d].append(_bs)

    if (
        len(border_styles) == 4
        and len(set([repr(v) for v in border_styles.values()])) == 1
    ):
        border_styles = {"": list(border_styles.values())[0]}

    for d, v in border_styles.items():
        _bs_name = "-".join(filter(None, ["border", d]))
        if len(v) == 1 and f"{_bs_name}-color" in border_names:
            _bs_name = f"{_bs_name}-color"
        new_style[_bs_name] = " ".join(v)

    return new_style


def compress_style(style: StyleType) -> StyleType:
    style = border_make_shorthand(style)
    return style
