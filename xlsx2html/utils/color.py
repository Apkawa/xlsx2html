from typing import Optional

from openpyxl.styles.colors import COLOR_INDEX, aRGB_REGEX, Color


def normalize_color(color: Color) -> Optional[str]:
    # TODO RGBA
    rgb = None
    if color.type == "rgb":
        rgb = color.rgb
    if color.type == "indexed":
        try:
            rgb = COLOR_INDEX[color.indexed]
        except IndexError:
            # The indices 64 and 65 are reserved for the system
            # foreground and background colours respectively
            pass
        if not rgb or not aRGB_REGEX.match(rgb):
            # TODO system fg or bg
            rgb = "00000000"
    if rgb:
        return "#" + rgb[2:]
    return None
