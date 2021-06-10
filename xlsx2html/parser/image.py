from dataclasses import dataclass

from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.drawing.image import Image as WSImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker
from openpyxl.utils import units

from xlsx2html.utils.image import bytes_to_datauri


@dataclass
class Point:
    x: int
    y: int


@dataclass
class ImageInfo:
    col: int
    row: int
    offset: Point
    width: int
    height: int
    src: str

    @classmethod
    def from_ws_image(cls, ws_image: WSImage) -> "ImageInfo":
        image = ws_image
        _from: AnchorMarker = image.anchor._from
        graphicalProperties: GraphicalProperties = image.anchor.pic.graphicalProperties
        transform = graphicalProperties.transform
        # http://officeopenxml.com/drwSp-location.php
        offsetX = units.EMU_to_pixels(_from.colOff)
        offsetY = units.EMU_to_pixels(_from.rowOff)

        return cls(
            col=_from.col + 1,
            row=_from.row + 1,
            offset=Point(x=offsetX, y=offsetY),
            width=units.EMU_to_pixels(transform.ext.width),
            height=units.EMU_to_pixels(transform.ext.height),
            src=bytes_to_datauri(image.ref, image.path),
        )
