import openpyxl
from packaging import version

OPENPYXL_24 = version.parse(openpyxl.__version__) < version.parse("2.5")

try:
    from typing.io import BinaryIO
except ImportError:
    # python >= 3.13
    from typing import BinaryIO  # noqa: F401
