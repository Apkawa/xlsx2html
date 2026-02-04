import openpyxl
from packaging import version

OPENPYXL_24 = version.parse(openpyxl.__version__) < version.parse("2.5")

try:
    import warnings

    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", category=DeprecationWarning)
        from typing.io import BinaryIO
except ImportError:
    # python >= 3.13
    from typing import BinaryIO  # noqa: F401


try:
    from babel.numbers import number_re  # type: ignore[attr-defined]
except ImportError:
    from babel.numbers import _number_pattern_re as number_re  # noqa: F401
