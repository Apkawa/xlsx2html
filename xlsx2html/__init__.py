# -*- coding: utf-8 -*-
import warnings
from typing import TextIO

from .core import xlsx2html, XLSX2HTMLConverter

__all__ = ["xls2html", "xlsx2html", "XLSX2HTMLConverter"]

__version__ = "0.4.0"


def xls2html(*args, **kwargs) -> TextIO:  # type: ignore
    warnings.warn("This func was renamed to xlsx2html.", DeprecationWarning)
    return xlsx2html(*args, **kwargs)
