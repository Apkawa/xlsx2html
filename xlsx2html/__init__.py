# -*- coding: utf-8 -*-
import warnings
from .core import xlsx2html

__all__ = ["xls2html", "xlsx2html"]

__version__ = "0.6.2"


def xls2html(*args, **kwargs):
    warnings.warn("This func was renamed to xlsx2html.", DeprecationWarning)
    return xlsx2html(*args, **kwargs)
