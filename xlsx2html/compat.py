import openpyxl
from packaging import version

OPENPYXL_24 = version.parse(openpyxl.__version__) < version.parse("2.5")
