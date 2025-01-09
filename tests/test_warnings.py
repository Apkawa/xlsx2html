import warnings

import openpyxl
import pytest
from _pytest.warning_types import PytestDeprecationWarning
from _pytest.recwarn import WarningsRecorder

from tests.test_files import XLSX_FILE
from xlsx2html import xlsx2html


@pytest.mark.filterwarnings("error")
def test_simply_warnings(temp_file):
    from xlsx2html import xlsx2html

    out_file = temp_file()
    from tests.test_files import XLSX_FILE

    xlsx2html(XLSX_FILE, out_file, locale="en").close()  # Can use index
    with open(out_file) as of:
        assert of.read().count("</table>")
    xlsx2html(XLSX_FILE, out_file, locale="en").close()  # Can use index


@pytest.mark.filterwarnings("error")
def test_sheet_by_name_no_deprecation_warning(temp_file):
    # By names
    out_file = temp_file()
    wb = openpyxl.load_workbook(XLSX_FILE)
    xlsx2html(
        XLSX_FILE, out_file, locale="en", sheet=wb.sheetnames
    ).close()  # Can use index


@pytest.mark.filterwarnings("error")
def test_count_warnings(temp_file):
    # Deep check warnings. Not necessary but funny
    with warnings.catch_warnings():
        # Mute pytest.PytestDeprecationWarning: A private pytest class or function was used.
        # because WarningsRecorder is private
        warnings.filterwarnings("ignore", category=PytestDeprecationWarning)

        with WarningsRecorder() as record:
            from xlsx2html import xlsx2html

            out_file = temp_file()
            from tests.test_files import XLSX_FILE

            xlsx2html(XLSX_FILE, out_file, locale="en").close()  # Can use index
            with open(out_file) as of:
                assert of.read().count("</table>")
            assert len(record.list) == 0, "\n".join(map(str, record.list))
