import os
import tempfile

import pytest
from diffimg import diff

from tests.pytest_helpers import get_test_info, build_filename


@pytest.fixture(scope='session')  # pragma: no cover
def splinter_screenshot_dir(request):
    """Browser screenshot directory."""
    if request.config.option.splinter_screenshot_dir.strip('.'):
        return os.path.abspath(request.config.option.splinter_screenshot_dir)
    return os.path.join(os.path.dirname(__file__), '.tests/splinter')


@pytest.fixture(scope='session')
def screenshot_reference_dir():
    return os.path.join(os.path.dirname(__file__), 'screenshots')


@pytest.fixture(scope="function")
def screenshot_match(browser, request, screenshot_reference_dir, splinter_screenshot_dir):
    def _factory(suffix='', threshold=0.00):
        test_info = get_test_info(request)
        reference_dir = os.path.join(screenshot_reference_dir, test_info['classname'])
        screenshot_dir = os.path.join(splinter_screenshot_dir, test_info['classname'])
        reference_name = os.path.join(reference_dir,
                                      build_filename(test_info['test_name'],
                                                     suffix=suffix + '-reference.png'))
        screenshot_name = os.path.join(screenshot_dir,
                                       build_filename(test_info['test_name'],
                                                      suffix=suffix + '-screenshot.png'))
        diff_name = os.path.join(screenshot_dir, build_filename(test_info['test_name'],
                                                                suffix=suffix + '-diff.png'))

        if not os.path.exists(reference_name):
            if not os.path.exists(reference_dir):
                os.makedirs(reference_dir)
            browser.driver.save_screenshot(reference_name)
            return

        if not os.path.exists(screenshot_dir):
            os.makedirs(screenshot_dir)
        browser.driver.save_screenshot(screenshot_name)
        diff_ratio = diff(reference_name, screenshot_name,
                          delete_diff_file=False,
                          diff_img_file=diff_name,
                          ignore_alpha=True)
        assert diff_ratio <= threshold, "Image not equals!"

        for f in [screenshot_name, diff_name]:
            if os.path.exists(f):
                os.unlink(f)

    yield _factory


@pytest.fixture(scope='function')
def temp_file(request):
    temp_files = []

    def tempfile_factory(extension='.html', prefix='xlsx2html_'):
        tf = tempfile.mktemp(suffix=extension, prefix='xlsx2html_')
        temp_files.append(tf)
        return tf

    yield tempfile_factory

    for tf in temp_files:
        if os.path.exists(tf):
            os.unlink(tf)


@pytest.fixture(scope='session')
def splinter_webdriver(request):
    return request.config.option.splinter_webdriver or 'chrome'


@pytest.fixture(scope='session')
def splinter_webdriver_executable(request, splinter_webdriver):
    """Webdriver executable directory."""
    executable = request.config.option.splinter_webdriver_executable
    if not executable and splinter_webdriver == 'chrome':
        from chromedriver_binary import chromedriver_filename
        executable = chromedriver_filename
    return os.path.abspath(executable) if executable else None


def pytest_addoption(parser):
    parser.addoption(
        '--skip-webtest',
        action='store_true',
        dest="skip_webtest",
        default=False,
        help="skip marked webtest tests")


def pytest_configure(config):
    mark_expr = []

    if config.option.markexpr:
        mark_expr.append(config.option.markexpr)

    if config.option.skip_webtest:
        mark_expr.append('not webtest')
    if mark_expr:
        setattr(config.option, 'markexpr', ' and '.join(mark_expr))
