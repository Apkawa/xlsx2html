[![PyPi](https://img.shields.io/pypi/v/xlsx2html.svg)](https://pypi.python.org/pypi/xlsx2html)
[![ci](https://github.com/Apkawa/xlsx2html/actions/workflows/ci.yml/badge.svg)](https://github.com/Apkawa/xlsx2html/actions/workflows/ci.yml)
[![Codecov](https://codecov.io/gh/Apkawa/xlsx2html/branch/master/graph/badge.svg)](https://codecov.io/gh/Apkawa/xlsx2html)
[![Python versions](https://img.shields.io/pypi/pyversions/xlsx2html.svg)]()
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

# xlsx2html

A simple export from xlsx format to html tables with keep cell formatting


# Install

```bash
pip install xlsx2html
```


# Usage
Simple usage
```python
from xlsx2html import xlsx2html

out_stream = xlsx2html('path/to/example.xlsx')
out_stream.seek(0)
print(out_stream.read())

```

or pass filepath
```python
from xlsx2html import xlsx2html

xlsx2html('path/to/example.xlsx', 'path/to/output.html')
```
or use file like objects

```python
import io
from xlsx2html import xlsx2html

# must be binary mode
xlsx_file = open('path/to/example.xlsx', 'rb')
out_file = io.StringIO()
xlsx2html(xlsx_file, out_file, locale='en')
out_file.seek(0)
result_html = out_file.read()
```

or from shell

```bash
python -m xlsx2html path/to/example.xlsx path/to/output.html
```
