[![PyPi](https://img.shields.io/pypi/v/xlsx2html.svg)](https://pypi.python.org/pypi/xlsx2html)
[![Build Status](https://travis-ci.org/Apkawa/xlsx2html.svg?branch=master)](https://travis-ci.org/Apkawa/xlsx2html)
[![Codecov](https://codecov.io/gh/Apkawa/xlsx2html/branch/master/graph/badge.svg)](https://codecov.io/gh/Apkawa/xlsx2html)
[![Requirements Status](https://requires.io/github/Apkawa/xlsx2html/requirements.svg?branch=master)](https://requires.io/github/Apkawa/xlsx2html/requirements/?branch=master)
[![PyUP](https://pyup.io/repos/github/Apkawa/xlsx2html/shield.svg)](https://pyup.io/repos/github/Apkawa/xlsx2html)
[![Python versions](https://img.shields.io/pypi/pyversions/xlsx2html.svg)](sd)
[![License](https://img.shields.io/badge/license-MIT-blue.svg)](LICENSE)

# xlsx2html

A simple export from xlsx format to html tables with keep cell formatting


## Install

* From pip
```bash
pip install xlsx2html
```

* From github dev version
```
pip install -e "git+https://github.com/Apkawa/xlsx2html.git#egg=xlsx2html"
```

### Compatibly

| Python  | xlsx2html  |
|---|---|
| 2.7  | 0.1.10  |
| 3.5  | 0.2.1  |
| >=3.6 | latest  |



## Usage

### Simple usage

```python
from xlsx2html import xlsx2html

out_stream = xlsx2html('path/to/example.xlsx')
out_stream.seek(0)
print(out_stream.read())

```

pass output file
```python
from xlsx2html import xlsx2html

xlsx2html('path/to/example.xlsx', 'path/to/output.html')
```

use file like objects

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

from shell

```bash
python -m xlsx2html path/to/example.xlsx path/to/output.html
```

### Advanced usage

Use converter class:

```python
from xlsx2html import XLSX2HTMLConverter
converter = XLSX2HTMLConverter(
    filepath='path/to/example.xlsx',
    locale='de_DE',
    parse_formula=True,
    inline_styles=False
)
html = converter.get_html(sheet="sheet name")
```

Export sheet to only table:

```python
from xlsx2html import XLSX2HTMLConverter
converter = XLSX2HTMLConverter(
    filepath='path/to/example.xlsx',
    locale='de_DE',
    parse_formula=True,
    inline_styles=False
)
result = converter.get_table(sheet="sheet name", extra_attrs={'id': 'table_id'})

print(f"""
<html>
    <head>
    <style type="text/css">
        {result.css}
    </style>
    </head>
    <body>
        {result.html}
    </body>
</html>""")
```

Export all sheets:

```python
from xlsx2html import XLSX2HTMLConverter
converter = XLSX2HTMLConverter(
    filepath='path/to/example.xlsx',
    locale='de_DE',
    parse_formula=True,
    inline_styles=False
)
results = converter.get_tables(extra_attrs={'class': 'xlsx_sheet'})

css_str = '\n'.join([r.css for r in results])
tables_str = '\n'.join([r.html for r in results])

print(f"""
<html>
    <head>
    <style type="text/css">
        {css_str}
    </style>
    </head>
    <body>
        {tables_str}
    </body>
</html>""")
```
