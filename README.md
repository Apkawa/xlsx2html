[![Build Status](https://travis-ci.org/Apkawa/xlsx2html.svg?branch=master)](https://travis-ci.org/Apkawa/xlsx2html)
[![Coverage Status](https://coveralls.io/repos/github/Apkawa/xlsx2html/badge.svg?branch=master)](https://coveralls.io/github/Apkawa/xlsx2html?branch=master)
[![Requirements Status](https://requires.io/github/Apkawa/xlsx2html/requirements.svg?branch=master)](https://requires.io/github/Apkawa/django-multitype-file-field/requirements/?branch=master)
[![PyPI](https://img.shields.io/pypi/pyversions/xlsx2html.svg)]()

# xlsx2html

A simple export from xlsx format to html tables with keep cell formatting


# Install

```bash
pip install xlsx2html
```


# Usage

```python
from xlsx2html import xlsx2html

xlsx2html('path/to/example.xlsx', 'path/to/output.html')
```


# Contributors
```bash
pip install -r requirements.txt
pytest # run tests
tox # run test matrix
```

## publish
```bash
python setup.py sdist upload -r pypi
```
