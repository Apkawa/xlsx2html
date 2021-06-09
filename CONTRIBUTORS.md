# Run tests
```bash
pip install -r requirements.txt
pytest # run tests
tox # run test matrix
```

# Run tests with pyenv with specific python and pypy

```
pyenv install 3.10-dev pypy3.7-7.3.5
pyenv local 3.10-dev pypy3.7-7.3.5
pip install -r requirements.txt
tox -e py310,pypy3
```


# Before commit

```
pre-commit install
```

For pycharm needs install `tox` to global
