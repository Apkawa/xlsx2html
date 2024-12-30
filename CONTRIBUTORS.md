# Run tests
```bash
pip install -r requirements.txt
pytest # run tests
tox # run test matrix
```

# Run tests with pyenv with specific python and pypy

1. Ensure `VIRTUALENV_DISCOVERY=pyenv` in env https://github.com/un-def/virtualenv-pyenv/blob/master/README.md


```
pyenv install 3.12 pypy3.7-7.3.5
pyenv local 3.12 pypy3.7-7.3.5
pip install -r requirements.txt
tox -e py312,pypy3
```


# Before commit

```
pre-commit install
```

For pycharm needs install `tox` to global
