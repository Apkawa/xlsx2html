# Run tests
```bash
pip install -r requirements.txt
pytest # run tests
python -m tox # run test matrix
```

# Run tests with pyenv with specific python and pypy

```
pyenv install 3.12 pypy7.3.5
pyenv local 3.12 pypy7.3.5
pip install -r requirements.txt
python -m tox -e py312,pypy7
```


# Before commit

```
pre-commit install
```

For pycharm needs install `tox` to global

