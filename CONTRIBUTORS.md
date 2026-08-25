# Install

```bash
uv sync
```

# QA

QA stack (uv / ruff / mypy / pytest / tox + tox-uv / pre-commit / bump-my-version):

```bash
uv run tox                        # lint, format, type, py310-py313
uv run pytest --skip-webtest      # test suite without browser tests
uv run mypy xlsx2html             # type check (basic; see [tool.mypy] TODO)
uv run ruff check                 # lint
uv run ruff format                # format
uv run pre-commit run --all-files
uv run bump-my-version bump minor|patch   # zero versioning 0.{minor}.{patch}
```

Sandbox (writable cache):

```bash
export UV_CACHE_DIR=$PWD/.uv-cache UV_LINK_MODE=clone UV_PYTHON_INSTALL_DIR=$PWD/.uv-python
export PRE_COMMIT_HOME=$PWD/.pre-commit-home
```

## Browser tests (webtest)

Manual, outside `env_list` (playwright + chromedriver):

### Setup

- `uv sync --group webtest`
- `uv run playwright install --with-deps`

### Run webtest

```bash
uv run tox -e webtest
```

## wsl2 (may be obsolete)

With wsl2 need additional steps

For Ubuntu-24.04

1) install chrome deps
```bash
sudo apt-get update
sudo apt-get install -y curl unzip xvfb libxi6 libgconf-2-4
```
2) Install chrome-stable
```bash
wget https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb
sudo apt install ./google-chrome-stable_current_amd64.deb
```

3) Ensure chromedriver avaiable in path

ref: https://www.gregbrisebois.com/posts/chromedriver-in-wsl2/

# Before commit

```bash
pre-commit install
```
