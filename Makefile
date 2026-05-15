PY := ./venv/Scripts/python.exe
export PYTHONDONTWRITEBYTECODE := 1

.PHONY: help install check test cli clean

help:
	@echo "Available targets:"
	@echo "  make install     Install package and dev dependencies"
	@echo "  make check       Syntax check without starting Excel"
	@echo "  make test        Run unit tests without starting Excel"
	@echo "  make cli         Run CLI; pass args with ARGS=\"...\""
	@echo "  make clean       Remove generated Python/test/build artifacts"

install:
	$(PY) -m pip install -e ".[dev]"

check:
	$(PY) -m py_compile pymacros/*.py pymacros_cli/*.py tests/*.py

test:
	$(PY) -m pytest tests

cli:
	$(PY) -m pymacros_cli $(ARGS)

clean:
	rm -rf __pycache__ .pytest_cache build dist *.egg-info
	rm -rf pymacros/__pycache__ pymacros_cli/__pycache__ tests/__pycache__ procedures/__pycache__
