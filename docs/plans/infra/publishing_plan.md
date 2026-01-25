# Plan: Publish ExcelTamer as a Python Package

This document outlines the steps to package `ExcelTamer` for distribution via PyPI.

## 1. Project Structure Analysis

Current structure:
```
ExcelTamer/
  ExcelTamer/         # Source package
    __init__.py
    ExcelTamer.py
    ExcelAutomation.py
    ExcelTamerAgent/
      ...
  test/               # Tests
  requirements.txt    # Dependencies
  LICENSE
  README.md
```

## 2. Packaging Strategy
We will use modern Python packaging standards with `pyproject.toml` and `setuptools`.

### Key Components:
- **Build Backend**: `setuptools.build_meta`
- **Configuration**: `pyproject.toml` (replacing `setup.py`)
- **Version Management**: We will use semantic versioning, starting with `0.1.0`.

## 3. Implementation Steps

### 3.1. Prepare Source Code
- [x] ensure `ExcelTamer/__init__.py` exposes the main classes (`ExcelAutomation`, `ExcelTamer`, etc.) to make imports cleaner for users.
  - Current state: Empty.
  - Desired state: 
    - `from .ExcelAutomation import ExcelAutomation`
    - `from .ExcelTamerAgent import ExcelTamerTools` (and/or specific tools if needed)
    - `from .ExcelTamerAgent.AgentBuilder import create_agent`
    - This ensures `ExcelTamerTools` is easily usable as requested.

### 3.2. create `pyproject.toml`
Create a `pyproject.toml` file in the root directory with the following configuration:

```toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "ExcelTamer"
version = "0.1.0"
description = "An agentic tool for Excel automation using LLMs"
readme = "README.md"
authors = [
  { name = "Shamit", email = "your.email@example.com" },
]
license = { file = "LICENSE" }
classifiers = [
    "Programming Language :: Python :: 3",
    "License :: OSI Approved :: GNU General Public License v3 (GPLv3)",
    "Operating System :: OS Independent",
]
dependencies = [
    "xlwings~=0.33.6",
    "pandas~=2.2.3",
    "langchain~=0.3.15",
    "langchain_community",
    "langchain_openai",
    "python-dotenv~=1.0.1",
    "langchain-core~=0.3.31",
    "langchain-openai~=0.3.2",
    "Pillow",
    "chainlit~=2.0.601",
    "pydantic~=2.10.6",
    "tabulate",
]
requires-python = ">=3.10"

[project.urls]
"Homepage" = "https://github.com/shamitv/ExcelTamer"
"Bug Tracker" = "https://github.com/shamitv/ExcelTamer/issues"

[tool.setuptools.packages.find]
where = ["."]
include = ["ExcelTamer*"]
exclude = ["test*"]
```

### 3.3. Create `MANIFEST.in` (Optional but recommended)
To ensure non-code files (like `README.md`, `LICENSE`) are included.
```
include LICENSE
include README.md
include requirements.txt
```

### 3.4. Build the Package
Use `build` to generate distribution archives.
```bash
pip install build
python -m build
```
This will create `dist/` containing `.tar.gz` and `.whl` files.

### 3.5. Publish to PyPI
Use `twine` to upload the package.
```bash
pip install twine
twine upload dist/*
```
(Note: Will need PyPI credentials)

### 3.6. Documentation
- [ ] Write User Guide (explaining how to use `ExcelTamerTools` and `ExcelAutomation`)

## 4. Verification Plan

1.  **Local Installation Test**:
    ```bash
    pip install .
    # Test import in a separate script/shell
    python -c "from ExcelTamer import ExcelAutomation"
    ```

2.  **Test PyPI Upload (TestPyPI)**:
    Upload to TestPyPI first to verify metadata.
    ```bash
    twine upload --repository testpypi dist/*
    ```

3.  **Final Release**:
    Upload to real PyPI.
