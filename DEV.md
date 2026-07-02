# sharepointlib — Developer Guide

- [Overview](#overview)
- [Requirements](#requirements)
- [Version Access](#version-access)
- [Project Structure](#project-structure)
- [Building](#building)
- [API Reference](#api-reference)
- [Tests](#tests)

## Overview

sharepointlib is implemented in Rust using [PyO3](https://pyo3.rs/) for Python bindings and
[reqwest](https://crates.io/crates/reqwest) for Microsoft Graph HTTP operations.
[Maturin](https://www.maturin.rs/) is used for building and packaging.

## Requirements

- Python 3.10 or higher
- Rust stable toolchain (edition 2024 crate)
- Maturin 1.7 or higher

```bash
pip install maturin
```

Install local development dependencies:

```bash
python -m pip install -e .[dev]
```

## Version Access

Recommended:

```python
from importlib.metadata import version

print(version("sharepointlib"))
```

Convenience alias:

```python
import sharepointlib

print(sharepointlib.__version__)
```

## Project Structure

```
sharepointlib/
├── src/
│   ├── lib.rs              # Rust SharePoint core implementation
│   ├── models.rs           # Field sets used in Graph $select queries
│   └── python_bindings.rs  # PyO3 module and Python-facing wrappers
├── tests/
│   └── test_smoke.py       # Public API smoke tests
├── old_python_version/
│   └── sharepointlib/      # Archived pure-Python implementation
├── Cargo.toml              # Rust crate configuration
├── pyproject.toml          # Python packaging configuration (maturin backend)
├── README.md               # User-facing documentation
└── DEV.md                  # This file
```

## Building

Development build:

```bash
maturin develop --features python
```

Release build:

```bash
maturin develop --release --features python
```

Build wheel:

```bash
maturin build --release --features python --out dist
```

## API Reference

### `SharePoint(client_id, tenant_id, client_secret, sp_domain, custom_logger=None)`

Initialise the client and authenticate.

### Methods

- `renew_token()`
- `get_site_info(name, save_as=None)`
- `get_hostname_info(site_id, save_as=None)`
- `list_drives(site_id, save_as=None)`
- `get_dir_info(drive_id, path=None, save_as=None)`
- `list_dir(drive_id, path=None, alias=None, save_as=None)`
- `create_dir(drive_id, path, name, save_as=None)`
- `delete_dir(drive_id, path)`
- `rename_folder(drive_id, path, new_name, save_as=None)`
- `get_file_info(drive_id, filename, save_as=None)`
- `check_out_file(drive_id, filename)`
- `check_in_file(drive_id, filename, comment=None)`
- `copy_file(drive_id, filename, target_path, new_name=None)`
- `get_download_url(drive_id, filename)`
- `copy_file_stream(source_drive_id, source_path, target_drive_id, target_path, new_name=None, chunk_size=20971520, timeout=3600, save_as=None)`
- `move_file(drive_id, filename, target_path, new_name=None, save_as=None)`
- `delete_file(drive_id, filename)`
- `rename_file(drive_id, filename, new_name, save_as=None)`
- `download_file(drive_id, remote_path, local_path)`
- `download_file_to_memory(drive_id, remote_path)`
- `download_all_files(drive_id, remote_path, local_path)`
- `upload_file(drive_id, local_path, remote_path, save_as=None)`
- `list_lists(site_id, save_as=None)`
- `list_list_columns(site_id, list_id, save_as=None)`
- `list_list_items(site_id, list_id, fields, save_as=None)`
- `delete_list_item(site_id, list_id, item_id)`
- `add_list_item(site_id, list_id, item, save_as=None)`

### Response Shape

All operations returning `Response` expose:

- `status_code`: HTTP status code.
- `content`: decoded payload (`dict`, `list`, `bytes`, or `None`).

### Errors

Python-facing methods raise `RuntimeError` when Rust operations fail.

## Tests

Smoke tests:

```bash
python -m pytest tests/test_smoke.py -q
```
