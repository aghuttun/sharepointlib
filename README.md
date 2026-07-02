# sharepointlib

Microsoft SharePoint client package for Python powered by Rust.

- [Installation](#installation)
- [Usage](#usage)
- [Version](#version)
- [Licence](#licence)

## Installation

```bash
pip install sharepointlib
```

## Usage

```python
import sharepointlib

sharepoint = sharepointlib.SharePoint(
    client_id="<client-id>",
    tenant_id="<tenant-id>",
    client_secret="<client-secret>",
    sp_domain="companygroup.sharepoint.com",
)
```

```python
# Site metadata
response = sharepoint.get_site_info(name="My Site")
if response.status_code == 200:
    print(response.content)
```

```python
# Drives
response = sharepoint.list_drives(site_id="companygroup.sharepoint.com,abc,def")
if response.status_code == 200:
    print(response.content)
```

```python
# List a folder
response = sharepoint.list_dir(
    drive_id="b!123",
    path="Sellout/Support",
    alias=r"_\d{8}",
)
if response.status_code == 200:
    print(response.content)
```

```python
# Download one file
response = sharepoint.download_file(
    drive_id="b!123",
    remote_path="Sellout/Support/report.xlsx",
    local_path=r"C:\\Temp\\report.xlsx",
)
print(response.status_code)
```

```python
# Download to memory (bytes)
response = sharepoint.download_file_to_memory(
    drive_id="b!123",
    remote_path="Sellout/Support/report.xlsx",
)
if response.status_code == 200:
    payload_bytes = response.content
    print(len(payload_bytes))
```

```python
# Copy across drives via streaming
response = sharepoint.copy_file_stream(
    source_drive_id="b!SRC",
    source_path="Sellout/2024/Sellout_November.xlsx",
    target_drive_id="b!DST",
    target_path="Archive/2024",
    new_name="Sellout_November_2024_Backup.xlsx",
    timeout=7200,
)
print(response.status_code)
```

```python
# SharePoint lists
response = sharepoint.list_lists(site_id="companygroup.sharepoint.com,abc,def")
if response.status_code == 200:
    print(response.content)
```

For technical and contributor documentation, see [DEV.md](DEV.md).

## Version

Recommended way to read installed package version:

```python
from importlib.metadata import version

print(version("sharepointlib"))
```

Convenience attribute:

```python
import sharepointlib

print(sharepointlib.__version__)
```

## Licence

BSD-3-Clause Licence (see [LICENSE](LICENSE))
