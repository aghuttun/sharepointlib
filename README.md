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
# Force token refresh (for long-running sessions)
sharepoint.renew_token()
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

```python
# Resolve hostname details from a site ID
response = sharepoint.get_hostname_info(site_id="companygroup.sharepoint.com,abc,def")
if response.status_code == 200:
    print(response.content)
```

```python
# Get folder metadata
response = sharepoint.get_dir_info(
    drive_id="b!123",
    path="Sellout/Support",
)
if response.status_code == 200:
    print(response.content)
```

```python
# Create / rename / delete a folder
created = sharepoint.create_dir(
    drive_id="b!123",
    path="Sellout",
    name="Support_New",
)

renamed = sharepoint.rename_folder(
    drive_id="b!123",
    path="Sellout/Support_New",
    new_name="Support_Renamed",
)

deleted = sharepoint.delete_dir(
    drive_id="b!123",
    path="Sellout/Support_Renamed",
)

print(created.status_code, renamed.status_code, deleted.status_code)
```

```python
# File metadata and download URL
info = sharepoint.get_file_info(
    drive_id="b!123",
    filename="Sellout/Support/report.xlsx",
)
url = sharepoint.get_download_url(
    drive_id="b!123",
    filename="Sellout/Support/report.xlsx",
)

print(info.status_code)
print(url)
```

```python
# Upload a local file
response = sharepoint.upload_file(
    drive_id="b!123",
    local_path=r"C:\\Temp\\report.xlsx",
    remote_path="Sellout/Support",
)
print(response.status_code)
```

```python
# Move / rename / copy / delete a file
moved = sharepoint.move_file(
    drive_id="b!123",
    filename="Sellout/Support/report.xlsx",
    target_path="Archive/2024",
    new_name="report_2024.xlsx",
)

renamed = sharepoint.rename_file(
    drive_id="b!123",
    filename="Archive/2024/report_2024.xlsx",
    new_name="report_final.xlsx",
)

copied = sharepoint.copy_file(
    drive_id="b!123",
    filename="Archive/2024/report_final.xlsx",
    target_path="Archive/Backup",
)

deleted = sharepoint.delete_file(
    drive_id="b!123",
    filename="Archive/Backup/report_final.xlsx",
)

print(moved.status_code, renamed.status_code, copied.status_code, deleted.status_code)
```

```python
# Check out and check in a file
checked_out = sharepoint.check_out_file(
    drive_id="b!123",
    filename="Sellout/Support/report.xlsx",
)

checked_in = sharepoint.check_in_file(
    drive_id="b!123",
    filename="Sellout/Support/report.xlsx",
    comment="Updated after reconciliation",
)

print(checked_out.status_code, checked_in.status_code)
```

```python
# Download all files from a folder to a local directory
response = sharepoint.download_all_files(
    drive_id="b!123",
    remote_path="Sellout/Support",
    local_path=r"C:\\Temp\\support",
)
print(response.status_code)
```

```python
# List list columns
response = sharepoint.list_list_columns(
    site_id="companygroup.sharepoint.com,abc,def",
    list_id="1234-5678-90ab-cdef",
)
if response.status_code == 200:
    print(response.content)
```

```python
# List list items with selected fields
response = sharepoint.list_list_items(
    site_id="companygroup.sharepoint.com,abc,def",
    list_id="1234-5678-90ab-cdef",
    fields="Title,Status,Owner",
)
if response.status_code == 200:
    print(response.content)
```

```python
# Add and delete a list item
created = sharepoint.add_list_item(
    site_id="companygroup.sharepoint.com,abc,def",
    list_id="1234-5678-90ab-cdef",
    item={"Title": "New task", "Status": "Open"},
)

deleted = sharepoint.delete_list_item(
    site_id="companygroup.sharepoint.com,abc,def",
    list_id="1234-5678-90ab-cdef",
    item_id="42",
)

print(created.status_code, deleted.status_code)
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
