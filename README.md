# sharepointlib

- [Description](#package-description)
- [Usage](#usage)
- [Installation](#installation)
- [Docstring](#docstring)
- [License](#license)

## Package Description

Microsoft SharePoint client Python package that uses the [requests](https://pypi.org/project/requests/) library.

> [!IMPORTANT]  
> This packages uses pydantic v1!

## Usage

From a script:

```python
import sharepointlib
import logging
import pandas as pd

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

client_id = "123..."
tenant_id = "456..."
client_secret = "xxxx"
sp_domain = "companygroup.sharepoint.com"

sp_site_name = "My Site"
sp_site_id = "companygroup.sharepoint.com,1233124124"
sp_site_drive_id = "b!1234567890"

# Initialize SharePoint client
sharepoint = sharepointlib.SharePoint(client_id=client_id,
                                      tenant_id=tenant_id,
                                      client_secret=client_secret,
                                      sp_domain=sp_domain)
```

```python
# Gets the site ID for a given site name
response = sharepoint.get_site_info(name=sp_site_name)
if response.status_code == 200:
    print(response.content["id"])
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Gets the hostname and site details for a specified site ID
response = sharepoint.get_hostname_info(site_id=sp_site_id)
if response.status_code == 200:
    df = pd.DataFrame([response.content])
    print(df)
```

Drives:

```python
# Gets a list of the Drive IDs for a given site ID
response = sharepoint.list_drives(site_id=sp_site_id)
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# Gets the folder ID for a specified folder within a drive ID
response = sharepoint.get_dir_info(drive_id=sp_site_drive_id, path="Sellout/Support")
if response.status_code == 200:
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# List content (files and folders) of a specific folder
response = sharepoint.list_dir(drive_id=sp_site_drive_id, path="Sellout/Support")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# List content (files and folders) of a specific folder changing the filename
#
# The "alias" column allows simplifying the value of the "name" column using regular expressions.
# Example: column name: My_Filename_20251107.xlsx –> alias parameter: r"_\d{8}" –> alias column: My_Filename.xlsx.
#
# If the alias is set to None, then the regular expression r"$^" is used; which means the alias column will be identical
# to the name column - in other words, the value of the "alias" column will be the same as that of the "name" column.
response = sharepoint.list_dir(drive_id=sp_site_drive_id, path="Sellout/Support", alias=r"_\d{8}")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    print(df)
```

```python
# Creates a new folder in a specified drive ID
response = sharepoint.create_dir(drive_id=sp_site_drive_id, path="Sellout/Support", name="Archive")
if response.status_code in (200, 201):
    df = pd.DataFrame([response.content])
    print(df)

response = sharepoint.create_dir(drive_id=sp_site_drive_id, path="Sellout/Support", name="Test")
if response.status_code in (200, 201):
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Deletes a folder from a specified drive ID
response = sharepoint.delete_dir(drive_id=sp_site_drive_id, path="Sellout/Support/Test")
if response.status_code in (200, 204):
    print("Folder deleted successfully")
```

```python
# Renames a folder in a specified drive ID
response = sharepoint.rename_folder(drive_id=sp_site_drive_id, path="Sellout/Support", new_name="Old")
if response.status_code == 200:
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Retrieves information about a specific file in a drive ID
response = sharepoint.get_file_info(drive_id=sp_site_drive_id, filename="Sellout/Support/Sellout.xlsx")
if response.status_code in (200, 202):
    print(response.content["id"])
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Perform check-out of a file
response = sharepoint.check_out_file(drive_id=sp_site_drive_id, filename="Sellout/Support/Sellout.xlsx")
if response.status_code in (200, 204):
    print("Check-out completed successfully")
```

```python
# Perform check-in of a file
response = sharepoint.check_in_file(
    drive_id=sp_site_drive_id,
    filename="Sellout/Support/Sellout.xlsx",
    comment="Reviewed and approved"
)
if response.status_code in (200, 204):
    print("Check-in completed successfully")
```

```python
# Copy a file from one folder to another within the same drive ID
response = sharepoint.copy_file(drive_id=sp_site_drive_id,
                                filename="Sellout/Support/Archive/My Book.xlsx",
                                target_path="Sellout/Support/",
                                new_name="My Book Copy.xlsx")
if response.status_code in (200, 202):
    print("File copied successfully")
```

```python
# Moves a file from one folder to another within the same drive ID
response = sharepoint.move_file(drive_id=sp_site_drive_id,
                                filename="Sellout/Support/Book1.xlsx",
                                target_path="Sellout/Support/Archive/",
                                new_name="My Book.xlsx")
if response.status_code == 200:
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Deletes a file from a specified drive ID
response = sharepoint.delete_file(drive_id=sp_site_drive_id, filename="Sellout/Sellout.xlsx")
if response.status_code in (200, 204):
    print("File deleted successfully")
```

```python
# Renames a file in a specified drive ID.
response = sharepoint.rename_file(drive_id=sp_site_drive_id,
                                  filename="Sellout/Support/Archive/Sellout.xlsx",
                                  new_name="Sellout_New_Name.xlsx")
if response.status_code == 200:
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Downloads a file from a specified remote path in a drive ID to a local path
# Examples for local_path (databricks):
#   local_path=r"/Workspace/Users/admin@admin.com/Sellout.xlsm"
#   local_path=r"/Volumes/lakehouse/sadp/Sellout.xlsm"
response = sharepoint.download_file(drive_id=sp_site_drive_id,
                                    remote_path=r"Sellout/Support/Sellout.xlsx",
                                    local_path=r"C:\Users\admin\Downloads\Sellout.xlsx")
if response.status_code == 200:
    print("File downloaded successfully")
```

```python
# Downloads all files from a specified remote path in a drive ID to a local path
# Examples for local_path (databricks):
#   local_path=r"/Workspace/Users/admin@admin.com/"
#   local_path=r"/Volumes/lakehouse/sadp/"
response = sharepoint.download_all_files(drive_id=sp_site_drive_id,
                                         remote_path=r"Sellout/Support",
                                         local_path=r"C:\Users\admin\Downloads")
if response.status_code == 200:
    df = pd.DataFrame(response.content)
    display(df)  # print(df)
```

```python
# Downloads an Excel file from SharePoint directly into memory and loads it into a Pandas DataFrame
from io import BytesIO

response = sharepoint.download_file_to_memory(drive_id=sp_site_drive_id,
                                              remote_path="Sellout/Support/Sellout.xlsx")

if response.status_code == 200 and response.content is not None:
    excel_data = BytesIO(response.content)
    df = pd.read_excel(excel_data)
    print(df)
```

```python
# Uploads a file to a specified remote path in a SharePoint drive ID
response = sharepoint.upload_file(drive_id=sp_site_drive_id,
                                  local_path=r"C:\Users\admin\Downloads\Sellout.xlsx",
                                  remote_path=r"Sellout/Support/Archive/Sellout.xlsx")
if response.status_code in (200, 201):
    df = pd.DataFrame([response.content])
    print(df)
```

```python
# Close
del(sharepoint)
```

## Installation

Install python and pip if you have not already.

Then run:

```bash
pip install pip --upgrade
```

For production:

```bash
pip install sharepointlib
```

This will install the package and all of it's python dependencies.

If you want to install the project for development:

```bash
git clone https://github.com/aghuttun/sharepointlib.git
cd sharepointlib
pip install -e ".[dev]"
```

## Docstring

The script's docstrings follow the numpydoc style.

## License

BSD License (see license file)

[top](#sharepointlib)
