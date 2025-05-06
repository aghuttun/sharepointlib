[![PyPI version](https://badge.fury.io/py/aghuttun.svg)](https://badge.fury.io/py/aghuttun)
[![PyPI - Python Version](https://img.shields.io/pypi/pyversions/aghuttun)](https://pypi.org/project/aghuttun/)
![Linux](https://img.shields.io/badge/os-Linux-blue.svg)
![macOS Silicon](https://img.shields.io/badge/os-macOS_Silicon-lightgrey.svg)

# sharepointlib

This library provides access to some SharePoint functionalities.

* [Description](#package-description)
* [Usage](#usage)
* [Installation](#installation)
* [Development/Contributing](#developmentcontributing)
* [License](#license)

## Package Description

Microsoft SharePoint interaction (Under development)!

## Usage

* [sharepointlib](#sharepointlib)

from a script:

```python
import sharepointlib
import logging

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(name)s - %(levelname)s - %(message)s")

client_id = "123..."
tenant_id = "123..."
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
result = sharepoint.get_site_info(name=sp_site_name)
if result.status_code == 200:
    print(result.content)
```

```python
# Gets the hostname and site details for a specified site ID
result = sharepoint.get_hostname_info(site_id=sp_site_id)
if result.status_code == 200:
    print(result.content)
```

```python
# Gets the hostname and site details for a specified site ID
result = sharepoint.get_hostname_info(site_id=sp_site_id)
if result.status_code == 200:
    print(result.content)
```

Drives:

```python
# Gets a list of the Drive IDs for a given site ID
result = sharepoint.list_drives(site_id=sp_site_id)
if result.status_code == 200:
    print(result.content)
```

```python
# Gets the folder ID for a specified folder within a drive ID
result = sharepoint.get_dir_info(drive_id=sp_site_drive_id,
                                 path="Sellout/Support")
if result.status_code == 200:
    print(result.content)
```

```python
# List content (files and folders) of a specific folder
result = sharepoint.list_dir(drive_id=sp_site_drive_id, 
                             path="Sellout/Support")
if result.status_code == 200:
    print(result.content)
```

```python
# Creates a new folder in a specified drive ID
result = sharepoint.create_dir(drive_id=sp_site_drive_id, 
                               path="Sellout/Support",
                               name="Archive")
if result.status_code in (200, 201):
    print(result.content)

result = sharepoint.create_dir(drive_id=sp_site_drive_id, 
                               path="Sellout/Support",
                               name="Test")
if result.status_code in (200, 201):
    print(result.content)
```

```python
# Deletes a folder from a specified drive ID
result = sharepoint.delete_dir(drive_id=sp_site_drive_id, 
                               path="Sellout/Support/Test")
if result.status_code == 204:
    print("Folder deleted successfully")
```

```python
# Retrieves information about a specific file in a drive ID
result = sharepoint.get_file_info(drive_id=sp_site_drive_id, 
                                  filename="Sellout/Support/Sellout.xlsx")
if result.status_code == 200:
    print(result.content)
```

```python
# Copy a file from one folder to another within the same drive ID
result = sharepoint.copy_file(drive_id=sp_site_drive_id, 
                              filename="Sellout/Support/Sellout.xlsx",
                              target_path="Sellout/Support/Archive")
if result.status_code == 200:
    print(result.content)
```

```python
# Moves a file from one folder to another within the same drive ID
result = sharepoint.move_file(drive_id=sp_site_drive_id, 
                              filename="Sellout/Support/Sellout1.xlsx", 
                              target_path="Sellout")
if result.status_code == 200:
    print(result.content)
```

```python
# Deletes a file from a specified drive ID
result = sharepoint.delete_file(drive_id=sp_site_drive_id, 
                                filename="Sellout/Support/Sellout.xlsx")
if result.status_code == 204:
    print("File deleted successfully")
```

```python
# Downloads a file from a specified remote path in a drive ID to a local path
result = sharepoint.download_file(drive_id=sp_site_drive_id, 
                                  remote_path=r"Sellout/Sellout.xlsx", 
                                  local_path=r"C:\Users\admin\Downloads\Sellout.xlsx")
if result.status_code == 200:
    print("File downloaded successfully")
```

```python
# Uploads a file to a specified remote path in a SharePoint drive ID
result = sharepoint.upload_file(drive_id=sp_site_drive_id, 
                                local_path=r"C:\Users\admin\Downloads\Sellout.xlsx",
                                remote_path=r"Sellout/Support/Sellout.xlsx")
if result.status_code == 201:
    print(result.content)
```

## Installation

* [sharepointlib](#sharepointlib)

Install python and pip if you have not already.

Then run:

```bash
pip3 install pip --upgrade
pip3 install wheel
```

For production:

```bash
pip3 install sharepointlib
```

This will install the package and all of it's python dependencies.

If you want to install the project for development:

```bash
git clone https://github.com/aghuttun/sharepointlib.git
cd sharepointlib
pip3 install -e .[dev]
```

To test the development package: [Testing](#testing)

## Development/Contributing

* [sharepointlib](#sharepointlib)

1. Fork it!
2. Create your feature branch: `git checkout -b my-new-feature`
3. Test it
5. Commit your changes: `git commit -am 'Add some feature'`
6. Push to the branch: `git push origin my-new-feature`
7. Ensure github actions are passing tests
8. Email me at portela.paulo@gmail.com if it's been a while and I haven't seen it

## License

* [sharepointlib](#sharepointlib)

BSD License (see license file)
