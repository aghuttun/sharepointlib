"""
Generic SharePoint Client Library.

This module provides a generic Python client for interacting with SharePoint sites, drives, folders, files, and lists
via the Microsoft Graph API.
It supports authentication using OAuth2 client credentials flow and offers methods for managing files, folders, and
SharePoint lists.
The library is designed to be reusable and should not be modified for project-specific requirements.

Features:
- OAuth2 authentication with Azure AD
- List, create, delete, rename, move, copy, download, and upload files and folders
- Access and manipulate SharePoint lists and list items
- Export API responses to JSON files
- Pydantic-based response validation

Intended for use as a utility library in broader data engineering and automation workflows.
"""

# import base64
import dataclasses
import json
import logging
from typing import Any, Type
import re
from urllib.parse import quote
import requests

# TypeAdapter v2 vs parse_obj_as v1
from pydantic import BaseModel, parse_obj_as  # pylint: disable=no-name-in-module
from .models import (
    GetSiteInfo,
    GetHostNameInfo,
    ListDrives,
    GetDirInfo,
    ListDir,
    CreateDir,
    RenameFolder,
    GetFileInfo,
    CheckOutFile,
    CheckInFile,
    MoveFile,
    RenameFile,
    UploadFile,
    ListLists,
    ListListColumns,
    AddListItem,
)

# Creates a logger for this module
logger = logging.getLogger(__name__)


class SharePoint(object):
    """
    Interact with SharePoint sites, drives, folders, files, and lists via the Microsoft Graph API.

    Authenticate using OAuth2 client credentials flow. Manage files and folders, including listing, creating, deleting,
    renaming, moving, copying, downloading, and uploading. Access and manipulate SharePoint lists and list items.

    Parameters
    ----------
    client_id : str
        Specify the Azure client ID for authentication.
    tenant_id : str
        Specify the Azure tenant ID associated with the client.
    client_secret : str
        Specify the secret key for the Azure client.
    sp_domain : str
        Specify the SharePoint domain (e.g., "companygroup.sharepoint.com").
    custom_logger : logging.Logger, optional
        Provide a custom logger instance. If None, use the default logger.

    Attributes
    ----------
    _logger : logging.Logger
        Logger instance for logging informational and error messages.
    _session : requests.Session
        Session object for making HTTP requests.
    _configuration : SharePoint.Configuration
        Configuration dataclass containing API and authentication details.

    Methods
    -------
    renew_token()
        Force re-authentication to obtain a new access token.
    get_site_info(name, save_as=None)
        Retrieve the site ID for a given site name.
    get_hostname_info(site_id, save_as=None)
        Retrieve the hostname and site details for a specified site ID.
    list_drives(site_id, save_as=None)
        List the Drive IDs for a given site ID.
    get_dir_info(drive_id, path=None, save_as=None)
        Retrieve the folder ID for a specified folder within a drive ID.
    list_dir(drive_id, path=None, save_as=None)
        List content (files and folders) of a specific folder.
    create_dir(drive_id, path, name, save_as=None)
        Create a new folder in a specified drive ID.
    delete_dir(drive_id, path)
        Delete a folder from a specified drive ID.
    rename_folder(drive_id, path, new_name, save_as=None)
        Rename a folder in a specified drive ID.
    get_file_info(drive_id, filename, save_as=None)
        Retrieve information about a specific file in a drive ID.
    copy_file(drive_id, filename, target_path, new_name=None)
        Copy a file from one folder to another within the same drive ID.
    move_file(drive_id, filename, target_path, new_name=None, save_as=None)
        Move a file from one folder to another within the same drive ID.
    delete_file(drive_id, filename)
        Delete a file from a specified drive ID.
    rename_file(drive_id, filename, new_name, save_as=None)
        Rename a file in a specified drive ID.
    download_file(drive_id, remote_path, local_path)
        Download a file from a specified remote path in a drive ID to a local path.
    download_file_to_memory(drive_id, remote_path)
        Download a file from a specified remote path in a drive ID to memory.
    download_all_files(drive_id, remote_path, local_path)
        Download all files from a specified folder to a local folder.
    upload_file(drive_id, local_path, remote_path, save_as=None)
        Upload a file to a specified remote path in a SharePoint drive ID.
    list_lists(site_id, save_as=None)
        Retrieve a list of SharePoint lists for a specified site.
    list_list_columns(site_id, list_id, save_as=None)
        Retrieve the columns from a specified list in SharePoint.
    list_list_items(site_id, list_id, fields, save_as=None)
        Retrieve the items from a specified list in SharePoint.
    delete_list_item(site_id, list_id, item_id)
        Delete a specified item from a list in SharePoint.
    add_list_item(site_id, list_id, item, save_as=None)
        Add a new item to a specified list in SharePoint.
    """

    @dataclasses.dataclass
    class Configuration:
        """
        Define configuration parameters for the SharePoint client.

        Set API domain, API version, SharePoint domain, Azure client credentials, and OAuth2 token.

        Parameters
        ----------
        api_domain : str or None, optional
            Specify the Microsoft Graph API domain.
        api_version : str or None, optional
            Specify the Microsoft Graph API version.
        sp_domain : str or None, optional
            Specify the SharePoint domain.
        client_id : str or None, optional
            Specify the Azure client ID for authentication.
        tenant_id : str or None, optional
            Specify the Azure tenant ID associated with the client.
        client_secret : str or None, optional
            Specify the secret key for the Azure client.
        token : str or None, optional
            Specify the OAuth2 access token for API requests.
        """

        api_domain: str | None = None
        api_version: str | None = None
        sp_domain: str | None = None
        client_id: str | None = None
        tenant_id: str | None = None
        client_secret: str | None = None
        token: str | None = None

    @dataclasses.dataclass
    class Response:
        """
        Represent the response from SharePoint client methods.

        Parameters
        ----------
        status_code : int
            Specify the HTTP status code returned by the SharePoint API request.
        content : Any, optional
            Provide the content returned by the SharePoint API request. Can be a dictionary, list, bytes, or None.

        Examples
        --------
        >>> resp = Response(status_code=200, content={"id": "123", "name": "example"})
        >>> print(resp.status_code)
        200
        >>> print(resp.content)
        {'id': '123', 'name': 'example'}
        """

        status_code: int
        content: Any = None

    def __init__(
        self,
        client_id: str,
        tenant_id: str,
        client_secret: str,
        sp_domain: str,
        custom_logger: logging.Logger | None = None,
    ) -> None:
        """
        Initialize the SharePoint client.

        Set up the configuration, initialize the HTTP session, and authenticate the client using OAuth2 client
        credentials flow.

        Parameters
        ----------
        client_id : str
            Specify the Azure client ID for authentication.
        tenant_id : str
            Specify the Azure tenant ID associated with the client.
        client_secret : str
            Specify the secret key for the Azure client.
        sp_domain : str
            Specify the SharePoint domain (e.g., "companygroup.sharepoint.com").
        custom_logger : logging.Logger, optional
            Provide a custom logger instance. If None, use the default logger.

        Notes
        -----
        Use the provided logger or create a default one. Store credentials and configuration. Authenticate immediately.
        """
        # Init logging
        # Use provided logger or create a default one
        self._logger = custom_logger or logging.getLogger(name=__name__)

        # Init variables
        self._session: requests.Session = requests.Session()
        api_domain = "graph.microsoft.com"
        api_version = "v1.0"

        # Credentials/Configuration
        self._configuration = self.Configuration(
            api_domain=api_domain,
            api_version=api_version,
            sp_domain=sp_domain,
            client_id=client_id,
            tenant_id=tenant_id,
            client_secret=client_secret,
            token=None,
        )

        # Authenticate
        self._authenticate()

    def __del__(self) -> None:
        """
        Finalize the SharePoint client instance and release resources.

        Close the internal HTTP session and log an informational message indicating cleanup.

        Parameters
        ----------
        self : SharePoint
            The SharePoint client instance.

        Returns
        -------
        None

        Notes
        -----
        This method is called when the instance is about to be destroyed. Ensure the HTTP session is closed and log
        cleanup.
        """
        self._logger.info(msg="Cleaning the house at the exit")
        self._session.close()

    def _authenticate(self) -> None:
        """
        Authenticate the SharePoint client using OAuth2 client credentials flow.

        Send a POST request to the Azure AD v2.0 token endpoint using the tenant ID, client ID, and client secret
        stored in the configuration. On success, extract the access token from the response and assign it to the
        configuration for subsequent API requests.

        Parameters
        ----------
        self : SharePoint
            Instance of the SharePoint client.

        Returns
        -------
        None

        Raises
        ------
        requests.exceptions.RequestException
            If the HTTP request fails due to network issues, DNS resolution, or SSL errors.
        RuntimeError
            If the token endpoint returns a non-200 HTTP status code.
        ValueError
            If the response body does not contain a valid JSON access_token.
        """
        self._logger.info(msg="Authenticating")

        # Request headers
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        # Authorization endpoint
        url_auth = f"https://login.microsoftonline.com/{self._configuration.tenant_id}/oauth2/v2.0/token"

        # Request body
        body = {
            "grant_type": "client_credentials",
            "client_id": self._configuration.client_id,
            "client_secret": self._configuration.client_secret,
            "scope": "https://graph.microsoft.com/.default",
        }

        # Send request
        response = self._session.post(url=url_auth, data=body, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Return valid response
        if response.status_code == 200:
            self._configuration.token = json.loads(response.content.decode("utf-8"))["access_token"]

    def renew_token(self) -> None:
        """
        This method forces a re-authentication to obtain a new access token.

        The new token is stored in the token attribute.
        """
        self._authenticate()

    def _export_to_json(self, content: bytes, save_as: str | None) -> None:
        """
        Export response content to a JSON file.

        Save the given bytes content to a file in binary mode if a file path is provided.

        Parameters
        ----------
        content : bytes
            Response content to export.
        save_as : str or None
            File path to save the JSON content. If None, do not save.

        Returns
        -------
        None
            This function does not return any value.

        Notes
        -----
        If `save_as` is specified, write the content to the file in binary mode.
        """
        if save_as is not None:
            self._logger.info(msg="Exporting response to JSON file")
            with open(file=save_as, mode="wb") as file:
                file.write(content)

    def _handle_response(
        self, response: requests.Response, model: Type[BaseModel], rtype: str = "scalar"
    ) -> dict | list[dict]:
        """
        Handle and deserialize the JSON content from an API response.

        Parameters
        ----------
        response : requests.Response
            Response object from the API request.
        model : Type[BaseModel]
            Pydantic BaseModel class for deserialization and validation.
        rtype : str, optional
            Specify "scalar" for a single record or "list" for a list of records. Default is "scalar".

        Returns
        -------
        dict or list of dict
            Deserialized content as a dictionary (for scalar) or a list of dictionaries (for list).

        Examples
        --------
        >>> self._handle_response(response, MyModel, rtype="scalar")
        {'field1': 'value1', 'field2': 'value2'}

        >>> self._handle_response(response, MyModel, rtype="list")
        [{'field1': 'value1'}, {'field1': 'value2'}]
        """
        if rtype.lower() == "scalar":
            # Deserialize json (scalar values)
            content_raw = response.json()
            # Pydantic v1 validation
            validated = model(**content_raw)
            # Convert to dict
            return validated.dict()

        # List of records
        # Deserialize json
        content_raw = response.json()["value"]
        # Pydantic v1 validation
        validated_list = parse_obj_as(list[model], content_raw)
        # return [dict(data) for data in parse_obj_as(list[model], content_raw)]
        # Convert to a list of dicts
        return [item.dict() for item in validated_list]

    def get_site_info(self, name: str, save_as: str | None = None) -> Response:
        """
        Retrieve the site ID for a given site name.

        Send a request to the Microsoft Graph API and return the site ID and related information.
        Optionally, save the response to a JSON file.

        Parameters
        ----------
        name : str
            Specify the name of the site to retrieve the site ID for.
        save_as : str or None, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, including the site ID and other
            relevant information.

        Notes
        -----
        Validate the returned content using the GetSiteInfo Pydantic model.
        """
        self._logger.info(msg="Retrieving the site ID for the specified site name")
        self._logger.info(msg=name)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        sp_domain = self._configuration.sp_domain

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        # get_sites_id: url_query = f"https://graph.microsoft.com/v1.0/sites?search='{filter}'"
        url_query = rf"https://{api_domain}/{api_version}/sites/{sp_domain}:/sites/{name}"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in GetSiteInfo.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=GetSiteInfo, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def get_hostname_info(self, site_id: str, save_as: str | None = None) -> Response:
        """
        Retrieve the hostname and site details for a specified site ID.

        Send a request to the Microsoft Graph API and return the hostname, site name, and other relevant details
        associated with the given site ID. Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            Specify the ID of the site for which to retrieve the hostname and details.
        save_as : str or None, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes the hostname,
            site name, and other relevant information.

        Examples
        --------
        >>> sp = SharePoint(client_id, tenant_id, client_secret, sp_domain)
        >>> resp = sp.get_hostname_info(site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1")
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Retrieving the hostname and site details for the specified site ID")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}"

        # Query parameters
        # Pydantic v1
        alias_list = [
            field.alias for field in GetHostNameInfo.__fields__.values() if field.field_info.alias is not None
        ]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=GetHostNameInfo, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    # DRIVES
    def list_drives(self, site_id: str, save_as: str | None = None) -> Response:
        """
        List Drive IDs for a given site.

        Send a request to the Microsoft Graph API to retrieve Drive IDs associated with the specified site ID.
        Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            Specify the ID of the site for which to list Drive IDs.
        save_as : str or None, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the content, which includes the list of Drive IDs and
            related information.

        Notes
        -----
        Validate the returned content using the ListDrives Pydantic model.
        """
        self._logger.info(msg="Retrieving a list of Drive IDs for the specified site")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/drives"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in ListDrives.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = self._handle_response(response=response, model=ListDrives, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def get_dir_info(self, drive_id: str, path: str | None = None, save_as: str | None = None) -> Response:
        """
        Get the folder ID for a specified folder within a drive.

        Send a request to the Microsoft Graph API and return the folder ID and related information.
        Optionally, save the response to a JSON file.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the folder.
        path : str, optional
            Specify the path of the folder. If None, use the root folder.
        save_as : str, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the content, including the folder ID and other relevant
            information.

        Notes
        -----
        Validate the returned content using the GetDirInfo Pydantic model.
        """
        self._logger.info(msg="Retrieving the folder ID for the specified folder within the drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        path_quote = "///" if path is None else f"/{quote(string=path)}"
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:{path_quote}"
        # print(url_query)

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in GetDirInfo.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=GetDirInfo, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def list_dir(
        self,
        drive_id: str,
        path: str | None = None,
        alias: str | None = None,
        save_as: str | None = None,
    ) -> Response:
        """
        List the contents (files and folders) of a folder in a SharePoint drive.

        Parameters
        ----------
        drive_id : str
            ID of the drive containing the folder.
        path : str, optional
            Path of the folder to list. If None, use the root folder.
        alias : str, optional
            Regex pattern to extract an alias from the item names. If provided, the alias will be added to each item.
        save_as : str, optional
            Path to save the results as a JSON file. If None, do not save.

        Returns
        -------
        Response
            Response object containing the HTTP status code and a list of items (files and folders) in the folder.

        Notes
        -----
        Send a request to the Microsoft Graph API to retrieve the list of children in the specified folder.
        If successful, return the HTTP status code and a list of items. Add the folder path to each item in the result.
        """
        self._logger.info(msg="Listing the contents of the specified folder in the drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        path_quote = "/" if path is None else f"{quote(string=path)}"
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/items/root:/{path_quote}:/children"
        # print(url_query)

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in ListDir.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=ListDir, rtype="list")

            # Add path to each item
            # Also extract alias from name if regex provided
            for item in content:
                item["path"] = path or "/"
                item["alias"] = (
                    re.sub(pattern=alias, repl="", string=item.get("name", ""))
                    if alias
                    else re.sub(pattern=r"$^", repl="", string=item.get("name", ""))
                )

        return self.Response(status_code=response.status_code, content=content)

    def create_dir(self, drive_id: str, path: str, name: str, save_as: str | None = None) -> Response:
        """
        Create a new folder in a specified drive.

        Parameters
        ----------
        drive_id : str
            ID of the drive where the folder will be created.
        path : str
            Path within the drive where the new folder will be created.
        name : str
            Name of the new folder.
        save_as : str or None, optional
            File path to save the response as a JSON file. If None, do not save.

        Returns
        -------
        Response
            Response object containing the HTTP status code and details of the created folder.

        Notes
        -----
        Validate the returned content using the CreateDir Pydantic model.
        """
        self._logger.info(msg="Creating a new folder in the specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        path_quote = quote(string=path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}:/children"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in CreateDir.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        # @microsoft.graph.conflictBehavior: fail, rename, replace
        data = {
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace",
        }

        # Send request
        response = self._session.post(url=url_query, headers=headers, params=params, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 201):
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=CreateDir, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def delete_dir(self, drive_id: str, path: str) -> Response:
        """
        Delete a folder from a specified drive.

        Send a DELETE request to the Microsoft Graph API to remove a folder at the given path within the specified
        drive.

        Parameters
        ----------
        drive_id : str
            ID of the drive containing the folder to delete.
        path : str
            Full path of the folder to delete.

        Returns
        -------
        Response
            Response object containing the HTTP status code and details of the operation.

        Notes
        -----
        Return a successful HTTP status code if the folder is deleted.
        """
        self._logger.info(msg="Deleting a folder from the specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        path_quote = quote(string=path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}"

        # Send request
        response = self._session.delete(url=url_query, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 204):
            self._logger.info(msg="Request successful")

        return self.Response(status_code=response.status_code, content=content)

    def rename_folder(self, drive_id: str, path: str, new_name: str, save_as: str | None = None) -> Response:
        """
        Rename a folder in a specified drive.

        Send a PATCH request to the Microsoft Graph API to rename a folder at the given path within the specified drive.

        Parameters
        ----------
        drive_id : str
            ID of the drive containing the folder to rename.
        path : str
            Full path of the folder to rename.
        new_name : str
            New name for the folder.
        save_as : str, optional
            File path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Response dataclass instance containing the HTTP status code and the content.

        Notes
        -----
        Validate the returned content using the RenameFolder Pydantic model.
        """
        self._logger.info(msg="Renaming a folder in the specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)
        self._logger.info(msg=new_name)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        path_quote = quote(string=path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}"

        # Request body
        data = {"name": new_name}

        alias_list = [field.alias for field in RenameFolder.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.patch(url=url_query, headers=headers, params=params, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=RenameFolder, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def get_file_info(self, drive_id: str, filename: str, save_as: str | None = None) -> Response:
        """
        Retrieve information about a specific file in a drive.

        Send a request to the Microsoft Graph API to obtain details about a file located at the specified path within
        the given drive ID. Optionally, save the response to a JSON file.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the file.
        filename : str
            Specify the full path of the file, including the filename.
        save_as : str or None, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes file details such
            as ID, name, web URL, size, created date, and last modified date.

        Notes
        -----
        Validate the returned content using the GetFileInfo Pydantic model.
        """
        self._logger.info(msg="Retrieving information about a specific file")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        filename_quote = quote(string=filename)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in GetFileInfo.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 202):
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=GetFileInfo, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def check_out_file(self, drive_id: str, filename: str) -> Response:
        """
        Perform a check-out on a SharePoint file.

        Lock the file for exclusive editing by the authenticated user. This operation is only supported in document
        libraries where check-out is required.

        Parameters
        ----------
        drive_id : str
            Identify the drive (document library) that contains the file.
        filename : str
            Specify the full path of the file, including the filename (e.g. "Folder/Subfolder/report.docx").

        Returns
        -------
        Response
            Contain the HTTP status code and, on success, the updated driveItem (including the publication facet
            showing the checkout state).

        Raises
        ------
        RuntimeError
            Propagate any error returned by the underlying get_file_info call.

        See Also
        --------
        check_in_file : Release the lock and publish a new version.
        """
        self._logger.info(msg="Initiating check-out process for the file")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)

        # Obtain the file ID using get_file_info
        file_info = self.get_file_info(drive_id=drive_id, filename=filename)

        if file_info.status_code != 200:
            return self.Response(status_code=file_info.status_code, content=None)

        file_id = file_info.content["id"]

        # Build request
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        # Endpoint
        url = f"https://{api_domain}/{api_version}/drives/{drive_id}/items/{file_id}/checkout"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in CheckOutFile.__fields__.values() if field.field_info.alias]
        params = {"$select": ",".join(alias_list)} if alias_list else None

        # Send request
        response = self._session.post(url=url, headers=headers, params=params, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 204):
            self._logger.info(msg="Check-out completed successfully")

        return self.Response(status_code=response.status_code, content=content)

    def check_in_file(self, drive_id: str, filename: str, comment: str | None = None) -> Response:
        """
        Perform a check-in on a SharePoint file.

        Release the exclusive lock and, if versioning is enabled, create a new version. The file must be in a
        checked-out state.

        Parameters
        ----------
        drive_id : str
            Identify the drive (document library) that contains the file.
        filename : str
            Specify the full path of the file, including the filename.
        comment : str | None, optional
            Provide a version comment. Required when the library uses major versioning; ignored otherwise.

        Returns
        -------
        Response
            Contain the HTTP status code and, on success, the updated driveItem (publication.level will be "published").

        Raises
        ------
        RuntimeError
            Propagate any error returned by the underlying get_file_info call.
        """
        self._logger.info(msg="Initiating check-in process for the file")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)

        # Obtain the file ID using get_file_info
        file_info = self.get_file_info(drive_id=drive_id, filename=filename)

        if file_info.status_code != 200:
            return self.Response(status_code=file_info.status_code, content=None)

        file_id = file_info.content["id"]

        # Build request
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        # Endpoint
        url = f"https://{api_domain}/{api_version}/drives/{drive_id}/items/{file_id}/checkin"

        data: dict = {}
        if comment:
            data["comment"] = comment

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in CheckInFile.__fields__.values() if field.field_info.alias]
        params = {"$select": ",".join(alias_list)} if alias_list else None

        # Send request
        response = self._session.post(url=url, headers=headers, params=params, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 204):
            self._logger.info(msg="Check-in completed successfully")

        return self.Response(status_code=response.status_code, content=content)

    def copy_file(self, drive_id: str, filename: str, target_path: str, new_name: str | None = None) -> Response:
        """
        Copy a file from one folder to another within the same drive.

        Send a request to the Microsoft Graph API to copy a file from the specified source path to the destination path
        within the given drive ID. Operate only within the same drive. Return the HTTP status code and details of the
        copied file.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the file to copy.
        filename : str
            Specify the full path of the file to copy, including the filename.
        target_path : str
            Specify the path of the destination folder where to copy the file.
        new_name : str, optional
            Specify a new name for the copied file. If not provided, retain the original name.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes details of the
            copied file.

        Notes
        -----
        The copy operation is asynchronous. The response may indicate that the operation is in progress.
        """
        self._logger.info(msg="Copying a file from one folder to another within the same drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)
        self._logger.info(msg=target_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        filename_quote = quote(string=filename)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}:/copy"

        # Request body
        data = {
            "parentReference": {
                "driveId": drive_id,
                "driveType": "documentLibrary",
                "path": f"/drives/{drive_id}/root:/{target_path}",
            }
        }
        # Add to the request body if new_name is provided
        if new_name is not None:
            data["name"] = new_name  # type: ignore

        # Send request
        response = self._session.post(url=url_query, headers=headers, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 202):
            self._logger.info(msg="Request successful")

        return self.Response(status_code=response.status_code, content=content)

    def move_file(
        self, drive_id: str, filename: str, target_path: str, new_name: str | None = None, save_as: str | None = None
    ) -> Response:
        """
        Move a file from one folder to another within the same drive.

        Move the specified file from the source path to the destination path within the given drive ID using the
        Microsoft Graph API. Optionally, rename the file during the move and save the response to a JSON file.

        Parameters
        ----------
        drive_id : str
            ID of the drive containing the file to move.
        filename : str
            Full path of the file to move, including the filename.
        target_path : str
            Path of the destination folder where to move the file.
        new_name : str, optional
            New name for the file after moving. If not provided, retain the original name.
        save_as : str, optional
            File path to save the response in JSON format. If not provided, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the content, including details of the moved file.

        Raises
        ------
        RuntimeError
            If the source file or destination folder cannot be found or the move operation fails.

        Notes
        -----
        Validate the returned content using the MoveFile Pydantic model.
        """
        self._logger.info(msg="Moving a file from one folder to another within the same drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)
        self._logger.info(msg=target_path)

        # Source file: Uses the get_file_info function to obtain the source file_id
        file_info_response = self.get_file_info(drive_id=drive_id, filename=filename, save_as=None)

        if file_info_response.status_code != 200:
            content = None
            return self.Response(status_code=file_info_response.status_code, content=content)

        # Destination folder: Uses the get_dir_info function to obtain the source folder_id
        dir_info_response = self.get_dir_info(drive_id=drive_id, path=target_path, save_as=None)

        if dir_info_response.status_code != 200:
            content = None
            return self.Response(status_code=dir_info_response.status_code, content=content)

        # Do the move
        file_id = file_info_response.content["id"]
        folder_id = dir_info_response.content["id"]

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/items/{file_id}"

        # Request body
        data = {"parentReference": {"id": folder_id}}
        # Add to the request body if new_name is provided
        if new_name is not None:
            data["name"] = new_name  # type: ignore

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in MoveFile.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.patch(url=url_query, headers=headers, params=params, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=MoveFile, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def delete_file(self, drive_id: str, filename: str) -> Response:
        """
        Delete a file from a specified drive.

        Send a DELETE request to the Microsoft Graph API to remove a file at the given path within the specified drive.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the file to delete.
        filename : str
            Specify the full path of the file to delete, including the filename.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes details of the
            deleted file.

        Notes
        -----
        Return a successful HTTP status code if the file is deleted.
        """
        self._logger.info(msg="Deleting a file from the specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        filename_quote = quote(string=filename)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"

        # Send request
        response = self._session.delete(url=url_query, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 204):
            self._logger.info(msg="Request successful")

        return self.Response(status_code=response.status_code, content=content)

    def rename_file(self, drive_id: str, filename: str, new_name: str, save_as: str | None = None) -> Response:
        """
        Rename a file in a specified drive.

        Send a PATCH request to the Microsoft Graph API to rename a file at the given path within the specified drive.
        If the request is successful, return the HTTP status code and the details of the renamed file.
        Optionally, save the response to a JSON file.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the file to rename.
        filename : str
            Specify the full path of the file to rename, including the filename.
        new_name : str
            Specify the new name for the file.
        save_as : str or None, optional
            Specify the file path to save the response in JSON format. If None, do not save the response.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes the details of the
            renamed file.

        Examples
        --------
        >>> sp = SharePoint(client_id, tenant_id, client_secret, sp_domain)
        >>> resp = sp.rename_file(drive_id="drive_id", filename="old_name.txt", new_name="new_name.txt")
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Renaming a file in the specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)
        self._logger.info(msg=new_name)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        filename_quote = quote(string=filename)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"

        # Request body
        data = {"name": new_name}

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in RenameFile.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.patch(url=url_query, headers=headers, params=params, json=data, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=RenameFile, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def download_file(self, drive_id: str, remote_path: str, local_path: str) -> Response:
        """
        Download a file from a specified remote path in a SharePoint drive to a local path.

        Send a request to the Microsoft Graph API to download a file located at the specified remote path within the
        given drive ID. Save the file to the specified local path on the machine running the code.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the drive containing the file.
        remote_path : str
            Specify the path of the file in the SharePoint drive, including the filename.
        local_path : str
            Specify the local file path where the downloaded file will be saved.

        Returns
        -------
        Response
            Return an instance of the Response class containing the HTTP status code and content indicating the result
            of the download operation.

        Notes
        -----
        If the request is successful, write the file to disk.
        """
        self._logger.info(msg="Downloading a file from the specified remote path in the drive to the local path")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=remote_path)
        self._logger.info(msg=local_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}"}

        # Endpoint
        remote_path_quote = quote(string=remote_path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"

        # Send request
        response = self._session.get(url=url_query, headers=headers, stream=True, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Create file
            with open(file=local_path, mode="wb") as file:
                for chunk in response.iter_content(chunk_size=1024):
                    file.write(chunk)

        return self.Response(status_code=response.status_code, content=content)

    def download_file_to_memory(self, drive_id: str, remote_path: str) -> Response:
        """
        Download a file from a specified remote path in a SharePoint drive to memory.

        Parameters
        ----------
        drive_id : str
            ID of the drive containing the file.
        remote_path : str
            Path of the file in the SharePoint drive, including the filename.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the file content as bytes.

        Notes
        -----
        Store the file content in memory as bytes. Use with caution for large files, as this may consume significant
        memory.
        """
        self._logger.info(msg="Downloading a file from the specified remote path in the drive to memory")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=remote_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}"}

        # Endpoint
        remote_path_quote = quote(string=remote_path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"

        # Send request
        response = self._session.get(url=url_query, headers=headers, stream=True, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")
            content = b"".join(response.iter_content(chunk_size=1024))
            file_size = len(content)
            self._logger.info(msg=f"{file_size} bytes downloaded")

        return self.Response(status_code=response.status_code, content=content)

    def download_all_files(self, drive_id: str, remote_path: str, local_path: str) -> Response:
        """
        Download all files from a specified folder in SharePoint to a local directory.

        List all files in the given SharePoint folder and download each file with an extension to the specified local
        directory. Log the status of each download and return a summary of the results.

        Parameters
        ----------
        drive_id : str
            Specify the ID of the SharePoint drive.
        remote_path : str
            Specify the path of the folder in SharePoint.
        local_path : str
            Specify the local directory where files will be saved.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and a list of download results for each file.

        Notes
        -----
        Only download files with an extension. Each result includes file metadata and download status.
        """
        self._logger.info(msg="Initiating the process of downloading all files from the specified folder")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=remote_path)
        self._logger.info(msg=local_path)

        # List all items in the folder
        response = self.list_dir(drive_id=drive_id, path=remote_path)
        if response.status_code != 200:
            self._logger.error(msg="Failed to list folder contents")
            return self.Response(status_code=response.status_code, content=None)

        # Output
        items = response.content
        content = []
        if response.status_code == 200:
            for item in items:
                # Only files with extension
                if item.get("extension") is None:
                    continue

                filename = item.get("name")
                self._logger.info(msg=f"File {filename}")

                # Download file
                sub_response = self.download_file(
                    drive_id=drive_id,
                    remote_path=rf"{remote_path}/{filename}",
                    local_path=rf"{local_path}/{filename}",
                )

                # Status
                status = "pass" if sub_response.status_code == 200 else "fail"
                if status == "pass":
                    self._logger.info(msg="File downloaded successfully")
                else:
                    self._logger.warning(msg=f"Failed to download {filename}")

                content.append(
                    {
                        "id": item.get("id"),
                        "name": filename,
                        "extension": item.get("extension"),
                        "size": item.get("size"),
                        "path": item.get("path"),
                        "created_date_time": item.get("created_date_time"),
                        "last_modified_date_time": item.get("last_modified_date_time"),
                        "last_modified_by_name": item.get("last_modified_by_name"),
                        "last_modified_by_email": item.get("last_modified_by_email"),
                        "status": status,
                    }
                )

        return self.Response(status_code=response.status_code, content=content)

    def upload_file(self, drive_id: str, local_path: str, remote_path: str, save_as: str | None = None) -> Response:
        """
        Upload a file to a specified remote path in a SharePoint drive.

        Upload a file from the local file system to the specified remote path in a SharePoint drive using the Microsoft
        Graph API.
        Create the target folder in SharePoint if it does not exist. Return the HTTP status code and a response
        indicating the result of the operation.

        Parameters
        ----------
        drive_id : str
            ID of the drive where to upload the file.
        local_path : str
            Local file path of the file to upload.
        remote_path : str
            Path in the SharePoint drive where to upload the file, including the filename.
        save_as : str or None, optional
            If provided, save the results to a JSON file at the specified path.

        Returns
        -------
        Response
            Response object containing the HTTP status code and content indicating the result of the upload operation.

        Examples
        --------
        >>> resp = sp.upload_file(
        ...     drive_id="drive_id",
        ...     local_path="/tmp/example.txt",
        ...     remote_path="Documents/example.txt"
        ... )
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Uploading a file to the specified remote path in the drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=local_path)
        self._logger.info(msg=remote_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/octet-stream",
        }

        # Endpoint
        remote_path_quote = quote(string=remote_path)
        url_query = rf"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in UploadFile.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        data = open(file=local_path, mode="rb").read()

        # Send request
        response = self._session.put(url=url_query, headers=headers, params=params, data=data, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 201):
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=UploadFile, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    # LISTS
    def list_lists(self, site_id: str, save_as: str | None = None) -> Response:
        """
        Retrieve SharePoint lists for a specified site.

        Send a request to the Microsoft Graph API to obtain details about lists within the given site ID.
        Return the list information and HTTP status code. Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            Specify the ID of the site containing the lists.
        save_as : str, optional
            Specify the file path to save the response in JSON format. If not provided, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the content, which includes list details such as ID,
            name, display name, description, web URL, created date, and last modified date.

        Examples
        --------
        >>> resp = sp.list_lists(site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1")
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Retrieving a list of lists for the specified site")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/lists"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in ListLists.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = self._handle_response(response=response, model=ListLists, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def list_list_columns(self, site_id: str, list_id: str, save_as: str | None = None) -> Response:
        """
        Retrieve columns from a specified SharePoint list.

        Send a request to the Microsoft Graph API to get columns for the given list ID within a site.
        Return column details and HTTP status code. Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            ID of the site containing the list.
        list_id : str
            ID of the list for which to retrieve columns.
        save_as : str, optional
            File path to save the response in JSON format. If not provided, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and content with column details such as ID, name,
            display name, description, column group, enforce unique values, hidden, indexed, read-only, and required.

        Examples
        --------
        >>> resp = sp.list_list_columns(
        ...     site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1",
        ...     list_id="e11f111b-1111-11a1-1111-11f11d1a11f1"
        ... )
        >>> print(resp.status_code)
        >>> print(resp.content)
        """

        self._logger.info(msg="Retrieving columns from the specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/columns"

        # Query parameters
        # Pydantic v1
        alias_list = [
            field.alias for field in ListListColumns.__fields__.values() if field.field_info.alias is not None
        ]
        params = {"$select": ",".join(alias_list)}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = self._handle_response(response=response, model=ListListColumns, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def list_list_items(self, site_id: str, list_id: str, fields: dict, save_as: str | None = None) -> Response:
        """
        Retrieve items from a specified SharePoint list.

        Send a request to the Microsoft Graph API to obtain items for the given list ID within a site.
        Return item details and HTTP status code. Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            Specify the ID of the site containing the list.
        list_id : str
            Specify the ID of the list for which to retrieve items.
        fields : dict
            Specify the fields to retrieve for each item in the list.
        save_as : str, optional
            Specify the file path to save the response in JSON format. If not provided, do not save the response.

        Returns
        -------
        Response
            Return a Response object containing the HTTP status code and the content, which includes item details such
            as ID, title, description, etc.

        Examples
        --------
        >>> resp = sp.list_list_items(
        ...     site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1",
        ...     list_id="e11f111b-1111-11a1-1111-11f11d1a11f1",
        ...     fields="fields/Id,fields/Title,fields/Description"
        ... )
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Retrieving items from the specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json;odata.metadata=none",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items"

        # Query parameters
        params = {"select": fields, "expand": "fields"}

        # Send request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json
            content = [item["fields"] for item in response.json()["value"]]

        return self.Response(status_code=response.status_code, content=content)

    def delete_list_item(self, site_id: str, list_id: str, item_id: str) -> Response:
        """
        Delete an item from a SharePoint list.

        Send a DELETE request to the Microsoft Graph API to remove the specified item from a list.

        Parameters
        ----------
        site_id : str
            ID of the site containing the list.
        list_id : str
            ID of the list containing the item.
        item_id : str
            ID of the item to delete.

        Returns
        -------
        Response
            Response object containing the HTTP status code.

        Examples
        --------
        >>> resp = sp.delete_list_item(
        ...     site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1",
        ...     list_id="e11f111b-1111-11a1-1111-11f11d1a11f1",
        ...     item_id="1"
        ... )
        >>> print(resp.status_code)
        """
        self._logger.info(msg="Deleting the specified item from the list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)
        self._logger.info(msg=item_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items/{item_id}"

        # Send request
        response = self._session.delete(url=url_query, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 204:
            self._logger.info(msg="Request successful")

        return self.Response(status_code=response.status_code, content=content)

    def add_list_item(self, site_id: str, list_id: str, item: dict, save_as: str | None = None) -> Response:
        """
        Add a new item to a SharePoint list.

        Add a new item to the specified list in SharePoint using the Microsoft Graph API.
        Return the details of the added item and the HTTP status code. Optionally, save the response to a JSON file.

        Parameters
        ----------
        site_id : str
            ID of the site containing the list.
        list_id : str
            ID of the list to which the item will be added.
        item : dict
            Item data to add to the list. (e.g. {"Title": "Hello World"})
        save_as : str, optional
            File path to save the response in JSON format. If not provided, do not save the response.

        Returns
        -------
        Response
            Response object containing the HTTP status code and the content, which includes the details of the added
            list item.

        Examples
        --------
        >>> resp = sp.add_list_item(
        ...     site_id="companygroup.sharepoint.com,1111a11e-...,...-ed11bff1baf1",
        ...     list_id="e11f111b-1111-11a1-1111-11f11d1a11f1",
        ...     item={"Title": "Hello World"}
        ... )
        >>> print(resp.status_code)
        >>> print(resp.content)
        """
        self._logger.info(msg="Adding a new item to the specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        }

        # Endpoint
        url_query = rf"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items"

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in AddListItem.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        # @microsoft.graph.conflictBehavior: fail, rename, replace
        data = {"fields": item}

        # Send request
        response = self._session.post(url=url_query, headers=headers, json=data, params=params, verify=True)
        # print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 201:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=AddListItem, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

# eof
