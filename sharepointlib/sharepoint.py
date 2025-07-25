"""
This library provides access to some SharePoint functionalities.
Its features are designed to remain generic and should not be modified to meet the specific needs of individual projects.

Get single base models values:
Example:
print(response.content.id)

Convert list of base models to pandas dataframes:
Example:
df = pd.DataFrame(response.content)
display(df)
"""

# import base64
import dataclasses
from datetime import datetime
import json
import logging
# import pathlib
# TypeAdapter v2 vs parse_obj_as v1
from pydantic import BaseModel, Field, parse_obj_as, validator
import requests
from typing import Any, Type
from urllib.parse import quote

# Creates a logger for this module
logger = logging.getLogger(__name__)

class SharePoint(object):
    @dataclasses.dataclass
    class Configuration(object):
        api_domain: str | None = None
        api_version: str | None = None
        sp_domain: str | None = None
        client_id: str | None = None
        tenant_id: str | None = None
        client_secret: str | None = None
        token: str | None = None
    
    @dataclasses.dataclass
    class Response:
        status_code: int
        content: Any = None
    
    def __init__(self, client_id: str, tenant_id: str, client_secret: str, sp_domain: str, logger: logging.Logger | None = None) -> None:
        """
        Initializes the SharePoint client with the provided credentials and configuration.

        Args:
            client_id (str): The Azure client ID used for authentication.
            tenant_id (str): The Azure tenant ID associated with the client.
            client_secret (str): The secret key for the Azure client.
            sp_domain (str): The SharePoint domain.
                              Example: "companygroup.sharepoint.com"
            logger (logging.Logger, optional): Logger instance to use. If None, a default logger is created.
        """
        # Init logging
        # Use provided logger or create a default one
        self._logger = logger or logging.getLogger(name=__name__)

        # Init variables
        self._session: requests.Session = requests.Session()
        api_domain = "graph.microsoft.com"
        api_version = "v1.0"

        # Credentials/Configuration
        self._configuration = self.Configuration(api_domain=api_domain,
                                                 api_version=api_version,
                                                 sp_domain=sp_domain,
                                                 client_id=client_id, 
                                                 tenant_id=tenant_id,
                                                 client_secret=client_secret,
                                                 token=None)
        
        # Authenticate
        self.auth()
    
    def __del__(self) -> None:
        """
        Cleans the house at the exit.
        """
        self._logger.info(msg="Cleans the house at the exit")
        self._session.close()
    
    def auth(self) -> None:
        """
        Authentication.
        This method performs the authentication process to obtain an access token
        using the client credentials flow. The token is stored in the Configuration
        dataclass for subsequent API requests.
        """
        self._logger.info(msg="Authentication")

        # Request headers
        headers = {"Content-Type": "application/x-www-form-urlencoded"}

        # Authorization URL
        url_auth = f"https://login.microsoftonline.com/{self._configuration.tenant_id}/oauth2/v2.0/token"

        # Request body
        body = {"grant_type": "client_credentials",
                "client_id": self._configuration.client_id,
                "client_secret": self._configuration.client_secret,
                "scope": "https://graph.microsoft.com/.default"}

        # Request
        response = self._session.post(url=url_auth, data=body, headers=headers, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Return valid response
        if response.status_code == 200:
            self._configuration.token = json.loads(response.content.decode("utf-8"))["access_token"]
    
    def _export_to_json(self, content: bytes, save_as: str | None) -> None:
        """
        Export response content to a JSON file.

        This method takes the content to be exported and saves it to a specified file in JSON format.
        If the `save_as` parameter is provided, the content will be written to that file.

        Args:
            content (bytes): The content to be exported, typically the response content from an API call.
            save_as (str): The file path where the JSON content will be saved. If None, the content will not be saved.
        """
        if save_as is not None:
            self._logger.info(msg="Exports response to JSON file.")
            with open(file=save_as, mode="wb") as file:
                file.write(content)
    
    def _handle_response(self, response: requests.Response, model: Type[BaseModel], rtype: str = "scalar") -> BaseModel | list[BaseModel]:
        """
        Handles the response from an API request and deserializes the JSON content.

        This method processes the response from an API request and deserializes the JSON content
        that contains a single instance of a Pydantic BaseModel (scalar) or a list of instances (list).

        Args:
            response (requests.Response): The response object from the API request.
            model (BaseModel): The Pydantic BaseModel class to use for deserialization and validation.
            rtype (str, optional): The type of response to handle. Options are "scalar" for a single record
                                   and "list" for a list of records. Defaults to "scalar".

        Returns:
            BaseModel | list[BaseModel]: The deserialized content as a Pydantic BaseModel instance or a list of instances.
        """
        if rtype.lower() == "scalar": 
            # Deserialize json (scalar values)
            content_raw = response.json()
            # Pydantic v1
            return model(**content_raw)
        else:
            # Deserialize json
            content_raw = response.json()["value"]
            # Pydantic v1
            return parse_obj_as(list[model], content_raw)
            # return [dict(data) for data in parse_obj_as(list[model], content_raw)]

    def get_site_info(self, name: str, save_as: str | None = None) -> Response:
        """
        Gets the site ID for a given site name.

        This method sends a request to the Microsoft Graph API to retrieve the site ID
        associated with the specified site name. If the request is successful, it will
        return the site ID along with the HTTP status code. Optionally, the response can
        be saved to a JSON file.

        Args:
            name (str): The name of the site for which to retrieve the site ID.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the site ID and other relevant information.
        """
        self._logger.info(msg="Gets the site ID for a given site name")
        self._logger.info(msg=name)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        sp_domain = self._configuration.sp_domain

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        # get_sites_id: url_query = f"https://graph.microsoft.com/v1.0/sites?search='{filter}'"
        url_query = fr"https://{api_domain}/{api_version}/sites/{sp_domain}:/sites/{name}"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            display_name: str = Field(None, alias="displayName")
            web_url: str = Field(None, alias="webUrl")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)
    
    def get_hostname_info(self, site_id: str, save_as: str | None = None) -> Response:
        """
        Gets the hostname and site details for a specified site ID.

        This method sends a request to the Microsoft Graph API to retrieve the hostname,
        site name, and other relevant details associated with the specified site ID. 
        If the request is successful, it will return the site information along with 
        the HTTP status code. Optionally, the response can be saved to a JSON file.

        Args:
            site_id (str): The ID of the site for which to retrieve the hostname and details.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the hostname, site name, and other relevant information.
        """
        self._logger.info(msg="Gets the hostname and site details for a specified site ID")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            display_name: str = Field(None, alias="displayName")
            description: str = Field(None, alias="description")
            web_url: str = Field(None, alias="webUrl")
            site_collection: dict = Field(None, alias="siteCollection")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")
        
        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)
    
    # DRIVES
    def list_drives(self, site_id: str, save_as: str | None = None) -> Response:
        """
        Gets a list of the Drive IDs for a given site ID.

        This method sends a request to the Microsoft Graph API to retrieve the Drive IDs
        associated with the specified site ID. If the request is successful, it will return
        the Drive IDs along with the HTTP status code. Optionally, the response can be saved
        to a JSON file.

        Args:
            site_id (str): The ID of the site for which to retrieve the Drive IDs.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the list of Drive IDs and other relevant information.
        """
        self._logger.info(msg="Gets a list of the Drive IDs for a given site")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        
        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/drives"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            description: str = Field(None, alias="description")
            web_url: str = Field(None, alias="webUrl")
            drive_type: str = Field(None, alias="driveType")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="list")

        return self.Response(status_code=response.status_code, content=content)
    
    def get_dir_info(self, drive_id: str, path: str | None = None, save_as: str | None = None) -> Response:
        """
        Gets the folder ID for a specified folder within a drive ID.

        This method sends a request to the Microsoft Graph API to retrieve the folder ID
        associated with the specified folder path within the given drive ID. If the request
        is successful, it will return the folder ID along with the HTTP status code. 
        Optionally, the response can be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive where the folder is located.
            path (str, optional): The path of the folder for which to retrieve the folder ID.
                                   If not provided, the root folder will be used.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the folder ID and other relevant information.
        """
        self._logger.info(msg="Gets the folder ID for a specified folder within a drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        path_quote = "///" if path is None else f"/{quote(string=path)}"
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:{path_quote}"
        # print(url_query)

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            web_url: str = Field(None, alias="webUrl")
            size: int = Field(None, alias="size")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def list_dir(self, drive_id: str, path: str | None = None, save_as: str | None = None) -> Response:
        """
        List content (files and folders) of a specific folder.

        This method sends a request to the Microsoft Graph API to retrieve the list of 
        children (files and folders) within a specified folder in a drive. If the request 
        is successful, it will return the HTTP status code and a list of the children.

        Args:
            drive_id (str): The ID of the drive containing the folder.
            path (str, optional): The path of the folder for which to list the children.
                                   If not provided, the root folder will be used.
            save_as (str, optional): If provided, the results will be saved to a JSON file 
                                      at the specified path.

        Returns:
            Response: An instance of the Response class containing the HTTP status code 
                      and a list of children (files and folders) within the specified folder.
        """
        self._logger.info(msg="List content (files and folders) of a folder")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        path_quote = "/" if path is None else f"{quote(string=path)}"
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/items/root:/{path_quote}:/children"
        # print(url_query)

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            extension: str | None = None
            size: int = Field(None, alias="size")
            web_url: str = Field(None, alias="webUrl")
            folder: dict = Field(None, alias="folder")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")
            last_modified_by: dict = Field(None, alias="lastModifiedBy")
            last_modified_by_email: str | None = None

            @validator("extension", always=True)
            def set_extension(cls, v, values):
                if values.get("folder") is None:
                    return values["name"].split(".")[-1] if "." in values["name"] else None
                return None
            
            @validator("last_modified_by_email", pre=True, always=True)
            def set_last_modified_by_email(cls, v, values):
                # Handle cases where lastModifiedBy or user.email is missing
                last_modified_by = values.get("last_modified_by")
                if last_modified_by and "user" in last_modified_by and "email" in last_modified_by["user"]:
                    return last_modified_by["user"]["email"]
                return None
            
            # Exclude last_modified_by from dict() method
            def dict(self, *args, **kwargs):
                # Override dict() to exclude last_modified_by from output
                kwargs.setdefault("exclude", {"last_modified_by"})
                return super().dict(*args, **kwargs)

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="list")

        return self.Response(status_code=response.status_code, content=content)

    def create_dir(self, drive_id: str, path: str, name: str, save_as: str | None = None) -> Response:
        """
        Creates a new folder in a specified drive ID.

        This method sends a request to the Microsoft Graph API to create a new folder 
        at the specified path within the given drive ID. If the request is successful, 
        it will return the HTTP status code and the details of the created folder. 
        Optionally, the response can be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive where the folder will be created.
            path (str): The path of the new folder to be created.
            name (str): The name of the new folder to be created.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the details of the created folder.
        """
        self._logger.info(msg="Creates a new folder in a specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        path_quote = quote(string=path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}:/children"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            web_url: str = Field(None, alias="webUrl")
            created_date_time: datetime = Field(alias="createdDateTime")
        
        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        # @microsoft.graph.conflictBehavior: fail, rename, replace
        data = {"name": name,
                "folder": {},
                "@microsoft.graph.conflictBehavior": "replace"}
        
        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")
        
        return self.Response(status_code=response.status_code, content=content)
    
    def delete_dir(self, drive_id: str, path: str) -> Response:
        """
        Deletes a folder from a specified drive ID.

        This method sends a request to the Microsoft Graph API to delete a folder located at the 
        specified path within the given drive ID. If the request is successful, it will return 
        the HTTP status code and the details of the deleted folder. Optionally, the response can 
        be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive containing the folder to be deleted.
            path (str): The full path of the folder to be deleted.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the details of the deleted folder.
        """
        self._logger.info(msg="Deletes a folder from a specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        path_quote = quote(string=path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}"

        # Request
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
        Renames a folder in a specified drive ID.
        This method sends a PATCH request to the Microsoft Graph API to rename a folder.

        Args:
            drive_id (str): The ID of the drive containing the folder to be renamed.
            path (str): The full path of the folder to be renamed.
            new_name (str): The new name for the folder.
            save_as (str, optional): The file path where the response will be saved in JSON format.

        Returns:
            Response: A Response dataclass instance with the status code and content.
        """
        self._logger.info(msg="Renames a folder in a specified drive")
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
            "Accept": "application/json"
        }

        # Request query
        path_quote = quote(string=path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{path_quote}"

        # Request body
        data = {"name": new_name}

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            web_url: str = Field(None, alias="webUrl")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)

    def get_file_info(self, drive_id: str, filename: str, save_as: str | None = None) -> Response:
        """
        Retrieves information about a specific file in a drive ID.

        This method sends a request to the Microsoft Graph API to obtain details about a file 
        located at the specified path within the given drive ID. If the request is successful, 
        it will return the file information along with the HTTP status code. Optionally, the 
        response can be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive containing the file.
            filename (str): The full path of the file for which to retrieve information, including the filename.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the file details such as ID, name, web URL, size,
                      created date, and last modified date.
        """
        self._logger.info(msg="Retrieves information about a specific file")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        filename_quote = quote(string=filename)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            web_url: str = Field(None, alias="webUrl")
            size: int = Field(None, alias="size")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")
            last_modified_by: dict = Field(None, alias="lastModifiedBy")
            last_modified_by_email: str | None = None

            @validator("last_modified_by_email", pre=True, always=True)
            def set_last_modified_by_email(cls, v, values):
                # Handle cases where lastModifiedBy or user.email is missing
                last_modified_by = values.get("last_modified_by")
                if last_modified_by and "user" in last_modified_by and "email" in last_modified_by["user"]:
                    return last_modified_by["user"]["email"]
                return None
            
            # Exclude last_modified_by from dict() method
            def dict(self, *args, **kwargs):
                # Override dict() to exclude last_modified_by from output
                kwargs.setdefault("exclude", {"last_modified_by"})
                return super().dict(*args, **kwargs)

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")
        
        return self.Response(status_code=response.status_code, content=content)

    def copy_file(self, drive_id: str, filename: str, target_path: str) -> Response:
        """
        Copy a file from one folder to another within the same drive ID.

        This method sends a request to the Microsoft Graph API to copy a file from the specified 
        source path to the destination path within the given drive ID. It works only on the same drive. 
        If the request is successful, it will return the HTTP status code and the details of the moved file.

        Args:
            drive_id (str): The ID of the drive containing the file to be moved.
            filename (str): The full path of the file to be moved, including the filename.
            target_path (str): The path of the destination folder where the file will be moved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code.
        """
        self._logger.info(msg="Copy a file from one folder to another within the same drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)
        self._logger.info(msg=target_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        filename_quote = quote(string=filename)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}:/copy"

        # Request body
        data = {"parentReference": {"driveId": drive_id,
                                    "driveType": "documentLibrary",
                                    "path": f"/drives/{drive_id}/root:/{target_path}"}}

        # Request
        response = self._session.post(url=url_query, headers=headers, json=data, verify=True)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 202):
            self._logger.info(msg="Request successful")
    
        return self.Response(status_code=response.status_code, content=content)
    
    def move_file(self, drive_id: str, filename: str, target_path: str, save_as: str | None = None) -> Response:
        """
        Moves a file from one folder to another within the same drive ID.

        This method sends a request to the Microsoft Graph API to move a file from the specified 
        source path to the destination path within the given drive ID. It works only on the same drive. 
        If the request is successful, it will return the HTTP status code and the details of the moved file. 
        Optionally, the response can be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive containing the file to be moved.
            filename (str): The full path of the file to be moved, including the filename.
            target_path (str): The path of the destination folder where the file will be moved.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the details of the moved file.
        """
        self._logger.info(msg="Moves a file from one folder to another within the same drive")
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
        file_id = file_info_response.content.id
        folder_id = dir_info_response.content.id

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/items/{file_id}"

        # Request body
        data = {"parentReference": {"id": folder_id}}

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            web_url: str = Field(None, alias="webUrl")
            size: int = Field(None, alias="size")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")
    
        return self.Response(status_code=response.status_code, content=content)

    def delete_file(self, drive_id: str, filename: str) -> Response:
        """
        Deletes a file from a specified drive ID.

        This method sends a request to the Microsoft Graph API to delete a file located at the 
        specified path within the given drive ID. If the request is successful, it will return 
        the HTTP status code and the details of the deleted file. Optionally, the response can 
        be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive containing the file to be deleted.
            filename (str): The full path of the file to be deleted, including the filename.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the details of the deleted file.
        """
        self._logger.info(msg="Deletes a file from a specified drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=filename)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        filename_quote = quote(string=filename)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"
        
        # Request
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
        Renames a file in a specified drive ID.

        This method sends a request to the Microsoft Graph API to rename a file located at the
        specified path within the given drive ID. If the request is successful, it will return
        the HTTP status code and the details of the renamed file. Optionally, the response can
        be saved to a JSON file.

        Args:
            drive_id (str): The ID of the drive containing the file to be renamed.
            filename (str): The full path of the file to be renamed, including the filename.
            new_name (str): The new name for the file.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                    the content, which includes the details of the renamed file.
        """
        self._logger.info(msg="Renames a file in a specified drive")
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
            "Accept": "application/json"
        }

        # Request query
        filename_quote = quote(string=filename)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{filename_quote}"

        # Request body
        data = {"name": new_name}

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            web_url: str = Field(None, alias="webUrl")
            size: int = Field(None, alias="size")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
        response = self._session.patch(url=url_query, headers=headers, params=params, json=data, verify=True)
        print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 200:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")

        return self.Response(status_code=response.status_code, content=content)
    
    def download_file(self, drive_id: str, remote_path: str, local_path: str) -> Response:
        """
        Downloads a file from a specified remote path in a drive ID to a local path.

        This method sends a request to the Microsoft Graph API to download a file located at the 
        specified remote path within the given drive ID. The file will be saved to the specified 
        local path on the machine running the code. If the request is successful, it will return 
        the HTTP status code and a response indicating the result of the operation.

        Args:
            drive_id (str): The ID of the drive containing the file.
            remote_path (str): The path of the file in the SharePoint drive, including the filename.
            local_path (str): The local file path where the downloaded file will be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code 
                      and content indicating the result of the download operation.
        """
        self._logger.info(msg="Downloads a file from a specified remote path in a drive to a local path")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=remote_path)
        self._logger.info(msg=local_path)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}"}
        
        # Request query
        remote_path_quote = quote(string=remote_path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"
        
        # Request
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
        Downloads a file from a specified remote path in a drive ID to a variable (memory).

        Note that large files will require a significant amount of memory!

        This method sends a request to the Microsoft Graph API to download a file located at the 
        specified remote path within the given drive ID. The file content is stored in a variable
        and returned as part of the Response object. If the request is successful, it will return 
        the HTTP status code and the file content.

        Args:
            drive_id (str): The ID of the drive containing the file.
            remote_path (str): The path of the file in the SharePoint drive, including the filename.

        Returns:
            Response: An instance of the Response class containing the HTTP status code 
                    and the content, which includes the file data as bytes.
        """
        self._logger.info(msg="Downloads a file from a specified remote path in a drive to a variable")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=remote_path)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}"}
        
        # Request query
        remote_path_quote = quote(string=remote_path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"
        
        # Request
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
    
    def upload_file(self, drive_id: str, local_path: str, remote_path: str, save_as: str | None = None) -> Response:
        """
        Uploads a file to a specified remote path in a SharePoint drive ID.

        This method sends a request to the Microsoft Graph API to upload a file from the local path 
        to the specified remote path within the given drive ID. If the folder does not exist in 
        SharePoint, it will be created. If the request is successful, it will return the HTTP status 
        code and a response indicating the result of the operation.

        Args:
            drive_id (str): The ID of the drive where the file will be uploaded.
            local_path (str): The local file path of the file to be uploaded.
            remote_path (str): The path in the SharePoint drive where the file will be uploaded, including the filename.
            save_as (str, optional): If provided, the results will be saved to a JSON file 
                                      at the specified path.

        Returns:
            Response: An instance of the Response class containing the HTTP status code 
                      and content indicating the result of the upload operation.
        """
        self._logger.info(msg="Uploads a file to a specified remote path in a drive")
        self._logger.info(msg=drive_id)
        self._logger.info(msg=local_path)
        self._logger.info(msg=remote_path)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/octet-stream"}
        
        # Request query
        remote_path_quote = quote(string=remote_path)
        url_query = fr"https://{api_domain}/{api_version}/drives/{drive_id}/root:/{remote_path_quote}:/content"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            size: int = Field(None, alias="size")
        
        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        data = open(file=local_path, mode="rb").read()

        # Request
        response = self._session.put(url=url_query, headers=headers, params=params, data=data, verify=True)
        print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code in (200, 201):
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")
        
        return self.Response(status_code=response.status_code, content=content)

    # LISTS
    def list_lists(self, site_id: str, save_as: str | None = None) -> Response:
        
        """
        Retrieves a list of SharePoint lists for a specified site.

        This method sends a request to the Microsoft Graph API to obtain details about the lists 
        within the given site ID. If the request is successful, it will return the list information 
        along with the HTTP status code. Optionally, the response can be saved to a JSON file.

        Example: list_lists(site_id="companygroup.sharepoint.com,1111a11e-f1bb-1111-b11f-a1111b11b1b0,db1111a1-11e1-1d1c-1111-ed11bff1baf1")

        Args:
            site_id (str): The ID of the site containing the lists.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the list details such as ID, name, display name,
                      description, web URL, created date, and last modified date.
        """
        self._logger.info(msg="Retrieves a list of lists for a specified site")
        self._logger.info(msg=site_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        
        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/lists"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            display_name: str = Field(None, alias="displayName")
            description: str = Field(None, alias="description")
            web_url: str = Field(None, alias="webUrl")
            created_date_time: datetime = Field(alias="createdDateTime")
            last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
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
            content = self._handle_response(response=response, model=DataStructure, rtype="list")

        return self.Response(status_code=response.status_code, content=content)
    
    def list_list_columns(self, site_id: str, list_id: str, save_as: str | None = None) -> Response:
        """
        Retrieves the columns from a specified list in SharePoint.

        This method sends a request to the Microsoft Graph API to retrieve the columns
        associated with the specified list ID within a site. If the request is successful,
        it will return the column details along with the HTTP status code. Optionally, 
        the response can be saved to a JSON file.

        Example: list_list_columns(site_id="companygroup.sharepoint.com,1111a11e-f1bb-1111-b11f-a1111b11b1b0,db1111a1-11e1-1d1c-1111-ed11bff1baf1", 
                                   list_id="e11f111b-1111-11a1-1111-11f11d1a11f1")

        Args:
            site_id (str): The ID of the site containing the list.
            list_id (str): The ID of the list for which to retrieve the columns.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the column details such as ID, name, display name,
                      description, column group, enforce unique values, hidden, indexed, read-only, and required.
        """
        self._logger.info(msg="Retrieves the columns from a specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        
        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/columns"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(alias="name")
            display_name: str = Field(alias="displayName")
            description: str = Field(alias="description")
            column_group: str = Field(alias="columnGroup")
            enforce_unique_values: bool = Field(alias="enforceUniqueValues")
            hidden: bool = Field(alias="hidden")
            indexed: bool = Field(alias="indexed")
            read_only: bool = Field(alias="readOnly")
            required: bool = Field(alias="required")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request
        response = self._session.get(url=url_query, headers=headers, params=params, verify=True)  # params=params,
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
            content = self._handle_response(response=response, model=DataStructure, rtype="list")

        return self.Response(status_code=response.status_code, content=content)
    
    def list_list_items(self, site_id: str, list_id: str, fields: dict, save_as: str | None = None) -> Response:
        """
        Retrieves the items from a specified list in SharePoint.

        This method sends a request to the Microsoft Graph API to retrieve the items
        associated with the specified list ID within a site. If the request is successful,
        it will return the item details along with the HTTP status code. Optionally, 
        the response can be saved to a JSON file.

        Example: list_list_items(site_id="companygroup.sharepoint.com,1111a11e-f1bb-1111-b11f-a1111b11b1b0,db1111a1-11e1-1d1c-1111-ed11bff1baf1", 
                                 list_id="e11f111b-1111-11a1-1111-11f11d1a11f1", 
                                 fields="fields/Id,fields/Title,fields/Description")

        Args:
            site_id (str): The ID of the site containing the list.
            list_id (str): The ID of the list for which to retrieve the items.
            fields (dict): The fields to be retrieved for each item in the list.
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the item details such as ID, title, description, etc.
        """
        self._logger.info(msg="Retrieves the items from a specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)

        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version
        
        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json;odata.metadata=none"}

        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items"

        # Query parameters
        params = {"select": fields, "expand": "fields"}

        # Request
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
        Deletes a specified item from a list in SharePoint.

        This method sends a request to the Microsoft Graph API to delete the item
        associated with the specified item ID within a list. If the request is successful,
        it will return the HTTP status code. 

        Example: delete_list_item(site_id="companygroup.sharepoint.com,1111a11e-f1bb-1111-b11f-a1111b11b1b0,db1111a1-11e1-1d1c-1111-ed11bff1baf1", 
                                  list_id="e11f111b-1111-11a1-1111-11f11d1a11f1", 
                                  item_id="1")

        Args:
            site_id (str): The ID of the site containing the list.
            list_id (str): The ID of the list containing the item.
            item_id (str): The ID of the item to be deleted.

        Returns:
            Response: An instance of the Response class containing the HTTP status code.
        """
        self._logger.info(msg="Deletes a specified item from a list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)
        self._logger.info(msg=item_id)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items/{item_id}"
        
        # Request
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
        Adds a new item to a specified list in SharePoint.

        This method sends a request to the Microsoft Graph API to add a new item
        to the specified list ID within a site. If the request is successful,
        it will return the item details along with the HTTP status code. Optionally, 
        the response can be saved to a JSON file.

        Example: add_list_item(site_id="companygroup.sharepoint.com,1111a11e-f1bb-1111-b11f-a1111b11b1b0,db1111a1-11e1-1d1c-1111-ed11bff1baf1", 
                               list_id="e11f111b-1111-11a1-1111-11f11d1a11f1", 
                               item={"Title": "Hello World"})

        Args:
            site_id (str): The ID of the site containing the list.
            list_id (str): The ID of the list to which the item will be added.
            item (dict): The item data to be added to the list.
                          Example: {"Title": "Hello World"}
            save_as (str, optional): The file path where the response will be saved in JSON format.
                                      If not provided, the response will not be saved.

        Returns:
            Response: An instance of the Response class containing the HTTP status code and
                      the content, which includes the details of the added list item.
        """
        self._logger.info(msg="Adds a new item to a specified list")
        self._logger.info(msg=site_id)
        self._logger.info(msg=list_id)
        
        # Configuration
        token = self._configuration.token
        api_domain = self._configuration.api_domain
        api_version = self._configuration.api_version

        # Request headers
        headers = {"Authorization": f"Bearer {token}",
                   "Content-Type": "application/json",
                   "Accept": "application/json"}
        
        # Request query
        url_query = fr"https://{api_domain}/{api_version}/sites/{site_id}/lists/{list_id}/items"

        # Pydantic output data structure
        class DataStructure(BaseModel):
            id: str = Field(alias="id")
            name: str = Field(None, alias="name")
            web_url: str = Field(None, alias="webUrl")
            created_date_time: datetime = Field(alias="createdDateTime")
            # last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

        # Query parameters
        # Pydantic v1
        alias_list = [field.alias for field in DataStructure.__fields__.values() if field.field_info.alias is not None]
        params = {"$select": ",".join(alias_list)}

        # Request body
        # @microsoft.graph.conflictBehavior: fail, rename, replace
        data = {"fields": item}
        
        # Request
        response = self._session.post(url=url_query, headers=headers, json=data, params=params, verify=True)
        print(response.content)

        # Log response code
        self._logger.info(msg=f"HTTP Status Code {response.status_code}")

        # Output
        content = None
        if response.status_code == 201:
            self._logger.info(msg="Request successful")

            # Export response to json file
            self._export_to_json(content=response.content, save_as=save_as)

            # Deserialize json (scalar values)
            content = self._handle_response(response=response, model=DataStructure, rtype="scalar")
        
        return self.Response(status_code=response.status_code, content=content)

# eom
