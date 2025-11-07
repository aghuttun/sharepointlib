"""
Pydantic models for SharePoint entity representations and API responses.

This module defines data structures for SharePoint sites, drives, directories, files, and lists, enabling type
validation and serialization of SharePoint API responses using Pydantic.
"""

from datetime import datetime
from pydantic import BaseModel, Field, validator


class GetSiteInfo(BaseModel):
    """
    Represent the data structure for get_site_info() responses.

    Parameters
    ----------
    id : str
        The unique identifier of the SharePoint site.
    name : str, optional
        The name of the SharePoint site.
    display_name : str, optional
        The display name of the SharePoint site.
    web_url : str, optional
        The web URL of the SharePoint site.
    created_date_time : datetime
        The creation date and time of the SharePoint site.
    last_modified_date_time : datetime, optional
        The last modification date and time of the SharePoint site.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetHostNameInfo(BaseModel):
    """
    Define the data structure for get_hostname_info() responses.

    Parameters
    ----------
    id : str
        The unique identifier of the SharePoint host.
    name : str, optional
        The name of the SharePoint host.
    display_name : str, optional
        The display name of the SharePoint host.
    description : str, optional
        The description of the SharePoint host.
    web_url : str, optional
        The web URL of the SharePoint host.
    site_collection : dict, optional
        The site collection information of the SharePoint host.
    created_date_time : datetime
        The creation date and time of the SharePoint host.
    last_modified_date_time : datetime, optional
        The last modification date and time of the SharePoint host.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    site_collection: dict = Field(None, alias="siteCollection")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListDrives(BaseModel):
    """
    Represent the data structure for list_drives() responses.

    Define fields for SharePoint drive metadata, including identifiers, names, descriptions, URLs, drive types, and
    timestamps for creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint drive.
    name : str, optional
        Specify the name of the SharePoint drive.
    description : str, optional
        Specify the description of the SharePoint drive.
    web_url : str, optional
        Specify the web URL of the SharePoint drive.
    drive_type : str, optional
        Specify the type of the SharePoint drive.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint drive.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint drive.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    drive_type: str = Field(None, alias="driveType")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetDirInfo(BaseModel):
    """
    Represent the data structure for get_dir_info() responses.

    Define fields for SharePoint directory metadata, including identifiers, names, URLs, sizes, and timestamps for
    creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint directory.
    name : str, optional
        Specify the name of the SharePoint directory.
    web_url : str, optional
        Specify the web URL of the SharePoint directory.
    size : int, optional
        Specify the size of the SharePoint directory in bytes.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint directory.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint directory.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListDir(BaseModel):
    """
    Represent the data structure for list_dir() responses.

    Define fields for SharePoint directory listing metadata, including identifiers, names, extensions, sizes, paths,
    URLs, folder information, timestamps, and user modification details.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint directory or file.
    name : str
        Specify the name of the SharePoint directory or file.
    extension : str, optional
        Specify the file extension if the item is a file; otherwise, None.
    size : int, optional
        Specify the size of the directory or file in bytes.
    path : str, optional
        Specify the path of the directory or file.
    web_url : str, optional
        Specify the web URL of the directory or file.
    folder : dict, optional
        Specify folder metadata if the item is a directory; otherwise, None.
    created_date_time : datetime
        Specify the creation date and time of the directory or file.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the directory or file.
    last_modified_by : dict, optional
        Specify metadata about the user who last modified the item.
    last_modified_by_name : str, optional
        Specify the display name of the user who last modified the item.
    last_modified_by_email : str, optional
        Specify the email of the user who last modified the item.
    """

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    extension: str | None = None
    size: int = Field(None, alias="size")
    path: str | None = None
    web_url: str = Field(None, alias="webUrl")
    folder: dict = Field(None, alias="folder")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")
    last_modified_by: dict = Field(None, alias="lastModifiedBy")
    last_modified_by_name: str | None = None
    last_modified_by_email: str | None = None

    @validator("extension", pre=True, always=True)
    def set_extension(cls, v, values):
        if values.get("folder") is None:
            return values["name"].split(".")[-1] if "." in values["name"] else None
        return None

    @validator("last_modified_by_name", pre=True, always=True)
    def set_last_modified_by_name(cls, v, values):
        last_modified_by = values.get("last_modified_by")
        if (
            last_modified_by
            and "user" in last_modified_by
            and "displayName" in last_modified_by["user"]
        ):
            return last_modified_by["user"]["displayName"]
        return None

    @validator("last_modified_by_email", pre=True, always=True)
    def set_last_modified_by_email(cls, v, values):
        last_modified_by = values.get("last_modified_by")
        if (last_modified_by and "user" in last_modified_by and "email" in last_modified_by["user"]):
            return last_modified_by["user"]["email"]
        return None

    def dict(self, *args, **kwargs):
        kwargs.setdefault("exclude", {"folder", "last_modified_by"})
        return super().dict(*args, **kwargs)


class CreateDir(BaseModel):
    """
    Define the data structure for create_dir() responses.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the created SharePoint directory.
    name : str, optional
        Specify the name of the created SharePoint directory.
    web_url : str, optional
        Specify the web URL of the created SharePoint directory.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint directory.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")


class RenameFolder(BaseModel):
    """
    Represent the data structure for rename_folder() responses.

    Define fields for SharePoint folder renaming metadata, including identifiers, names, URLs, and timestamps for
    creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the renamed SharePoint folder.
    name : str
        Specify the new name of the SharePoint folder.
    web_url : str, optional
        Specify the web URL of the renamed SharePoint folder.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint folder.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint folder.
    """

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetFileInfo(BaseModel):
    """
    Represent the data structure for get_file_info() responses.

    Define fields for SharePoint file metadata, including identifiers, names, URLs, sizes, and timestamps for creation
    and modification. Extract the last modified user's email if available.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint file.
    name : str
        Specify the name of the SharePoint file.
    web_url : str, optional
        Specify the web URL of the SharePoint file.
    size : int, optional
        Specify the size of the SharePoint file in bytes.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint file.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint file.
    last_modified_by : dict, optional
        Specify metadata about the user who last modified the file.
    last_modified_by_email : str, optional
        Specify the email of the user who last modified the file.
    """

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
        """Get last modified email."""
        # Handle cases where lastModifiedBy or user.email is missing
        last_modified_by = values.get("last_modified_by")
        if (last_modified_by and "user" in last_modified_by and "email" in last_modified_by["user"]):
            return last_modified_by["user"]["email"]
        return None

    # Exclude last_modified_by from dict() method
    def dict(self, *args, **kwargs):
        """Override dict() to exclude last_modified_by from output."""
        # Override dict() to exclude last_modified_by from output
        kwargs.setdefault("exclude", {"last_modified_by"})
        return super().dict(*args, **kwargs)


class MoveFile(BaseModel):
    """
    Represent the data structure for move_file() responses.

    Define fields for SharePoint file move metadata, including identifiers, names, URLs, sizes, and timestamps for
    creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the moved SharePoint file.
    name : str
        Specify the name of the moved SharePoint file.
    web_url : str, optional
        Specify the web URL of the moved SharePoint file.
    size : int, optional
        Specify the size of the moved SharePoint file in bytes.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint file.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint file.
    """

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class RenameFile(BaseModel):
    """
    Represent the data structure for rename_file() responses.

    Define fields for SharePoint file renaming metadata, including identifiers, names, URLs, sizes, and timestamps for
    creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the renamed SharePoint file.
    name : str
        Specify the new name of the SharePoint file.
    web_url : str, optional
        Specify the web URL of the renamed SharePoint file.
    size : int, optional
        Specify the size of the renamed SharePoint file in bytes.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint file.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint file.
    """

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class UploadFile(BaseModel):
    """
    Define the data structure for upload_file() responses.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the uploaded file.
    name : str, optional
        Specify the name of the uploaded file.
    size : int, optional
        Specify the size of the uploaded file in bytes.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    size: int = Field(None, alias="size")


class ListLists(BaseModel):
    """
    Represent the data structure for list_lists() responses.

    Define fields for SharePoint list metadata, including identifiers, names, descriptions, URLs, and timestamps for
    creation and modification.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint list.
    name : str, optional
        Specify the name of the SharePoint list.
    display_name : str, optional
        Specify the display name of the SharePoint list.
    description : str, optional
        Specify the description of the SharePoint list.
    web_url : str, optional
        Specify the web URL of the SharePoint list.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint list.
    last_modified_date_time : datetime, optional
        Specify the last modification date and time of the SharePoint list.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListListColumns(BaseModel):
    """
    Define the data structure for list_list_columns() responses.

    Specify fields for SharePoint list column metadata, including identifiers, names, display names, descriptions,
    column groups, and column properties such as uniqueness, visibility, indexing, read-only status, and requirement
    status.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint list column.
    name : str
        Specify the name of the SharePoint list column.
    display_name : str
        Specify the display name of the SharePoint list column.
    description : str
        Specify the description of the SharePoint list column.
    column_group : str
        Specify the group to which the column belongs.
    enforce_unique_values : bool
        Indicate whether the column enforces unique values.
    hidden : bool
        Indicate whether the column is hidden.
    indexed : bool
        Indicate whether the column is indexed.
    read_only : bool
        Indicate whether the column is read-only.
    required : bool
        Indicate whether the column is required.
    """

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


# Pydantic output data structure
class AddListItem(BaseModel):
    """
    Define the data structure for add_list_item() responses.

    Specify fields for SharePoint list item metadata, including identifiers, names, URLs, and creation timestamps.

    Parameters
    ----------
    id : str
        Specify the unique identifier of the SharePoint list item.
    name : str, optional
        Specify the name of the SharePoint list item.
    web_url : str, optional
        Specify the web URL of the SharePoint list item.
    created_date_time : datetime
        Specify the creation date and time of the SharePoint list item.
    """

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    # last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


# eof
