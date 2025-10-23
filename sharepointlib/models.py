"""
Pydantic data structures for SharePoint entities.
"""

from datetime import datetime
from pydantic import BaseModel, Field, validator


class GetSiteInfo(BaseModel):
    """Pydantic data structure for get_site_info()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetHostNameInfo(BaseModel):
    """Pydantic data structure for get_hostname_info()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    site_collection: dict = Field(None, alias="siteCollection")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListDrives(BaseModel):
    """Pydantic data structure for list_drives()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    drive_type: str = Field(None, alias="driveType")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetDirInfo(BaseModel):
    """Pydantic data structure for get_dir_info()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListDir(BaseModel):
    """Pydantic data structure for list_dir()."""

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
        if (
            last_modified_by
            and "user" in last_modified_by
            and "email" in last_modified_by["user"]
        ):
            return last_modified_by["user"]["email"]
        return None

    def dict(self, *args, **kwargs):
        kwargs.setdefault("exclude", {"folder", "last_modified_by"})
        return super().dict(*args, **kwargs)


class CreateDir(BaseModel):
    """Pydantic data structure for create_dir()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")


class RenameFolder(BaseModel):
    """Pydantic data structure for rename_folder()."""

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class GetFileInfo(BaseModel):
    """Pydantic data structure for get_file_info()."""

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
        if (
            last_modified_by
            and "user" in last_modified_by
            and "email" in last_modified_by["user"]
        ):
            return last_modified_by["user"]["email"]
        return None

    # Exclude last_modified_by from dict() method
    def dict(self, *args, **kwargs):
        """Override dict() to exclude last_modified_by from output."""
        # Override dict() to exclude last_modified_by from output
        kwargs.setdefault("exclude", {"last_modified_by"})
        return super().dict(*args, **kwargs)


class MoveFile(BaseModel):
    """Pydantic data structure for move_file()."""

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class RenameFile(BaseModel):
    """Pydantic data structure for rename_file()."""

    id: str = Field(alias="id")
    name: str = Field(alias="name")
    web_url: str = Field(None, alias="webUrl")
    size: int = Field(None, alias="size")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class UploadFile(BaseModel):
    """Pydantic data structure for upload_file()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    size: int = Field(None, alias="size")


class ListLists(BaseModel):
    """Pydantic data structure for list_lists()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    display_name: str = Field(None, alias="displayName")
    description: str = Field(None, alias="description")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")


class ListListColumns(BaseModel):
    """Pydantic data structure for list_list_columns()."""

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
    """Pydantic data structure for add_list_item()."""

    id: str = Field(alias="id")
    name: str = Field(None, alias="name")
    web_url: str = Field(None, alias="webUrl")
    created_date_time: datetime = Field(alias="createdDateTime")
    # last_modified_date_time: datetime = Field(None, alias="lastModifiedDateTime")

# eom
