pub(crate) const GET_SITE_INFO_SELECT: &[&str] = &[
    "id",
    "name",
    "displayName",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const GET_HOSTNAME_INFO_SELECT: &[&str] = &[
    "id",
    "name",
    "displayName",
    "description",
    "webUrl",
    "siteCollection",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const LIST_DRIVES_SELECT: &[&str] = &[
    "id",
    "name",
    "description",
    "webUrl",
    "driveType",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const GET_DIR_INFO_SELECT: &[&str] = &[
    "id",
    "name",
    "webUrl",
    "size",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const LIST_DIR_SELECT: &[&str] = &[
    "id",
    "name",
    "size",
    "webUrl",
    "folder",
    "createdDateTime",
    "lastModifiedDateTime",
    "lastModifiedBy",
];

pub(crate) const CREATE_DIR_SELECT: &[&str] = &["id", "name", "webUrl", "createdDateTime"];

pub(crate) const RENAME_FOLDER_SELECT: &[&str] = &[
    "id",
    "name",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const GET_FILE_INFO_SELECT: &[&str] = &[
    "id",
    "name",
    "size",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
    "lastModifiedBy",
    "@microsoft.graph.downloadUrl",
];

pub(crate) const MOVE_FILE_SELECT: &[&str] = &[
    "id",
    "name",
    "size",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
    "lastModifiedBy",
];

pub(crate) const RENAME_FILE_SELECT: &[&str] = &[
    "id",
    "name",
    "size",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
    "lastModifiedBy",
];

pub(crate) const UPLOAD_FILE_SELECT: &[&str] = &[
    "id",
    "name",
    "size",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
    "lastModifiedBy",
];

pub(crate) const LIST_LISTS_SELECT: &[&str] = &[
    "id",
    "name",
    "displayName",
    "description",
    "webUrl",
    "createdDateTime",
    "lastModifiedDateTime",
];

pub(crate) const LIST_LIST_COLUMNS_SELECT: &[&str] = &[
    "id",
    "name",
    "displayName",
    "description",
    "columnGroup",
    "enforceUniqueValues",
    "hidden",
    "indexed",
    "readOnly",
    "required",
];

pub(crate) const ADD_LIST_ITEM_SELECT: &[&str] = &["id", "createdDateTime", "lastModifiedDateTime", "webUrl"];
