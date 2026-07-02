use std::path::PathBuf;

use pyo3::prelude::*;
use pyo3::types::{PyBytes, PyDict, PyModule};
use serde_json::Value;

use crate::{Response, ResponseContent, SharePoint};

#[pyclass(name = "Response")]
pub struct PyResponse {
    #[pyo3(get)]
    pub status_code: u16,
    content_json: Option<String>,
    content_bytes: Option<Vec<u8>>,
}

#[pymethods]
impl PyResponse {
    #[getter]
    fn content(&self, py: Python<'_>) -> PyResult<PyObject> {
        if let Some(raw_json) = &self.content_json {
            let json_module = PyModule::import(py, "json")?;
            let value = json_module.call_method1("loads", (raw_json,))?;
            return Ok(value.into());
        }

        if let Some(bytes) = &self.content_bytes {
            return Ok(PyBytes::new(py, bytes).into());
        }

        Ok(py.None())
    }
}

#[pyclass(name = "SharePoint")]
pub struct PySharePoint {
    inner: SharePoint,
}

#[pymethods]
impl PySharePoint {
    #[new]
    #[pyo3(signature = (client_id, tenant_id, client_secret, sp_domain, custom_logger=None))]
    fn new(
        client_id: String,
        tenant_id: String,
        client_secret: String,
        sp_domain: String,
        custom_logger: Option<PyObject>,
    ) -> PyResult<Self> {
        let _ = custom_logger;
        let inner = SharePoint::new(client_id, tenant_id, client_secret, sp_domain).map_err(into_pyerr)?;
        Ok(Self { inner })
    }

    fn renew_token(&mut self) -> PyResult<()> {
        self.inner.renew_token().map_err(into_pyerr)
    }

    #[pyo3(signature = (name, save_as=None))]
    fn get_site_info(&self, name: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .get_site_info(&name, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, save_as=None))]
    fn get_hostname_info(&self, site_id: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .get_hostname_info(&site_id, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, save_as=None))]
    fn list_drives(&self, site_id: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_drives(&site_id, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, path=None, save_as=None))]
    fn get_dir_info(
        &self,
        drive_id: String,
        path: Option<String>,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .get_dir_info(&drive_id, path.as_deref(), save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, path=None, alias=None, save_as=None))]
    fn list_dir(
        &self,
        drive_id: String,
        path: Option<String>,
        alias: Option<String>,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_dir(
                &drive_id,
                path.as_deref(),
                alias.as_deref(),
                save_as.as_ref().map(PathBuf::from),
            )
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, path, name, save_as=None))]
    fn create_dir(
        &self,
        drive_id: String,
        path: String,
        name: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .create_dir(&drive_id, &path, &name, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn delete_dir(&self, drive_id: String, path: String) -> PyResult<PyResponse> {
        let response = self.inner.delete_dir(&drive_id, &path).map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, path, new_name, save_as=None))]
    fn rename_folder(
        &self,
        drive_id: String,
        path: String,
        new_name: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .rename_folder(&drive_id, &path, &new_name, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, filename, save_as=None))]
    fn get_file_info(&self, drive_id: String, filename: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .get_file_info(&drive_id, &filename, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn check_out_file(&self, drive_id: String, filename: String) -> PyResult<PyResponse> {
        let response = self
            .inner
            .check_out_file(&drive_id, &filename)
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, filename, comment=None))]
    fn check_in_file(&self, drive_id: String, filename: String, comment: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .check_in_file(&drive_id, &filename, comment.as_deref())
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, filename, target_path, new_name=None))]
    fn copy_file(
        &self,
        drive_id: String,
        filename: String,
        target_path: String,
        new_name: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .copy_file(&drive_id, &filename, &target_path, new_name.as_deref())
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn get_download_url(&self, drive_id: String, filename: String) -> PyResult<Option<String>> {
        self.inner
            .get_download_url(&drive_id, &filename)
            .map_err(into_pyerr)
    }

    #[pyo3(signature = (
        source_drive_id,
        source_path,
        target_drive_id,
        target_path,
        new_name=None,
        chunk_size=20971520,
        timeout=3600,
        save_as=None
    ))]
    fn copy_file_stream(
        &self,
        source_drive_id: String,
        source_path: String,
        target_drive_id: String,
        target_path: String,
        new_name: Option<String>,
        chunk_size: usize,
        timeout: u64,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .copy_file_stream(
                &source_drive_id,
                &source_path,
                &target_drive_id,
                &target_path,
                new_name.as_deref(),
                chunk_size,
                timeout,
                save_as.as_ref().map(PathBuf::from),
            )
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, filename, target_path, new_name=None, save_as=None))]
    fn move_file(
        &self,
        drive_id: String,
        filename: String,
        target_path: String,
        new_name: Option<String>,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .move_file(
                &drive_id,
                &filename,
                &target_path,
                new_name.as_deref(),
                save_as.as_ref().map(PathBuf::from),
            )
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn delete_file(&self, drive_id: String, filename: String) -> PyResult<PyResponse> {
        let response = self.inner.delete_file(&drive_id, &filename).map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, filename, new_name, save_as=None))]
    fn rename_file(
        &self,
        drive_id: String,
        filename: String,
        new_name: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .rename_file(&drive_id, &filename, &new_name, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn download_file(&self, drive_id: String, remote_path: String, local_path: String) -> PyResult<PyResponse> {
        let response = self
            .inner
            .download_file(&drive_id, &remote_path, PathBuf::from(local_path))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn download_file_to_memory(&self, drive_id: String, remote_path: String) -> PyResult<PyResponse> {
        let response = self
            .inner
            .download_file_to_memory(&drive_id, &remote_path)
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn download_all_files(&self, drive_id: String, remote_path: String, local_path: String) -> PyResult<PyResponse> {
        let response = self
            .inner
            .download_all_files(&drive_id, &remote_path, PathBuf::from(local_path))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (drive_id, local_path, remote_path, save_as=None))]
    fn upload_file(
        &self,
        drive_id: String,
        local_path: String,
        remote_path: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .upload_file(
                &drive_id,
                PathBuf::from(local_path),
                &remote_path,
                save_as.as_ref().map(PathBuf::from),
            )
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, save_as=None))]
    fn list_lists(&self, site_id: String, save_as: Option<String>) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_lists(&site_id, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, list_id, save_as=None))]
    fn list_list_columns(
        &self,
        site_id: String,
        list_id: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_list_columns(&site_id, &list_id, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, list_id, fields, save_as=None))]
    fn list_list_items(
        &self,
        site_id: String,
        list_id: String,
        fields: String,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let response = self
            .inner
            .list_list_items(&site_id, &list_id, &fields, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    fn delete_list_item(&self, site_id: String, list_id: String, item_id: String) -> PyResult<PyResponse> {
        let response = self
            .inner
            .delete_list_item(&site_id, &list_id, &item_id)
            .map_err(into_pyerr)?;
        py_response_from(response)
    }

    #[pyo3(signature = (site_id, list_id, item, save_as=None))]
    fn add_list_item(
        &self,
        py: Python<'_>,
        site_id: String,
        list_id: String,
        item: &Bound<'_, PyAny>,
        save_as: Option<String>,
    ) -> PyResult<PyResponse> {
        let item_value = py_any_to_json_value(py, item)?;
        let response = self
            .inner
            .add_list_item(&site_id, &list_id, item_value, save_as.as_ref().map(PathBuf::from))
            .map_err(into_pyerr)?;
        py_response_from(response)
    }
}

#[pymodule]
fn sharepointlib(_py: Python<'_>, module: &Bound<'_, PyModule>) -> PyResult<()> {
    module.add_class::<PySharePoint>()?;
    module.add_class::<PyResponse>()?;
    module.add("__version__", env!("CARGO_PKG_VERSION"))?;
    Ok(())
}

fn py_response_from(response: Response) -> PyResult<PyResponse> {
    let (content_json, content_bytes) = match response.content {
        Some(ResponseContent::Json(value)) => (
            Some(serde_json::to_string(&value).map_err(into_pyerr)?),
            None,
        ),
        Some(ResponseContent::Bytes(bytes)) => (None, Some(bytes)),
        None => (None, None),
    };

    Ok(PyResponse {
        status_code: response.status_code,
        content_json,
        content_bytes,
    })
}

fn py_any_to_json_value(py: Python<'_>, value: &Bound<'_, PyAny>) -> PyResult<Value> {
    let json_module = PyModule::import(py, "json")?;
    let dumps = json_module.getattr("dumps")?;
    let raw = dumps.call1((value,))?;
    let raw_str: String = raw.extract()?;
    serde_json::from_str(&raw_str).map_err(into_pyerr)
}

fn into_pyerr<E>(error: E) -> PyErr
where
    E: std::fmt::Display,
{
    pyo3::exceptions::PyRuntimeError::new_err(error.to_string())
}

#[allow(dead_code)]
fn _type_markers(_a: Option<Value>, _b: Option<PyDict>) {}
