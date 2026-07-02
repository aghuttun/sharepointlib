mod models;
#[cfg(feature = "python")]
mod python_bindings;

use std::fs;
use std::io::Read;
use std::path::Path;
use std::time::Duration;

use regex::Regex;
use reqwest::blocking::{Client, RequestBuilder, Response as HttpResponse};
use reqwest::header::{ACCEPT, AUTHORIZATION, CONTENT_TYPE, HeaderMap, HeaderValue, LOCATION};
use reqwest::{Method, StatusCode};
use serde::Deserialize;
use serde_json::{Map, Value, json};
use thiserror::Error;

use models::*;

#[derive(Debug, Clone)]
pub struct Configuration {
    pub api_domain: String,
    pub api_version: String,
    pub sp_domain: String,
    pub client_id: String,
    pub tenant_id: String,
    pub client_secret: String,
    pub token: Option<String>,
}

#[derive(Debug, Clone)]
pub enum ResponseContent {
    Json(Value),
    Bytes(Vec<u8>),
}

#[derive(Debug, Clone)]
pub struct Response {
    pub status_code: u16,
    pub content: Option<ResponseContent>,
}

#[derive(Debug)]
pub struct SharePoint {
    configuration: Configuration,
    client: Client,
    no_redirect_client: Client,
}

#[derive(Debug, Error)]
pub enum SharePointError {
    #[error("HTTP error: {0}")]
    Http(#[from] reqwest::Error),
    #[error("I/O error: {0}")]
    Io(#[from] std::io::Error),
    #[error("JSON error: {0}")]
    Json(#[from] serde_json::Error),
    #[error("missing bearer token; authenticate first")]
    MissingToken,
    #[error("invalid header value: {0}")]
    InvalidHeader(#[from] reqwest::header::InvalidHeaderValue),
    #[error("invalid regex pattern: {0}")]
    InvalidRegex(#[from] regex::Error),
    #[error("invalid API response: {0}")]
    InvalidResponse(String),
}

#[derive(Debug, Clone, Deserialize)]
struct AuthTokenResponse {
    access_token: String,
}

impl SharePoint {
    pub fn new(
        client_id: impl Into<String>,
        tenant_id: impl Into<String>,
        client_secret: impl Into<String>,
        sp_domain: impl Into<String>,
    ) -> Result<Self, SharePointError> {
        let client = Client::builder().build()?;
        let no_redirect_client = Client::builder()
            .redirect(reqwest::redirect::Policy::none())
            .build()?;

        let mut this = Self {
            configuration: Configuration {
                api_domain: "graph.microsoft.com".to_string(),
                api_version: "v1.0".to_string(),
                sp_domain: sp_domain.into(),
                client_id: client_id.into(),
                tenant_id: tenant_id.into(),
                client_secret: client_secret.into(),
                token: None,
            },
            client,
            no_redirect_client,
        };

        this.authenticate()?;
        Ok(this)
    }

    pub fn renew_token(&mut self) -> Result<(), SharePointError> {
        self.authenticate()
    }

    fn authenticate(&mut self) -> Result<(), SharePointError> {
        let url_auth = format!(
            "https://login.microsoftonline.com/{}/oauth2/v2.0/token",
            self.configuration.tenant_id
        );

        let response = self
            .client
            .post(url_auth)
            .header(CONTENT_TYPE, "application/x-www-form-urlencoded")
            .form(&[
                ("grant_type", "client_credentials"),
                ("client_id", self.configuration.client_id.as_str()),
                ("client_secret", self.configuration.client_secret.as_str()),
                ("scope", "https://graph.microsoft.com/.default"),
            ])
            .send()?;

        if response.status() == StatusCode::OK {
            let token: AuthTokenResponse = response.json()?;
            self.configuration.token = Some(token.access_token);
        }

        Ok(())
    }

    fn token(&self) -> Result<&str, SharePointError> {
        self.configuration
            .token
            .as_deref()
            .ok_or(SharePointError::MissingToken)
    }

    fn headers_json(&self) -> Result<HeaderMap, SharePointError> {
        let mut headers = HeaderMap::new();
        headers.insert(CONTENT_TYPE, HeaderValue::from_static("application/json"));
        headers.insert(ACCEPT, HeaderValue::from_static("application/json"));
        headers.insert(
            AUTHORIZATION,
            HeaderValue::from_str(&format!("Bearer {}", self.token()?))?,
        );
        Ok(headers)
    }

    fn headers_octet_stream(&self) -> Result<HeaderMap, SharePointError> {
        let mut headers = HeaderMap::new();
        headers.insert(
            AUTHORIZATION,
            HeaderValue::from_str(&format!("Bearer {}", self.token()?))?,
        );
        headers.insert(
            CONTENT_TYPE,
            HeaderValue::from_static("application/octet-stream"),
        );
        Ok(headers)
    }

    fn export_to_json(
        &self,
        content: &[u8],
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<(), SharePointError> {
        if let Some(path) = save_as {
            fs::write(path, content)?;
        }
        Ok(())
    }

    fn request(
        &self,
        method: Method,
        path: &str,
        query: Option<&[(String, String)]>,
        body: Option<&Value>,
    ) -> Result<HttpResponse, SharePointError> {
        let mut req: RequestBuilder = self
            .client
            .request(
                method,
                format!(
                    "https://{}/{}/{}",
                    self.configuration.api_domain,
                    self.configuration.api_version,
                    path.trim_start_matches('/')
                ),
            )
            .headers(self.headers_json()?);

        if let Some(params) = query {
            req = req.query(params);
        }

        if let Some(payload) = body {
            req = req.json(payload);
        }

        Ok(req.send()?)
    }

    fn parse_scalar_json(
        &self,
        response: HttpResponse,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Option<ResponseContent>, SharePointError> {
        let bytes = response.bytes()?;
        self.export_to_json(&bytes, save_as)?;
        let mut value: Value = serde_json::from_slice(&bytes)?;
        enrich_file_info_like_value(&mut value);
        Ok(Some(ResponseContent::Json(value)))
    }

    fn parse_list_json(
        &self,
        response: HttpResponse,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Option<ResponseContent>, SharePointError> {
        let bytes = response.bytes()?;
        self.export_to_json(&bytes, save_as)?;
        let raw: Value = serde_json::from_slice(&bytes)?;
        let list = raw
            .get("value")
            .and_then(Value::as_array)
            .cloned()
            .unwrap_or_default();
        let mut list_value = Value::Array(list);
        enrich_file_info_like_value(&mut list_value);
        Ok(Some(ResponseContent::Json(list_value)))
    }

    pub fn get_site_info(
        &self,
        name: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let site_name = urlencoding::encode(name);
        let endpoint = format!("sites/{}:/sites/{}", self.configuration.sp_domain, site_name);
        let query = vec![("$select".to_string(), GET_SITE_INFO_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn get_hostname_info(
        &self,
        site_id: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}");
        let query = vec![("$select".to_string(), GET_HOSTNAME_INFO_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn list_drives(
        &self,
        site_id: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/drives");
        let query = vec![("$select".to_string(), LIST_DRIVES_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_list_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn get_dir_info(
        &self,
        drive_id: &str,
        path: Option<&str>,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let path_quote = if let Some(value) = path {
            format!("/{}", urlencoding::encode(value))
        } else {
            "///".to_string()
        };

        let endpoint = format!("drives/{drive_id}/root:{path_quote}");
        let query = vec![("$select".to_string(), GET_DIR_INFO_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn list_dir(
        &self,
        drive_id: &str,
        path: Option<&str>,
        alias: Option<&str>,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let path_quote = if let Some(value) = path {
            urlencoding::encode(value).to_string()
        } else {
            "/".to_string()
        };

        let endpoint = format!("drives/{drive_id}/items/root:/{path_quote}:/children");
        let query = vec![("$select".to_string(), LIST_DIR_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            let bytes = response.bytes()?;
            self.export_to_json(&bytes, save_as)?;
            let raw: Value = serde_json::from_slice(&bytes)?;
            let mut list = raw
                .get("value")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default();

            let alias_regex = match alias {
                Some(pattern) => Some(Regex::new(pattern)?),
                None => None,
            };

            for item in &mut list {
                if let Some(obj) = item.as_object_mut() {
                    let name = obj
                        .get("name")
                        .and_then(Value::as_str)
                        .unwrap_or_default()
                        .to_string();

                    let extension = if obj.get("folder").is_none() {
                        name.rsplit_once('.').map(|(_, ext)| ext.to_string())
                    } else {
                        None
                    };

                    let last_modified_name = obj
                        .get("lastModifiedBy")
                        .and_then(Value::as_object)
                        .and_then(|it| it.get("user"))
                        .and_then(Value::as_object)
                        .and_then(|user| user.get("displayName"))
                        .cloned();

                    let last_modified_email = obj
                        .get("lastModifiedBy")
                        .and_then(Value::as_object)
                        .and_then(|it| it.get("user"))
                        .and_then(Value::as_object)
                        .and_then(|user| user.get("email"))
                        .cloned();

                    let path_value = path.unwrap_or("/").to_string();
                    let alias_value = match &alias_regex {
                        Some(regex) => regex.replace_all(&name, "").to_string(),
                        None => name.clone(),
                    };

                    obj.insert("path".to_string(), Value::String(path_value));
                    obj.insert("alias".to_string(), Value::String(alias_value));

                    if let Some(ext) = extension {
                        obj.insert("extension".to_string(), Value::String(ext));
                    } else {
                        obj.insert("extension".to_string(), Value::Null);
                    }

                    if let Some(display_name) = last_modified_name {
                        obj.insert("last_modified_by_name".to_string(), display_name);
                    }
                    if let Some(email) = last_modified_email {
                        obj.insert("last_modified_by_email".to_string(), email);
                    }
                }
            }

            Some(ResponseContent::Json(Value::Array(list)))
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn create_dir(
        &self,
        drive_id: &str,
        path: &str,
        name: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let path_quote = urlencoding::encode(path);
        let endpoint = format!("drives/{drive_id}/root:/{path_quote}:/children");

        let query = vec![("$select".to_string(), CREATE_DIR_SELECT.join(","))];
        let body = json!({
            "name": name,
            "folder": {},
            "@microsoft.graph.conflictBehavior": "replace"
        });

        let response = self.request(Method::POST, &endpoint, Some(&query), Some(&body))?;
        let status = response.status().as_u16();

        let content = if status == 200 || status == 201 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn delete_dir(&self, drive_id: &str, path: &str) -> Result<Response, SharePointError> {
        let path_quote = urlencoding::encode(path);
        let endpoint = format!("drives/{drive_id}/root:/{path_quote}");

        let response = self.request(Method::DELETE, &endpoint, None, None)?;

        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn rename_folder(
        &self,
        drive_id: &str,
        path: &str,
        new_name: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let path_quote = urlencoding::encode(path);
        let endpoint = format!("drives/{drive_id}/root:/{path_quote}");
        let query = vec![("$select".to_string(), RENAME_FOLDER_SELECT.join(","))];
        let body = json!({ "name": new_name });

        let response = self.request(Method::PATCH, &endpoint, Some(&query), Some(&body))?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn get_file_info(
        &self,
        drive_id: &str,
        filename: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let filename_quote = urlencoding::encode(filename);
        let endpoint = format!("drives/{drive_id}/root:/{filename_quote}");
        let query = vec![("$select".to_string(), GET_FILE_INFO_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 || status == 202 {
            let mut content = self.parse_scalar_json(response, save_as)?;
            if let Some(ResponseContent::Json(Value::Object(ref mut map))) = content {
                let path_value = filename
                    .rsplit_once('/')
                    .map(|(left, _)| left.to_string())
                    .unwrap_or_else(|| "/".to_string());
                map.insert("path".to_string(), Value::String(path_value));
            }
            content
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn check_out_file(&self, drive_id: &str, filename: &str) -> Result<Response, SharePointError> {
        let file_info = self.get_file_info(drive_id, filename, Option::<&Path>::None)?;
        if file_info.status_code != 200 {
            return Ok(Response {
                status_code: file_info.status_code,
                content: None,
            });
        }

        let file_id = extract_id(&file_info.content)?;
        let endpoint = format!("drives/{drive_id}/items/{file_id}/checkout");
        let response = self.request(Method::POST, &endpoint, None, None)?;

        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn check_in_file(
        &self,
        drive_id: &str,
        filename: &str,
        comment: Option<&str>,
    ) -> Result<Response, SharePointError> {
        let file_info = self.get_file_info(drive_id, filename, Option::<&Path>::None)?;
        if file_info.status_code != 200 {
            return Ok(Response {
                status_code: file_info.status_code,
                content: None,
            });
        }

        let file_id = extract_id(&file_info.content)?;
        let endpoint = format!("drives/{drive_id}/items/{file_id}/checkin");
        let body = if let Some(value) = comment {
            json!({ "comment": value })
        } else {
            json!({})
        };

        let response = self.request(Method::POST, &endpoint, None, Some(&body))?;

        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn copy_file(
        &self,
        drive_id: &str,
        filename: &str,
        target_path: &str,
        new_name: Option<&str>,
    ) -> Result<Response, SharePointError> {
        let filename_quote = urlencoding::encode(filename);
        let endpoint = format!("drives/{drive_id}/root:/{filename_quote}:/copy");

        let mut body = json!({
            "parentReference": {
                "driveId": drive_id,
                "driveType": "documentLibrary",
                "path": format!("/drives/{drive_id}/root:/{target_path}")
            }
        });

        if let Some(name) = new_name {
            body["name"] = Value::String(name.to_string());
        }

        let response = self.request(Method::POST, &endpoint, None, Some(&body))?;
        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn get_download_url(
        &self,
        drive_id: &str,
        filename: &str,
    ) -> Result<Option<String>, SharePointError> {
        let filename_quote = urlencoding::encode(filename);
        let endpoint = format!(
            "https://{}/{}/drives/{drive_id}/root:/{filename_quote}:/content",
            self.configuration.api_domain, self.configuration.api_version
        );

        let response = self
            .no_redirect_client
            .get(endpoint)
            .header(
                AUTHORIZATION,
                HeaderValue::from_str(&format!("Bearer {}", self.token()?))?,
            )
            .send()?;

        if response.status().is_redirection() {
            let location = response
                .headers()
                .get(LOCATION)
                .and_then(|h| h.to_str().ok())
                .map(ToOwned::to_owned);
            return Ok(location);
        }

        Ok(None)
    }

    pub fn copy_file_stream(
        &self,
        source_drive_id: &str,
        source_path: &str,
        target_drive_id: &str,
        target_path: &str,
        new_name: Option<&str>,
        chunk_size: usize,
        timeout: u64,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let download_url = match self.get_download_url(source_drive_id, source_path)? {
            Some(url) => url,
            None => {
                return Ok(Response {
                    status_code: 400,
                    content: Some(ResponseContent::Json(json!("downloadUrl not available"))),
                });
            }
        };

        let source_info = self.get_file_info(source_drive_id, source_path, Option::<&Path>::None)?;
        if source_info.status_code != 200 {
            return Ok(Response {
                status_code: source_info.status_code,
                content: None,
            });
        }

        let file_size = extract_size(&source_info.content)?;
        let source_file_name = source_path.rsplit('/').next().unwrap_or(source_path);
        let final_name = new_name.unwrap_or(source_file_name);

        let remote_target_path = format!("{}/{}", target_path.trim_end_matches('/'), final_name);
        let remote_target_path = urlencoding::encode(&remote_target_path).to_string();

        let endpoint = format!(
            "drives/{target_drive_id}/root:/{remote_target_path}:/createUploadSession"
        );

        let payload = json!({
            "item": {"@microsoft.graph.conflictBehavior": "replace"},
            "name": final_name
        });

        let session_response = self.request(Method::POST, &endpoint, None, Some(&payload))?;
        let session_status = session_response.status().as_u16();
        if session_status != 200 && session_status != 201 {
            let raw = session_response.text().unwrap_or_default();
            return Ok(Response {
                status_code: session_status,
                content: Some(ResponseContent::Json(json!(raw))),
            });
        }

        let session_payload: Value = session_response.json()?;
        let upload_url = session_payload
            .get("uploadUrl")
            .and_then(Value::as_str)
            .ok_or_else(|| SharePointError::InvalidResponse("missing uploadUrl".to_string()))?
            .to_string();

        let mut source_response = self
            .client
            .get(download_url)
            .timeout(Duration::from_secs(timeout))
            .send()?;

        let mut uploaded: usize = 0;
        let mut buffer = vec![0_u8; chunk_size.max(320 * 1024)];

        loop {
            let bytes_read = source_response.read(&mut buffer)?;
            if bytes_read == 0 {
                break;
            }

            let start_byte = uploaded;
            let end_byte = uploaded + bytes_read - 1;
            let chunk = buffer[..bytes_read].to_vec();

            let upload_response = self
                .client
                .put(&upload_url)
                .header("Content-Length", bytes_read.to_string())
                .header(
                    "Content-Range",
                    format!("bytes {start_byte}-{end_byte}/{file_size}"),
                )
                .timeout(Duration::from_secs(timeout))
                .body(chunk)
                .send()?;

            let upload_status = upload_response.status().as_u16();

            if upload_status == 200 || upload_status == 201 {
                let bytes = upload_response.bytes()?;
                self.export_to_json(&bytes, save_as)?;
                let value: Value = serde_json::from_slice(&bytes)?;
                return Ok(Response {
                    status_code: upload_status,
                    content: Some(ResponseContent::Json(value)),
                });
            }

            if upload_status != 202 {
                let text = upload_response.text().unwrap_or_default();
                return Ok(Response {
                    status_code: upload_status,
                    content: Some(ResponseContent::Json(json!(text))),
                });
            }

            uploaded += bytes_read;
        }

        Ok(Response {
            status_code: 500,
            content: Some(ResponseContent::Json(json!("Copy interrupted unexpectedly"))),
        })
    }

    pub fn move_file(
        &self,
        drive_id: &str,
        filename: &str,
        target_path: &str,
        new_name: Option<&str>,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let file_info = self.get_file_info(drive_id, filename, Option::<&Path>::None)?;
        if file_info.status_code != 200 {
            return Ok(Response {
                status_code: file_info.status_code,
                content: None,
            });
        }

        let dir_info = self.get_dir_info(drive_id, Some(target_path), Option::<&Path>::None)?;
        if dir_info.status_code != 200 {
            return Ok(Response {
                status_code: dir_info.status_code,
                content: None,
            });
        }

        let file_id = extract_id(&file_info.content)?;
        let folder_id = extract_id(&dir_info.content)?;

        let endpoint = format!("drives/{drive_id}/items/{file_id}");
        let query = vec![("$select".to_string(), MOVE_FILE_SELECT.join(","))];

        let mut body = json!({ "parentReference": {"id": folder_id} });
        if let Some(name) = new_name {
            body["name"] = Value::String(name.to_string());
        }

        let response = self.request(Method::PATCH, &endpoint, Some(&query), Some(&body))?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn delete_file(&self, drive_id: &str, filename: &str) -> Result<Response, SharePointError> {
        let filename_quote = urlencoding::encode(filename);
        let endpoint = format!("drives/{drive_id}/root:/{filename_quote}");

        let response = self.request(Method::DELETE, &endpoint, None, None)?;
        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn rename_file(
        &self,
        drive_id: &str,
        filename: &str,
        new_name: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let filename_quote = urlencoding::encode(filename);
        let endpoint = format!("drives/{drive_id}/root:/{filename_quote}");
        let query = vec![("$select".to_string(), RENAME_FILE_SELECT.join(","))];
        let body = json!({ "name": new_name });

        let response = self.request(Method::PATCH, &endpoint, Some(&query), Some(&body))?;
        let status = response.status().as_u16();
        let content = if status == 200 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn download_file(
        &self,
        drive_id: &str,
        remote_path: &str,
        local_path: impl AsRef<Path>,
    ) -> Result<Response, SharePointError> {
        let remote_quote = urlencoding::encode(remote_path);
        let url = format!(
            "https://{}/{}/drives/{drive_id}/root:/{remote_quote}:/content",
            self.configuration.api_domain, self.configuration.api_version
        );

        let mut response = self
            .client
            .get(url)
            .headers(self.headers_octet_stream()?)
            .send()?;

        let status = response.status().as_u16();
        if status == 200 {
            let mut file = fs::File::create(local_path)?;
            std::io::copy(&mut response, &mut file)?;
        }

        Ok(Response {
            status_code: status,
            content: None,
        })
    }

    pub fn download_file_to_memory(
        &self,
        drive_id: &str,
        remote_path: &str,
    ) -> Result<Response, SharePointError> {
        let remote_quote = urlencoding::encode(remote_path);
        let url = format!(
            "https://{}/{}/drives/{drive_id}/root:/{remote_quote}:/content",
            self.configuration.api_domain, self.configuration.api_version
        );

        let response = self
            .client
            .get(url)
            .header(
                AUTHORIZATION,
                HeaderValue::from_str(&format!("Bearer {}", self.token()?))?,
            )
            .send()?;

        let status = response.status().as_u16();
        let content = if status == 200 {
            Some(ResponseContent::Bytes(response.bytes()?.to_vec()))
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn download_all_files(
        &self,
        drive_id: &str,
        remote_path: &str,
        local_path: impl AsRef<Path>,
    ) -> Result<Response, SharePointError> {
        let listed = self.list_dir(drive_id, Some(remote_path), None, Option::<&Path>::None)?;
        if listed.status_code != 200 {
            return Ok(Response {
                status_code: listed.status_code,
                content: None,
            });
        }

        let mut report = Vec::<Value>::new();
        let list = match listed.content {
            Some(ResponseContent::Json(Value::Array(items))) => items,
            _ => Vec::new(),
        };

        for item in list {
            let name = item
                .get("name")
                .and_then(Value::as_str)
                .unwrap_or_default()
                .to_string();

            let extension = item.get("extension").and_then(Value::as_str);
            if extension.is_none() {
                continue;
            }

            let local_target = local_path.as_ref().join(&name);
            let remote_target = format!("{remote_path}/{name}");
            let sub = self.download_file(drive_id, &remote_target, local_target)?;
            let status = if sub.status_code == 200 { "pass" } else { "fail" };

            let mut entry = item.as_object().cloned().unwrap_or_default();
            entry.insert("status".to_string(), Value::String(status.to_string()));
            report.push(Value::Object(entry));
        }

        Ok(Response {
            status_code: 200,
            content: Some(ResponseContent::Json(Value::Array(report))),
        })
    }

    pub fn upload_file(
        &self,
        drive_id: &str,
        local_path: impl AsRef<Path>,
        remote_path: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let remote_quote = urlencoding::encode(remote_path);
        let url = format!(
            "https://{}/{}/drives/{drive_id}/root:/{remote_quote}:/content",
            self.configuration.api_domain, self.configuration.api_version
        );

        let query = vec![("$select".to_string(), UPLOAD_FILE_SELECT.join(","))];
        let data = fs::read(local_path)?;

        let response = self
            .client
            .put(url)
            .headers(self.headers_octet_stream()?)
            .query(&query)
            .body(data)
            .send()?;

        let status = response.status().as_u16();
        let content = if status == 200 || status == 201 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn list_lists(
        &self,
        site_id: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/lists");
        let query = vec![("$select".to_string(), LIST_LISTS_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();
        let content = if status == 200 {
            self.parse_list_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn list_list_columns(
        &self,
        site_id: &str,
        list_id: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/lists/{list_id}/columns");
        let query = vec![("$select".to_string(), LIST_LIST_COLUMNS_SELECT.join(","))];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();
        let content = if status == 200 {
            self.parse_list_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn list_list_items(
        &self,
        site_id: &str,
        list_id: &str,
        fields: &str,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/lists/{list_id}/items");
        let query = vec![
            ("select".to_string(), fields.to_string()),
            ("expand".to_string(), "fields".to_string()),
        ];

        let response = self.request(Method::GET, &endpoint, Some(&query), None)?;
        let status = response.status().as_u16();

        let content = if status == 200 {
            let bytes = response.bytes()?;
            self.export_to_json(&bytes, save_as)?;
            let raw: Value = serde_json::from_slice(&bytes)?;
            let fields_only = raw
                .get("value")
                .and_then(Value::as_array)
                .cloned()
                .unwrap_or_default()
                .into_iter()
                .map(|item| item.get("fields").cloned().unwrap_or(Value::Null))
                .collect::<Vec<Value>>();

            Some(ResponseContent::Json(Value::Array(fields_only)))
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }

    pub fn delete_list_item(
        &self,
        site_id: &str,
        list_id: &str,
        item_id: &str,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/lists/{list_id}/items/{item_id}");
        let response = self.request(Method::DELETE, &endpoint, None, None)?;

        Ok(Response {
            status_code: response.status().as_u16(),
            content: None,
        })
    }

    pub fn add_list_item(
        &self,
        site_id: &str,
        list_id: &str,
        item: Value,
        save_as: Option<impl AsRef<Path>>,
    ) -> Result<Response, SharePointError> {
        let endpoint = format!("sites/{site_id}/lists/{list_id}/items");
        let query = vec![("$select".to_string(), ADD_LIST_ITEM_SELECT.join(","))];
        let body = json!({ "fields": item });

        let response = self.request(Method::POST, &endpoint, Some(&query), Some(&body))?;
        let status = response.status().as_u16();
        let content = if status == 201 {
            self.parse_scalar_json(response, save_as)?
        } else {
            None
        };

        Ok(Response {
            status_code: status,
            content,
        })
    }
}

fn extract_id(content: &Option<ResponseContent>) -> Result<String, SharePointError> {
    match content {
        Some(ResponseContent::Json(Value::Object(map))) => map
            .get("id")
            .and_then(Value::as_str)
            .map(ToOwned::to_owned)
            .ok_or_else(|| SharePointError::InvalidResponse("missing id".to_string())),
        _ => Err(SharePointError::InvalidResponse(
            "missing JSON content with id".to_string(),
        )),
    }
}

fn extract_size(content: &Option<ResponseContent>) -> Result<usize, SharePointError> {
    match content {
        Some(ResponseContent::Json(Value::Object(map))) => {
            let size_u64 = map
                .get("size")
                .and_then(Value::as_u64)
                .ok_or_else(|| SharePointError::InvalidResponse("missing size".to_string()))?;
            usize::try_from(size_u64)
                .map_err(|_| SharePointError::InvalidResponse("size conversion overflow".to_string()))
        }
        _ => Err(SharePointError::InvalidResponse(
            "missing JSON content with size".to_string(),
        )),
    }
}

fn enrich_file_info_like_value(value: &mut Value) {
    match value {
        Value::Object(map) => enrich_object(map),
        Value::Array(items) => {
            for item in items {
                if let Value::Object(map) = item {
                    enrich_object(map);
                }
            }
        }
        _ => {}
    }
}

fn enrich_object(map: &mut Map<String, Value>) {
    if let Some(name) = map.get("name").and_then(Value::as_str)
        && !map.contains_key("extension")
    {
        let extension = name.rsplit_once('.').map(|(_, ext)| ext.to_string());
        map.insert(
            "extension".to_string(),
            extension.map(Value::String).unwrap_or(Value::Null),
        );
    }

    let last_modified_name = map
        .get("lastModifiedBy")
        .and_then(Value::as_object)
        .and_then(|v| v.get("user"))
        .and_then(Value::as_object)
        .and_then(|user| user.get("displayName"))
        .cloned();

    let last_modified_email = map
        .get("lastModifiedBy")
        .and_then(Value::as_object)
        .and_then(|v| v.get("user"))
        .and_then(Value::as_object)
        .and_then(|user| user.get("email"))
        .cloned();

    if let Some(display_name) = last_modified_name {
        map.entry("last_modified_by_name".to_string())
            .or_insert(display_name);
    }

    if let Some(email) = last_modified_email {
        map.entry("last_modified_by_email".to_string())
            .or_insert(email);
    }
}
