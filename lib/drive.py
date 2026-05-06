"""Google Drive I/O for the KAS workbook + processed PDFs.

Authenticates with a service account. The Excel file ID is configured in
secrets. Processed PDFs go into a folder owned/shared the same way; if it
doesn't exist yet, this module creates it on first use and shares it with the
workbook owner.
"""
from __future__ import annotations
import io
import json
import os
import time
from contextlib import contextmanager
from pathlib import Path

import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

SCOPES = ["https://www.googleapis.com/auth/drive"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PDF_MIME = "application/pdf"
FOLDER_MIME = "application/vnd.google-apps.folder"


@st.cache_resource
def _service():
    """Build a Drive service from secrets. Supports either a JSON path or
    inline JSON in secrets (as ``GCP_SERVICE_ACCOUNT``)."""
    secrets = st.secrets
    if "GCP_SERVICE_ACCOUNT" in secrets:
        info = secrets["GCP_SERVICE_ACCOUNT"]
        if isinstance(info, str):
            info = json.loads(info)
        creds = service_account.Credentials.from_service_account_info(dict(info), scopes=SCOPES)
    else:
        path = secrets["GCP_SERVICE_ACCOUNT_PATH"]
        if not os.path.isabs(path):
            path = str(Path(__file__).resolve().parent.parent / path)
        creds = service_account.Credentials.from_service_account_file(path, scopes=SCOPES)
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def file_metadata(file_id: str) -> dict:
    return _service().files().get(
        fileId=file_id,
        fields="id,name,size,modifiedTime,md5Checksum,parents",
    ).execute()


def download_xlsx(file_id: str) -> bytes:
    req = _service().files().get_media(fileId=file_id)
    buf = io.BytesIO()
    dl = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = dl.next_chunk()
    buf.seek(0)
    return buf.getvalue()


def upload_xlsx(file_id: str, content: bytes) -> dict:
    """Overwrite existing file (preserves the file ID and shares)."""
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=XLSX_MIME, resumable=False)
    return _service().files().update(
        fileId=file_id, media_body=media,
        fields="id,name,size,modifiedTime,md5Checksum",
    ).execute()


def find_or_create_folder(name: str, parent_id: str | None = None) -> str:
    svc = _service()
    q = (f"name='{name}' and mimeType='{FOLDER_MIME}' and trashed=false"
         + (f" and '{parent_id}' in parents" if parent_id else ""))
    res = svc.files().list(q=q, fields="files(id,name)").execute()
    files = res.get("files", [])
    if files:
        return files[0]["id"]
    body = {"name": name, "mimeType": FOLDER_MIME}
    if parent_id:
        body["parents"] = [parent_id]
    folder = svc.files().create(body=body, fields="id").execute()
    return folder["id"]


def upload_pdf(folder_id: str, filename: str, content: bytes) -> dict:
    media = MediaIoBaseUpload(io.BytesIO(content), mimetype=PDF_MIME, resumable=False)
    body = {"name": filename, "parents": [folder_id]}
    return _service().files().create(
        body=body, media_body=media,
        fields="id,name,webViewLink",
    ).execute()


def get_modified_time(file_id: str) -> str:
    return file_metadata(file_id)["modifiedTime"]


@contextmanager
def workbook_lock(file_id: str, max_wait_s: int = 30):
    """Best-effort optimistic-locking via modifiedTime check.

    Reads the modifiedTime before yielding. After the caller writes back, they
    should pass that timestamp into ``upload_xlsx_if_unchanged``. This context
    manager doesn't itself prevent concurrent edits — it just records the
    baseline so callers can detect collisions.
    """
    baseline = get_modified_time(file_id)
    yield baseline


def upload_xlsx_if_unchanged(file_id: str, content: bytes, expected_modified_time: str) -> dict:
    """Upload only if the remote file's modifiedTime matches what we expected.

    Raises ``ConcurrentEditError`` otherwise so the UI can prompt the user to
    reload and retry.
    """
    actual = get_modified_time(file_id)
    if actual != expected_modified_time:
        raise ConcurrentEditError(
            f"Workbook was modified externally (expected {expected_modified_time}, got {actual}). "
            "Reload to pick up changes and try again."
        )
    return upload_xlsx(file_id, content)


class ConcurrentEditError(RuntimeError):
    pass


def share_with(file_id: str, email: str, role: str = "reader") -> dict:
    return _service().permissions().create(
        fileId=file_id,
        body={"type": "user", "role": role, "emailAddress": email},
        sendNotificationEmail=False,
    ).execute()
