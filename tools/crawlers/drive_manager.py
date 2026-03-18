"""Google Drive 폴더/이미지 관리 — OAuth 사용자 인증 사용"""
import io
import requests
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from . import config

_drive_service = None
_folder_cache: dict[str, str] = {}   # "벽지/신한벽지/실크" -> folder_id


def _get_drive():
    global _drive_service
    if _drive_service is None:
        from .oauth_helper import get_oauth_credentials
        creds = get_oauth_credentials()
        _drive_service = build("drive", "v3", credentials=creds)
    return _drive_service


def _find_child_folder(parent_id: str, name: str) -> str | None:
    """parent_id 아래에서 name 폴더를 찾아 ID 반환, 없으면 None."""
    drive = _get_drive()
    q = f"'{parent_id}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    result = drive.files().list(q=q, fields="files(id,name)", pageSize=1).execute()
    files = result.get("files", [])
    return files[0]["id"] if files else None


def _create_folder(parent_id: str, name: str) -> str:
    """parent_id 아래에 name 폴더 생성, ID 반환."""
    drive = _get_drive()
    metadata = {
        "name": name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [parent_id],
    }
    folder = drive.files().create(body=metadata, fields="id").execute()
    return folder["id"]


def ensure_folder_path(path_parts: list[str]) -> str:
    """
    Drive 루트 폴더 아래에 path_parts 경로의 폴더를 보장하고 최종 폴더 ID 반환.
    예: ensure_folder_path(["벽지", "신한벽지", "실크"])
    """
    cache_key = "/".join(path_parts)
    if cache_key in _folder_cache:
        return _folder_cache[cache_key]

    current_id = config.DRIVE_ROOT_FOLDER_ID
    for i, part in enumerate(path_parts):
        sub_key = "/".join(path_parts[: i + 1])
        if sub_key in _folder_cache:
            current_id = _folder_cache[sub_key]
            continue
        child_id = _find_child_folder(current_id, part)
        if child_id is None:
            child_id = _create_folder(current_id, part)
        _folder_cache[sub_key] = child_id
        current_id = child_id

    return current_id


def upload_image_from_url(image_url: str, folder_id: str, filename: str) -> dict:
    """
    이미지 URL에서 다운로드 -> Drive 폴더에 업로드.
    반환: {"file_id": ..., "file_name": ...}
    """
    resp = requests.get(image_url, headers=config.REQUEST_HEADERS, timeout=config.REQUEST_TIMEOUT)
    resp.raise_for_status()

    content_type = resp.headers.get("Content-Type", "image/jpeg")
    if ";" in content_type:
        content_type = content_type.split(";")[0].strip()

    drive = _get_drive()
    metadata = {"name": filename, "parents": [folder_id]}
    media = MediaIoBaseUpload(io.BytesIO(resp.content), mimetype=content_type, resumable=True)
    uploaded = drive.files().create(body=metadata, media_body=media, fields="id,name").execute()

    return {"file_id": uploaded["id"], "file_name": uploaded["name"]}
