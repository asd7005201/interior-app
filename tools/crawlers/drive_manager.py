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


def clean_folder(folder_id: str | None = None, recursive: bool = True) -> dict:
    """
    폴더 내부 파일/하위폴더 전부 삭제 (폴더 자체는 유지).
    folder_id 미지정 시 DRIVE_ROOT_FOLDER_ID 사용.
    반환: {"deleted_files": n, "deleted_folders": n}
    """
    drive = _get_drive()
    target = folder_id or config.DRIVE_ROOT_FOLDER_ID
    stats = {"deleted_files": 0, "deleted_folders": 0}

    page_token = None
    while True:
        q = f"'{target}' in parents and trashed=false"
        result = drive.files().list(
            q=q, fields="nextPageToken, files(id,name,mimeType)", pageSize=100, pageToken=page_token
        ).execute()
        items = result.get("files", [])

        for item in items:
            try:
                if item["mimeType"] == "application/vnd.google-apps.folder":
                    if recursive:
                        sub = clean_folder(item["id"], recursive=True)
                        stats["deleted_files"] += sub["deleted_files"]
                        stats["deleted_folders"] += sub["deleted_folders"]
                    drive.files().delete(fileId=item["id"]).execute()
                    stats["deleted_folders"] += 1
                else:
                    drive.files().delete(fileId=item["id"]).execute()
                    stats["deleted_files"] += 1
            except Exception:
                stats.setdefault("skipped", 0)
                stats["skipped"] += 1

        page_token = result.get("nextPageToken")
        if not page_token:
            break

    # 캐시 초기화
    _folder_cache.clear()
    return stats
