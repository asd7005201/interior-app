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


def _collect_all_ids(drive, folder_id: str) -> list[dict]:
    """폴더 내 모든 항목(파일+하위폴더) ID를 재귀 수집. 깊은 것부터 반환."""
    all_items = []
    page_token = None
    while True:
        q = f"'{folder_id}' in parents and trashed=false"
        result = drive.files().list(
            q=q, fields="nextPageToken, files(id,mimeType)", pageSize=200, pageToken=page_token
        ).execute()
        items = result.get("files", [])
        for item in items:
            if item["mimeType"] == "application/vnd.google-apps.folder":
                # 하위 먼저 수집 (깊이 우선)
                all_items.extend(_collect_all_ids(drive, item["id"]))
            all_items.append(item)
        page_token = result.get("nextPageToken")
        if not page_token:
            break
    return all_items


def clean_folder(folder_id: str | None = None, recursive: bool = True) -> dict:
    """
    폴더 내부 파일/하위폴더 전부 삭제 (폴더 자체는 유지).
    배치 삭제로 속도 최적화. folder_id 미지정 시 DRIVE_ROOT_FOLDER_ID 사용.
    반환: {"deleted_files": n, "deleted_folders": n}
    """
    import time
    drive = _get_drive()
    target = folder_id or config.DRIVE_ROOT_FOLDER_ID
    stats = {"deleted_files": 0, "deleted_folders": 0, "skipped": 0}

    print("  항목 수집 중...", end=" ", flush=True)
    all_items = _collect_all_ids(drive, target)
    print(f"{len(all_items)}개 발견")

    # 배치 삭제 (50개씩 묶어서)
    batch_size = 50
    for i in range(0, len(all_items), batch_size):
        batch = all_items[i:i + batch_size]

        def _make_callback(item):
            def callback(req_id, response, exception):
                if exception:
                    stats["skipped"] += 1
                elif item["mimeType"] == "application/vnd.google-apps.folder":
                    stats["deleted_folders"] += 1
                else:
                    stats["deleted_files"] += 1
            return callback

        batch_req = drive.new_batch_http_request()
        for item in batch:
            batch_req.add(drive.files().delete(fileId=item["id"]), callback=_make_callback(item))
        batch_req.execute()

        done = min(i + batch_size, len(all_items))
        print(f"  삭제: {done}/{len(all_items)}", flush=True)
        time.sleep(1)  # Google 보호

    # 캐시 초기화
    _folder_cache.clear()
    return stats
