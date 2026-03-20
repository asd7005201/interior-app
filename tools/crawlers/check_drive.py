"""Drive 폴더 구조 확인 스크립트."""
import time
from .drive_manager import _get_drive
from . import config


def main():
    drive = _get_drive()
    ROOT = config.DRIVE_ROOT_FOLDER_ID

    def list_tree(parent_id, depth=0):
        time.sleep(1)
        q = f"'{parent_id}' in parents and trashed=false"
        result = drive.files().list(
            q=q, fields="files(id,name,mimeType)", pageSize=100
        ).execute()
        files = result.get("files", [])
        folders = [f for f in files if f["mimeType"] == "application/vnd.google-apps.folder"]
        images = [f for f in files if f["mimeType"] != "application/vnd.google-apps.folder"]

        indent = "  " * depth
        for folder in sorted(folders, key=lambda x: x["name"]):
            print(f"{indent}[D] {folder['name']} ({folder['id'][:12]})")
            if depth < 3:
                list_tree(folder["id"], depth + 1)
        if images:
            print(f"{indent}[F] images: {len(images)}")

    print(f"Drive Materials root: {ROOT}")
    print("-" * 50)
    list_tree(ROOT)


if __name__ == "__main__":
    main()
