"""Drive 폴더 트리 출력"""
from tools.crawlers.drive_manager import _get_drive, _find_child_folder
from tools.crawlers import config


def show_tree(service, parent_id, depth=0):
    q = f"'{parent_id}' in parents and trashed=false"
    items = service.files().list(q=q, fields="files(id, name, mimeType)", pageSize=200).execute().get("files", [])

    folders = sorted([i for i in items if i["mimeType"] == "application/vnd.google-apps.folder"], key=lambda x: x["name"])
    files = [i for i in items if i["mimeType"] != "application/vnd.google-apps.folder"]

    for f in folders:
        sub_q = f"'{f['id']}' in parents and trashed=false"
        sub = service.files().list(q=sub_q, fields="files(id, mimeType)", pageSize=500).execute().get("files", [])
        sub_folders = [s for s in sub if s["mimeType"] == "application/vnd.google-apps.folder"]
        sub_files = [s for s in sub if s["mimeType"] != "application/vnd.google-apps.folder"]
        indent = "  " * depth
        if sub_folders:
            print(f"{indent}📁 {f['name']}/")
            show_tree(service, f["id"], depth + 1)
        else:
            print(f"{indent}📁 {f['name']}/ ({len(sub_files)} files)")

    if files and depth > 0:
        indent = "  " * depth
        print(f"{indent}📄 [{len(files)} files]")


def main():
    service = _get_drive()
    root_id = config.DRIVE_ROOT_FOLDER_ID
    print(f"Drive root: {root_id}")
    print("=" * 50)
    show_tree(service, root_id)
    print("=" * 50)


if __name__ == "__main__":
    main()
