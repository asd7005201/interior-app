"""Drive 폴더 구조 사전 생성.

최종 구조:
Materials/
  벽지/
    LX하우시스/ (디아망, 디아망포티스, 베스티, 테라피, 프리모, 휘앙세)
    개나리벽지/ (로하스, 아트북, 프리모, 트랜디, 스타일)
    신한벽지/ (파사드, 스케치, 아이리스, 에상스)
    현대벽지/ (큐피트, 큐브, 큐티에)
    FT벽지/ (이룸, 벨루체, 더뷰)
    DID벽지/ (더원, 디앤디, 나인, 컬러스, 작은방)
    코스모스벽지/ (앨리스, 소호, 모던)
    서울벽지/ (레노, 데이지, 플레인)
  시트지/
    단색/ 무늬목/ 대리석/ 하이그로시/ 디자인패턴/ 메탈/ 회벽벽돌/ 패브릭/
  페인트/
    벽지벽면용/ 방문가구용/ 타일욕실용/ 외부철재용/
  바닥재/
    장판/ (KCC, LX하우시스, 진양, 현대, 대진)
    데코타일/
    바닥시트지/
  타일/
    바닥타일/ 벽타일/ 데코모자이크/ 헥사곤/ 테라조/
  도기/
    양변기/ 세면기/ 욕조/ 수전/
  도어/
    실내도어/ 현관문/
  몰딩/
  조명/
  보수제/
"""
import time
from tools.crawlers.drive_manager import _get_drive
from tools.crawlers import config

_cache = {}


def ensure_folder(drive, parent_id, name):
    key = f"{parent_id}/{name}"
    if key in _cache:
        return _cache[key]

    q = f"'{parent_id}' in parents and name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    result = drive.files().list(q=q, fields="files(id)", pageSize=1).execute()
    files = result.get("files", [])

    if files:
        fid = files[0]["id"]
    else:
        meta = {"name": name, "mimeType": "application/vnd.google-apps.folder", "parents": [parent_id]}
        fid = drive.files().create(body=meta, fields="id").execute()["id"]
        print(f"  + {name}")
        time.sleep(0.5)

    _cache[key] = fid
    return fid


FOLDER_TREE = {
    "벽지": {
        "LX하우시스": ["디아망", "디아망포티스", "베스티", "테라피", "프리모", "휘앙세"],
        "개나리벽지": ["로하스", "아트북", "프리모", "트랜디", "스타일"],
        "신한벽지": ["파사드", "스케치", "아이리스"],
        "현대벽지": ["큐피트", "큐브", "큐티에"],
        "FT벽지": ["이룸", "벨루체", "더뷰"],
        "DID벽지": ["더원", "디앤디", "나인", "컬러스", "작은방"],
        "코스모스벽지": ["앨리스", "소호", "모던"],
        "서울벽지": ["레노", "데이지", "플레인"],
    },
    "시트지": {
        "단색": [],
        "무늬목": [],
        "대리석": [],
        "하이그로시": [],
        "디자인패턴": [],
        "메탈": [],
        "회벽벽돌": [],
        "패브릭": [],
    },
    "페인트": {
        "벽지벽면용": [],
        "방문가구용": [],
        "타일욕실용": [],
        "외부철재용": [],
    },
    "장판": {
        "KCC": [],
        "LX하우시스": [],
        "진양": [],
        "현대": [],
        "대진": [],
    },
    "데코타일": {},
    "바닥시트지": {},
    "타일": {
        "바닥타일": [],
        "벽타일": [],
        "데코모자이크": [],
    },
    "도기": {
        "대림바스": [],
        "아메리칸스탠다드": [],
    },
    "수전": {},
    "도어": {
        "실내도어": [],
        "현관문": [],
    },
    "몰딩": {},
    "조명": {},
    "보수제": {},
    "논슬립": {},
    "방수에폭시": {},
    "스테인": {},
    "프라이머": {},
}


def main():
    drive = _get_drive()
    root = config.DRIVE_ROOT_FOLDER_ID
    print("Drive 폴더 구조 생성 시작...")
    print(f"  root: {root}")

    for l1_name, l2_dict in FOLDER_TREE.items():
        l1_id = ensure_folder(drive, root, l1_name)
        if isinstance(l2_dict, dict):
            for l2_name, l3_list in l2_dict.items():
                l2_id = ensure_folder(drive, l1_id, l2_name)
                if isinstance(l3_list, list):
                    for l3_name in l3_list:
                        ensure_folder(drive, l2_id, l3_name)

    print("\n폴더 구조 생성 완료!")


if __name__ == "__main__":
    main()
