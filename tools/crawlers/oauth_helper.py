"""OAuth 사용자 인증 헬퍼.

첫 실행 시 브라우저에서 Google 로그인 → token.json 저장.
이후 실행부터는 자동으로 토큰 재사용.

사용법:
    python -m tools.crawlers.oauth_helper   # 토큰 발급/갱신
"""
import os
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from . import config

TOKEN_PATH = os.path.join(os.path.dirname(config.SERVICE_ACCOUNT_PATH), "token.json")
OAUTH_CREDS_PATH = os.path.join(os.path.dirname(config.SERVICE_ACCOUNT_PATH), "oauth_credentials.json")


def get_oauth_credentials() -> Credentials:
    """OAuth 사용자 인증 Credentials 반환. 없으면 브라우저 로그인 실행."""
    creds = None

    # 기존 토큰 로드
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, config.SCOPES)

    # 토큰 만료 시 갱신
    if creds and creds.expired and creds.refresh_token:
        try:
            creds.refresh(Request())
            _save_token(creds)
            return creds
        except Exception:
            creds = None

    # 토큰이 없거나 유효하지 않으면 브라우저 로그인
    if not creds or not creds.valid:
        if not os.path.exists(OAUTH_CREDS_PATH):
            raise FileNotFoundError(
                f"OAuth 클라이언트 JSON이 없습니다: {OAUTH_CREDS_PATH}\n"
                "GCP Console → API 및 서비스 → 사용자 인증 정보 → OAuth 클라이언트 ID → JSON 다운로드"
            )
        flow = InstalledAppFlow.from_client_secrets_file(OAUTH_CREDS_PATH, config.SCOPES)
        creds = flow.run_local_server(port=0)
        _save_token(creds)
        print(f"OAuth 토큰 저장 완료: {TOKEN_PATH}")

    return creds


def _save_token(creds: Credentials):
    """토큰을 파일에 저장."""
    with open(TOKEN_PATH, "w") as f:
        f.write(creds.to_json())


if __name__ == "__main__":
    c = get_oauth_credentials()
    print(f"인증 성공! (유효: {c.valid}, 만료: {c.expired})")
