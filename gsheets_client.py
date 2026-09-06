"""
Googleスプレッドシートへの最小限のAPIクライアント。

gspread/google-authではなく pycryptodome + requests だけで
サービスアカウントのJWT認証とSheets API v4呼び出しを行う。

背景: このPC環境では google-auth が依存する `cryptography` パッケージの
ネイティブDLL(_rust.pyd)が読み込めない(Git for Windowsのlibcrypto DLLとの
既知の競合)。pycryptodomeは同じ問題を回避できるため、依存を最小限にした
自前のクライアントで代替する。
"""

import base64
import json
import time

import requests
from Crypto.Hash import SHA256
from Crypto.PublicKey import RSA
from Crypto.Signature import pkcs1_15

TOKEN_URL = "https://oauth2.googleapis.com/token"
SCOPES = "https://www.googleapis.com/auth/spreadsheets"
API_BASE = "https://sheets.googleapis.com/v4/spreadsheets"


def _b64url(data: bytes) -> str:
    return base64.urlsafe_b64encode(data).rstrip(b"=").decode("ascii")


class SheetsClient:
    def __init__(self, service_account_info: dict, spreadsheet_id: str):
        self._info = service_account_info
        self.spreadsheet_id = spreadsheet_id
        self._token = None
        self._token_expiry = 0

    def _access_token(self) -> str:
        now = int(time.time())
        if self._token and now < self._token_expiry - 60:
            return self._token

        header = {"alg": "RS256", "typ": "JWT"}
        claims = {
            "iss": self._info["client_email"],
            "scope": SCOPES,
            "aud": TOKEN_URL,
            "iat": now,
            "exp": now + 3600,
        }
        signing_input = _b64url(json.dumps(header).encode()) + "." + _b64url(json.dumps(claims).encode())
        key = RSA.import_key(self._info["private_key"])
        digest = SHA256.new(signing_input.encode("ascii"))
        signature = pkcs1_15.new(key).sign(digest)
        assertion = signing_input + "." + _b64url(signature)

        resp = requests.post(
            TOKEN_URL,
            data={
                "grant_type": "urn:ietf:params:oauth:grant-type:jwt-bearer",
                "assertion": assertion,
            },
            timeout=15,
        )
        resp.raise_for_status()
        payload = resp.json()
        self._token = payload["access_token"]
        self._token_expiry = now + payload.get("expires_in", 3600)
        return self._token

    def _headers(self):
        return {"Authorization": f"Bearer {self._access_token()}"}

    def _spreadsheet_meta(self):
        resp = requests.get(f"{API_BASE}/{self.spreadsheet_id}", headers=self._headers(), timeout=15)
        resp.raise_for_status()
        return resp.json()

    def ensure_sheet(self, title: str, headers: list):
        sheets = self._spreadsheet_meta().get("sheets", [])
        if any(s["properties"]["title"] == title for s in sheets):
            return
        resp = requests.post(
            f"{API_BASE}/{self.spreadsheet_id}:batchUpdate",
            headers=self._headers(),
            json={"requests": [{"addSheet": {"properties": {"title": title}}}]},
            timeout=15,
        )
        resp.raise_for_status()
        self.append_row(title, headers)

    def sheet_id_for(self, title: str) -> int:
        for s in self._spreadsheet_meta().get("sheets", []):
            if s["properties"]["title"] == title:
                return s["properties"]["sheetId"]
        raise ValueError(f"sheet not found: {title}")

    def get_all_values(self, title: str) -> list:
        resp = requests.get(
            f"{API_BASE}/{self.spreadsheet_id}/values/{title}",
            headers=self._headers(),
            timeout=15,
        )
        resp.raise_for_status()
        return resp.json().get("values", [])

    def append_row(self, title: str, row: list):
        self.append_rows(title, [row])

    def append_rows(self, title: str, rows: list):
        resp = requests.post(
            f"{API_BASE}/{self.spreadsheet_id}/values/{title}:append",
            headers=self._headers(),
            params={"valueInputOption": "USER_ENTERED", "insertDataOption": "INSERT_ROWS"},
            json={"values": rows},
            timeout=30,
        )
        resp.raise_for_status()

    def delete_row(self, title: str, row_index: int):
        """row_index: get_all_values() の戻り値でのインデックス(0=ヘッダー行)。"""
        sheet_id = self.sheet_id_for(title)
        resp = requests.post(
            f"{API_BASE}/{self.spreadsheet_id}:batchUpdate",
            headers=self._headers(),
            json={
                "requests": [{
                    "deleteDimension": {
                        "range": {
                            "sheetId": sheet_id,
                            "dimension": "ROWS",
                            "startIndex": row_index,
                            "endIndex": row_index + 1,
                        }
                    }
                }]
            },
            timeout=15,
        )
        resp.raise_for_status()
