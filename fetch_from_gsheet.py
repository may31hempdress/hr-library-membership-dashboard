"""Google スプレッドシート(数字で見る人事図書館変遷)を xlsx としてエクスポートし、
extract.py に渡して public/data.json を更新する。

環境変数:
    GOOGLE_SERVICE_ACCOUNT_JSON — サービスアカウントの JSON 鍵。
        ファイルパス、または JSON 文字列そのもののどちらでも可。
    SPREADSHEET_ID — 対象スプレッドシートの ID。

使い方:
    python fetch_from_gsheet.py
"""
import io
import json
import os
import subprocess
import sys
from datetime import datetime, timedelta, timezone

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
JST = timezone(timedelta(hours=9))


def load_credentials():
    raw = os.environ["GOOGLE_SERVICE_ACCOUNT_JSON"]
    if os.path.isfile(raw):
        return Credentials.from_service_account_file(raw, scopes=SCOPES)
    return Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)


def download_xlsx(spreadsheet_id: str, creds) -> bytes:
    service = build("drive", "v3", credentials=creds)
    request = service.files().export_media(fileId=spreadsheet_id, mimeType=XLSX_MIME)
    buf = io.BytesIO()
    downloader = MediaIoBaseDownload(buf, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


def main():
    spreadsheet_id = os.environ["SPREADSHEET_ID"]
    creds = load_credentials()
    content = download_xlsx(spreadsheet_id, creds)

    today = datetime.now(JST).strftime("%y%m%d")
    base_dir = os.path.dirname(os.path.abspath(__file__))
    xlsx_path = os.path.join(base_dir, f"数字で見る人事図書館変遷_{today}時点.xlsx")
    with open(xlsx_path, "wb") as f:
        f.write(content)
    print(f"downloaded: {xlsx_path}")

    subprocess.run(
        [sys.executable, os.path.join(base_dir, "extract.py"), xlsx_path], check=True
    )


if __name__ == "__main__":
    main()
