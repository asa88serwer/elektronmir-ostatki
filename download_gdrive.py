"""Скачивание XLS-файлов складов с Google Диска через Service Account."""

import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

DIR = os.path.dirname(os.path.abspath(__file__))

FILES = {
    "Минск.xls": "1-xWFT4u3vKELBgdrc-ahQeloKVU4PU-j",
    "Москва.xls": "1MNKOhDazKAinY-FRRsc7kdIzWKSmridY",
}

def main():
    sa_json = os.environ.get("GOOGLE_SA_KEY")
    if not sa_json:
        print("GOOGLE_SA_KEY не задан, пропуск скачивания")
        return

    sa_json = sa_json.strip()
    creds_info = json.loads(sa_json)
    creds = service_account.Credentials.from_service_account_info(
        creds_info, scopes=["https://www.googleapis.com/auth/drive.readonly"]
    )
    service = build("drive", "v3", credentials=creds)

    for filename, file_id in FILES.items():
        print(f"Скачивание {filename} ...")
        request = service.files().get_media(fileId=file_id)
        buf = io.BytesIO()
        downloader = MediaIoBaseDownload(buf, request)
        done = False
        while not done:
            _, done = downloader.next_chunk()
        path = os.path.join(DIR, filename)
        with open(path, "wb") as f:
            f.write(buf.getvalue())
        print(f"  Сохранён: {path} ({len(buf.getvalue())} байт)")

if __name__ == "__main__":
    main()
