"""Скачивание XLS-файлов складов с Google Диска через Service Account."""

import os
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io

DIR = os.path.dirname(os.path.abspath(__file__))

FILES = {
    "Минск.xls": "1dDqOA36DIo11nIcYGbbRWoVBiWihlMd0",
    "Москва.xls": "1ioGPOb8mNRFHHyXtb9BL6XAEWS-M4SLr",
}

def main():
    sa_json = os.environ.get("GOOGLE_SA_KEY")
    if not sa_json:
        print("GOOGLE_SA_KEY не задан, пропуск скачивания")
        return

    sa_json = sa_json.strip()
    # Снять внешние кавычки, если секрет был сохранён как "{ ... }"
    if sa_json.startswith('"') and sa_json.endswith('"'):
        sa_json = sa_json[1:-1].replace('\\"', '"')

    print(f"GOOGLE_SA_KEY первые 20 символов: {repr(sa_json[:20])}")

    try:
        creds_info = json.loads(sa_json)
    except json.JSONDecodeError as e:
        print(f"Ошибка парсинга GOOGLE_SA_KEY: {e}")
        print(f"Длина строки: {len(sa_json)}, первые 50 символов: {repr(sa_json[:50])}")
        raise
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
