"""Опрос Telegram-бота на наличие команды обновления остатков.

Используется в GitHub Actions: запускается по расписанию, проверяет новые
сообщения через getUpdates, и если от авторизованного chat_id пришла
команда /update_stock — выставляет output triggered=true для workflow.
Offset последнего обработанного апдейта сохраняется в файл и коммитится
в репозиторий, чтобы между запусками Actions не терять позицию.
"""

import os
import sys

import requests

DIR = os.path.dirname(os.path.abspath(__file__))
OFFSET_FILE = os.path.join(DIR, ".github", "telegram_offset.txt")

TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
ALLOWED_CHAT_ID = os.environ.get("TELEGRAM_CHAT_ID")
COMMAND = "/update_stock"


def get_offset():
    if os.path.exists(OFFSET_FILE):
        content = open(OFFSET_FILE, encoding="utf-8").read().strip()
        return int(content) if content else 0
    return 0


def save_offset(offset):
    with open(OFFSET_FILE, "w", encoding="utf-8") as f:
        f.write(str(offset))


def send_message(chat_id, text):
    requests.post(
        f"https://api.telegram.org/bot{TOKEN}/sendMessage",
        data={"chat_id": chat_id, "text": text},
        timeout=15,
    )


def set_output(name, value):
    gh_output = os.environ.get("GITHUB_OUTPUT")
    if gh_output:
        with open(gh_output, "a", encoding="utf-8") as f:
            f.write(f"{name}={value}\n")


def main():
    if not TOKEN or not ALLOWED_CHAT_ID:
        print("TELEGRAM_BOT_TOKEN или TELEGRAM_CHAT_ID не заданы, пропуск опроса")
        set_output("triggered", "false")
        return

    offset = get_offset()
    resp = requests.get(
        f"https://api.telegram.org/bot{TOKEN}/getUpdates",
        params={"offset": offset + 1, "timeout": 0},
        timeout=15,
    )
    resp.raise_for_status()
    updates = resp.json().get("result", [])

    triggered = False
    max_update_id = offset
    for upd in updates:
        max_update_id = max(max_update_id, upd["update_id"])
        msg = upd.get("message") or upd.get("channel_post")
        if not msg:
            continue
        chat_id = str(msg["chat"]["id"])
        text = (msg.get("text") or "").strip()

        if chat_id != str(ALLOWED_CHAT_ID):
            send_message(chat_id, "У вас нет доступа к этому боту.")
            continue

        if text.startswith(COMMAND):
            triggered = True

    save_offset(max_update_id)
    set_output("triggered", "true" if triggered else "false")

    if triggered:
        send_message(ALLOWED_CHAT_ID, "Запускаю обновление остатков...")
        print("Команда получена, запускаю обновление")
    else:
        print("Новых команд нет")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"Ошибка опроса Telegram: {e}", file=sys.stderr)
        set_output("triggered", "false")
