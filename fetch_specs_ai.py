"""Парсинг характеристик для позиций из missing_1000.json.

Алгоритм для каждой позиции:
1. Запрашиваем страницу производителя или Google
2. Передаём HTML в Claude API → получаем структурированные характеристики
3. Сохраняем результат в fetched_specs.json (по одной — можно прерывать и продолжать)

Запуск: python fetch_specs_ai.py [--limit 50] [--brand dahua]
"""

import sys, os, json, time, argparse, re
import urllib.parse
import urllib.request
import urllib.error

sys.stdout.reconfigure(encoding="utf-8")

DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_FILE = os.path.join(DIR, "missing_1000.json")
OUTPUT_FILE = os.path.join(DIR, "fetched_specs.json")

ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml",
    "Accept-Language": "ru,en;q=0.9",
}

# Прямые URL по артикулу для известных брендов
BRAND_URL_BUILDERS = {
    "Dahua":     lambda m: f"https://www.dahua.com/en/search/?q={urllib.parse.quote(m)}",
    "Hikvision": lambda m: f"https://www.hikvision.com/en/search/?keywords={urllib.parse.quote(m)}",
    "ZKTeco":    lambda m: f"https://www.zkteco.eu/index.php?c=product&a=search&searchValue={urllib.parse.quote(m)}",
    "Tiandy":    lambda m: f"https://www.tiandy.com/products/search.html?keywords={urllib.parse.quote(m)}",
}

CLAUDE_MODEL = "claude-haiku-4-5-20251001"
MAX_HTML_CHARS = 12000


def fetch_url(url, timeout=15):
    req = urllib.request.Request(url, headers=HEADERS)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            raw = resp.read()
            for enc in ("utf-8", "cp1251", "latin-1"):
                try:
                    return raw.decode(enc)
                except Exception:
                    continue
    except Exception as e:
        return None


def strip_html(html):
    """Убираем теги, оставляем текст с пробелами."""
    html = re.sub(r"<(script|style)[^>]*>.*?</\1>", " ", html, flags=re.S | re.I)
    html = re.sub(r"<[^>]+>", " ", html)
    html = re.sub(r"[ \t]{2,}", " ", html)
    html = re.sub(r"\n{3,}", "\n\n", html)
    return html.strip()


def extract_image_url(html: str, base_url: str) -> str:
    """Ищем первое изображение товара на странице."""
    patterns = [
        r'<img[^>]+class="[^"]*(?:product|main|hero|detail)[^"]*"[^>]+src="([^"]+)"',
        r'<img[^>]+src="([^"]+(?:product|item|goods)[^"]*\.(?:jpg|png|webp))"',
        r'"image"\s*:\s*"(https?://[^"]+\.(?:jpg|png|webp))"',
        r'<img[^>]+src="(https?://[^"]+\.(?:jpg|png|webp))"',
    ]
    for pattern in patterns:
        m = re.search(pattern, html, re.I)
        if m:
            url = m.group(1)
            if url.startswith("//"):
                url = "https:" + url
            elif url.startswith("/"):
                from urllib.parse import urlparse
                p = urlparse(base_url)
                url = f"{p.scheme}://{p.netloc}{url}"
            return url
    return ""


def ask_claude(product_name: str, page_text: str) -> dict:
    """Отправляем текст страницы в Claude, просим вернуть характеристики JSON."""
    import json as _json
    prompt = f"""Ты эксперт по оборудованию видеонаблюдения и безопасности.
Ниже — текст страницы с сайта производителя для товара: {product_name}

Извлеки технические характеристики в виде JSON-объекта.
Ключи — названия параметров на русском (Разрешение, Тип матрицы, ИК-подсветка и т.д.).
Значения — строки.
Если характеристик нет на странице — верни пустой объект {{}}.
Верни ТОЛЬКО JSON без пояснений.

Текст страницы:
{page_text[:MAX_HTML_CHARS]}
"""

    payload = {
        "model": CLAUDE_MODEL,
        "max_tokens": 1024,
        "messages": [{"role": "user", "content": prompt}],
    }
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data=data,
        headers={
            "x-api-key": ANTHROPIC_API_KEY,
            "anthropic-version": "2023-06-01",
            "content-type": "application/json",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = _json.loads(resp.read().decode("utf-8"))
            text = result["content"][0]["text"].strip()
            # Извлекаем JSON из ответа
            m = re.search(r"\{.*\}", text, re.S)
            if m:
                return _json.loads(m.group())
            return {}
    except Exception as e:
        print(f"  Claude API ошибка: {e}")
        return {}


def process_item(item: dict) -> dict:
    name = item["name"]
    brand = item["brand"]

    # Определяем URL для запроса
    url_builder = BRAND_URL_BUILDERS.get(brand)
    if url_builder:
        url = url_builder(name)
    else:
        url = item.get("search_url", "")

    print(f"  → {url[:80]}")
    html = fetch_url(url) if url else None

    if not html or len(html) < 200:
        return {"name": name, "brand": brand, "specs": {}, "image_url": "", "source": "not_found"}

    image_url = extract_image_url(html, url)
    text = strip_html(html)

    if not ANTHROPIC_API_KEY:
        # Без API — простой regex-поиск характеристик
        specs = {}
        patterns = [
            (r"[Рр]азрешение[:\s]+([^\n<]{3,40})", "Разрешение"),
            (r"[Мм]атриц[аы][:\s]+([^\n<]{3,40})", "Тип матрицы"),
            (r"[Оо]бъектив[:\s]+([^\n<]{3,40})", "Объектив"),
            (r"ИК.{0,10}([0-9]+\s*м)", "ИК-подсветка"),
            (r"[Пп]итани[ея][:\s]+([^\n<]{3,40})", "Питание"),
            (r"[Зз]ащита[:\s]+(IP\d+)", "Степень защиты"),
        ]
        for pattern, key in patterns:
            m = re.search(pattern, text)
            if m:
                specs[key] = m.group(1).strip()
        return {"name": name, "brand": brand, "specs": specs, "image_url": image_url, "source": url}

    specs = ask_claude(name, text)
    return {"name": name, "brand": brand, "specs": specs, "image_url": image_url, "source": url}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--limit", type=int, default=50, help="Сколько позиций обработать за запуск")
    parser.add_argument("--brand", type=str, default="", help="Фильтр по бренду (например: Dahua)")
    parser.add_argument("--skip-done", action="store_true", default=True, help="Пропускать уже обработанные")
    args = parser.parse_args()

    if not ANTHROPIC_API_KEY:
        print("ВНИМАНИЕ: ANTHROPIC_API_KEY не задан — Claude API недоступен.")
        print("  Установите: set ANTHROPIC_API_KEY=sk-ant-...")
        print("  Без него используется базовый regex-парсинг.\n")

    with open(INPUT_FILE, encoding="utf-8") as f:
        items = json.load(f)

    # Загружаем уже обработанные
    done = {}
    if os.path.exists(OUTPUT_FILE) and args.skip_done:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            for entry in json.load(f):
                done[entry["name"]] = entry

    # Фильтрация
    todo = [x for x in items
            if x["name"] not in done
            and (not args.brand or x["brand"].lower() == args.brand.lower())]

    todo = todo[:args.limit]
    print(f"Обработать: {len(todo)} позиций (из {len(items)} в списке, {len(done)} уже готово)")

    results = list(done.values())
    new_count = 0
    found_count = 0

    for i, item in enumerate(todo, 1):
        print(f"[{i}/{len(todo)}] {item['name']}")
        try:
            result = process_item(item)
            results.append(result)
            if result["specs"]:
                found_count += 1
                print(f"  Найдено: {len(result['specs'])} параметров")
            else:
                print(f"  Не найдено")
            new_count += 1
        except Exception as e:
            print(f"  Ошибка: {e}")
            results.append({"name": item["name"], "brand": item["brand"], "specs": {}, "source": "error"})
            new_count += 1

        # Сохраняем после каждой позиции (можно прерывать)
        with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

        time.sleep(0.5)

    print(f"\nГотово: обработано {new_count}, с характеристиками: {found_count}")
    print(f"Результат: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
