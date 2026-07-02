"""Извлечение изображений и характеристик из XLSX прайс-листов.

Результат: папка product_images/ + обновление fetched_specs.json
Запуск: python extract_images.py [xlsx_file ...]
"""

import sys, os, json, zipfile, shutil, re
import xml.etree.ElementTree as ET

sys.stdout.reconfigure(encoding="utf-8")
DIR = os.path.dirname(os.path.abspath(__file__))
IMG_DIR = os.path.join(DIR, "product_images")
OUTPUT_FILE = os.path.join(DIR, "fetched_specs.json")

NS_XDR  = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
NS_A    = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

# Паттерны для извлечения характеристик из описаний
SPEC_PATTERNS = [
    (r"(\d+[\.,]?\d*)\s*[Мм][Пп]", "Разрешение"),
    (r"объектив\s+([\d.,/-]+\s*мм)", "Объектив"),
    (r"угол\s+обзора\s+([\d.,°/]+°?)", "Угол обзора"),
    (r"ИК[- ]подсветк[уа]\s+до\s+(\d+\s*м)", "ИК-подсветка"),
    (r"Led[- ]подсветк[уа]\s+(\d+\s*м)", "LED-подсветка"),
    (r"(H\.26[45]\+?(?:/H\.26[45]\+?)*)", "Сжатие"),
    (r"\bIP(\d{2,})\b", "Степень защиты"),
    (r"(PoE[^\s,;]*)", "Питание"),
    (r"DC\s*(\d+)\s*В", "Питание DC"),
    (r"([-\d]+\s*°C\s*[до]+\s*[+]?\d+\s*°C)", "Температура"),
    (r"(\d+[xх]\d+@\d+\s*к/с)", "Разрешение потока"),
    (r"матриц[аы]\s+(1/[\d.,]+''[^;,\n]{0,30})", "Матрица"),
    (r"слот.*?microSD\s+до\s+([\d]+\s*[ГТ]б)", "Карта памяти"),
    (r"([\d.,]+)\s*[Лл][кк]@", "Чувствительность"),
]

# Слова которые НЕ являются артикулами (заголовки, разделы)
SKIP_WORDS = {
    "фото", "изображение", "описание", "photo", "image", "модель", "model",
    "артикул", "наименование", "colorvu", "wizsense", "wizmind", "acusense",
    "eol", "замена", "новинка", "серия", "серии", "разрешение", "тип",
    "примечание", "содержание", "hd-tvi", "ip-камер", "регистратор",
    "коммутатор", "домофон", "скуд", "аксессуар", "замок", "сигнализ",
    "код заказа", "код", "заказа", "примечани", "рекомендован", "ррц",
    "цена", "price", "msrp",
}

BRAND_PREFIXES = [
    "hiwatch", "dahua", "hikvision", "zkteco", "tiandy", "temid",
    "ezviz", "optimus", "belpark",
]


def is_valid_model(text):
    if not text or len(text) < 3 or len(text) > 80:
        return False
    lower = text.lower().strip()
    if any(lower == w or lower.startswith(w + " ") for w in SKIP_WORDS):
        return False
    if "₽" in text or "руб" in lower:
        return False
    if re.match(r"^[\d.,\s]+$", text):
        return False
    # Артикул обычно содержит буквы + цифры или дефисы
    if not re.search(r"[A-Za-zА-Яа-я]", text):
        return False
    return True


def strip_brand(name):
    lower = name.lower().strip()
    for bp in BRAND_PREFIXES:
        if lower.startswith(bp + " "):
            return name[len(bp):].strip()
    return name


def normalize(s):
    return re.sub(r"\s+", " ", s.lower().strip())


def norm_no_size(s):
    """Убирает суффиксы размера объектива: '(2.8 mm)', '(4mm)'."""
    s = re.sub(r"\s*\([\d.,]+\s*mm?\)", "", s, flags=re.I)
    return re.sub(r"\s+", " ", s).strip()


def parse_rels(z, rels_path):
    try:
        tree = ET.fromstring(z.read(rels_path))
        return {r.get("Id"): r.get("Target") for r in tree}
    except Exception:
        return {}


def parse_drawing(z, drawing_path, rels_path):
    """Возвращает dict: row (0-based) -> имя файла изображения."""
    row_to_img = {}
    try:
        rels = parse_rels(z, rels_path)
        tree = ET.fromstring(z.read(drawing_path))
        for anchor in tree:
            tag = anchor.tag.split("}")[-1]
            if tag not in ("twoCellAnchor", "oneCellAnchor", "absoluteAnchor"):
                continue
            fr = anchor.find(f"{{{NS_XDR}}}from")
            if fr is None:
                continue
            row_el = fr.find(f"{{{NS_XDR}}}row")
            if row_el is None:
                continue
            row = int(row_el.text)
            blip = anchor.find(f".//{{{NS_A}}}blip")
            if blip is None:
                continue
            rid = blip.get(f"{{{NS_R}}}embed")
            if rid and rid in rels:
                row_to_img[row] = os.path.basename(rels[rid])
    except Exception as e:
        print(f"  [drawing error] {e}")
    return row_to_img


def parse_sheet(z, sheet_path):
    """Возвращает dict: row_idx (0-based) -> dict{col: value}."""
    rows = {}
    try:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            ss_tree = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in ss_tree.findall(f".//{{{NS_MAIN}}}si"):
                parts = [t.text for t in si.iter(f"{{{NS_MAIN}}}t") if t.text]
                shared.append("".join(parts))

        tree = ET.fromstring(z.read(sheet_path))
        for row_el in tree.findall(f".//{{{NS_MAIN}}}row"):
            r_idx = int(row_el.get("r", 1)) - 1
            cells = {}
            for cell in row_el.findall(f"{{{NS_MAIN}}}c"):
                ref = cell.get("r", "")
                m = re.match(r"([A-Z]+)", ref)
                if m:
                    col = sum((ord(c) - 64) * (26 ** i)
                              for i, c in enumerate(reversed(m.group(1)))) - 1
                else:
                    col = 0
                t = cell.get("t", "")
                v_el = cell.find(f"{{{NS_MAIN}}}v")
                val = ""
                if v_el is not None and v_el.text:
                    val = shared[int(v_el.text)] if t == "s" and v_el.text.isdigit() else v_el.text
                if val:
                    cells[col] = val.strip()
            rows[r_idx] = cells
    except Exception as e:
        print(f"  [sheet error] {e}")
    return rows


def get_sheet_drawing_pairs(z):
    pairs = []
    for name in z.namelist():
        if name.startswith("xl/worksheets/_rels/") and name.endswith(".rels"):
            sheet_name = name.replace("xl/worksheets/_rels/", "").replace(".rels", "")
            sheet_path = f"xl/worksheets/{sheet_name}"
            rels = parse_rels(z, name)
            for target in rels.values():
                if "drawing" in target.lower():
                    drawing_name = os.path.basename(target)
                    pairs.append((
                        sheet_path,
                        f"xl/drawings/{drawing_name}",
                        f"xl/drawings/_rels/{drawing_name}.rels",
                    ))
    return pairs


def parse_specs_from_description(desc):
    specs = {}
    if not desc:
        return specs
    for pattern, key in SPEC_PATTERNS:
        m = re.search(pattern, desc, re.IGNORECASE)
        if m:
            val = m.group(1) if m.lastindex else m.group(0)
            if val and val.strip() not in ("", "дБ"):
                specs[key] = val.strip()
    return specs


def find_model_and_desc(row_cells, rows, row_idx):
    model = ""
    description = ""
    for col in sorted(row_cells.keys())[:6]:
        val = row_cells[col]
        if is_valid_model(val):
            model = val
            for dcol in sorted(row_cells.keys()):
                dval = row_cells[dcol]
                if len(dval) > 50 and dcol != col:
                    description = dval
                    break
            break
    return model, description


def extract_from_xlsx(xlsx_path):
    """Возвращает dict: артикул -> {image_file, description, specs}."""
    result = {}
    os.makedirs(IMG_DIR, exist_ok=True)

    with zipfile.ZipFile(xlsx_path) as z:
        media_files = [f for f in z.namelist() if f.startswith("xl/media/")]
        print(f"  Медиафайлов: {len(media_files)}")
        for mf in media_files:
            fname = os.path.basename(mf)
            out_path = os.path.join(IMG_DIR, fname)
            if not os.path.exists(out_path):
                with z.open(mf) as src, open(out_path, "wb") as dst:
                    shutil.copyfileobj(src, dst)

        pairs = get_sheet_drawing_pairs(z)
        print(f"  Листов с рисунками: {len(pairs)}")

        for sheet_path, drawing_path, drawing_rels in pairs:
            if sheet_path not in z.namelist() or drawing_path not in z.namelist():
                continue
            row_to_img = parse_drawing(z, drawing_path, drawing_rels)
            if not row_to_img:
                continue
            rows = parse_sheet(z, sheet_path)
            sheet_matches = 0

            for row_idx, img_file in row_to_img.items():
                cells = rows.get(row_idx, {})
                model, desc = find_model_and_desc(cells, rows, row_idx)
                if not model:
                    for nearby in [row_idx - 1, row_idx + 1, row_idx + 2]:
                        cells2 = rows.get(nearby, {})
                        model, desc = find_model_and_desc(cells2, rows, nearby)
                        if model:
                            break
                if model and img_file:
                    specs = parse_specs_from_description(desc)
                    result[model] = {
                        "image_file": img_file,
                        "description": desc,
                        "specs": specs,
                    }
                    sheet_matches += 1

            print(f"  {sheet_path}: {sheet_matches} товаров с фото")

    return result


def update_fetched_specs(extracted):
    """Обновляет fetched_specs.json: добавляет image_url и specs из прайсов."""
    if not os.path.exists(OUTPUT_FILE):
        specs_list = []
    else:
        with open(OUTPUT_FILE, encoding="utf-8") as f:
            specs_list = json.load(f)

    # Строим несколько индексов из прайса для гибкого совпадения
    ex_full   = {normalize(k): v for k, v in extracted.items()}
    ex_nb     = {normalize(strip_brand(k)): v for k, v in extracted.items()}
    ex_nosize = {norm_no_size(normalize(strip_brand(k))): v for k, v in extracted.items()}

    updated_img = 0
    updated_specs = 0
    added_new = 0

    for entry in specs_list:
        name = entry.get("name", "")
        n_full  = normalize(name)
        n_nb    = normalize(strip_brand(name))
        n_ns    = norm_no_size(n_nb)

        match = (ex_full.get(n_full)
                 or ex_nb.get(n_nb)
                 or ex_nosize.get(n_ns))

        # Частичное совпадение: артикул из склада содержится в ключе прайса
        if not match and n_nb:
            for k, v in ex_nb.items():
                if n_nb and (n_nb in k or k in n_nb) and len(n_nb) > 5:
                    match = v
                    break

        if match:
            if match.get("image_file") and not entry.get("image_url"):
                entry["image_url"] = f"product_images/{match['image_file']}"
                updated_img += 1
            if match.get("specs"):
                if not entry.get("specs"):
                    entry["specs"] = match["specs"]
                    updated_specs += 1
                else:
                    for k, v in match["specs"].items():
                        if k not in entry["specs"]:
                            entry["specs"][k] = v

    # Добавляем новые записи из прайса (которых нет в fetched_specs)
    existing_names = {normalize(e["name"]) for e in specs_list}
    for model, data in extracted.items():
        if normalize(model) not in existing_names and is_valid_model(model):
            if data.get("image_file") or data.get("specs"):
                brand = model.split()[0]
                specs_list.append({
                    "name": model,
                    "brand": brand,
                    "specs": data.get("specs", {}),
                    "image_url": f"product_images/{data['image_file']}" if data.get("image_file") else "",
                    "source": "pricelist",
                })
                added_new += 1

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        json.dump(specs_list, f, ensure_ascii=False, indent=2)

    print(f"  Добавлено изображений в существующие: {updated_img}")
    print(f"  Добавлено характеристик в существующие: {updated_specs}")
    print(f"  Новых записей из прайса: {added_new}")


def main():
    files = sys.argv[1:]
    if not files:
        files = [
            os.path.join(DIR, f) for f in os.listdir(DIR)
            if f.endswith(".xlsx") and any(
                w in f.lower() for w in ["pricelist", "прайс"]
            )
        ]
    if not files:
        print("Укажи XLSX файл(ы)")
        sys.exit(1)

    all_extracted = {}
    for xlsx_path in files:
        if not os.path.exists(xlsx_path):
            print(f"Файл не найден: {xlsx_path}")
            continue
        print(f"\n=== {os.path.basename(xlsx_path)} ===")
        extracted = extract_from_xlsx(xlsx_path)
        print(f"  Итого товаров с фото: {len(extracted)}")
        all_extracted.update(extracted)

    print(f"\n=== Всего из прайсов: {len(all_extracted)} артикулов ===")
    update_fetched_specs(all_extracted)
    print("\nГотово! Запусти: python build_html.py")


if __name__ == "__main__":
    main()
