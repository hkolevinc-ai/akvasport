from __future__ import annotations

import csv
import html
import json
import logging
import os
import re
import shutil
import sys
import time
from dataclasses import asdict, dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable
from urllib.parse import urljoin, urlparse, urlunparse

import requests
from bs4 import BeautifulSoup, Tag
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

BASE_URL = "https://akvasport.com/"
CATEGORY_URL = os.getenv(
    "CATEGORY_URL", "https://akvasport.com/category/291/primamki.html"
).strip()
PRODUCT_LIMIT = int(os.getenv("PRODUCT_LIMIT", "10"))
REQUEST_DELAY = float(os.getenv("REQUEST_DELAY", "1.2"))
DEFAULT_IN_STOCK_QTY = int(os.getenv("DEFAULT_IN_STOCK_QTY", "1"))
TEMPLATE_PATH = Path(os.getenv("TEMPLATE_PATH", "input/TEMU_TEMPLATE.xlsx"))
OUTPUT_DIR = Path(os.getenv("OUTPUT_DIR", "output"))
OUTPUT_XLSX = OUTPUT_DIR / "TEMU_AKVASPORT_UPLOAD.xlsx"
OUTPUT_CSV = OUTPUT_DIR / "akvasport_raw_export.csv"
LOG_PATH = OUTPUT_DIR / "scraper.log"

MANUFACTURER = "Tianyun Fishing Tackle Co.,Ltd"
EU_RESPONSIBLE_PERSON = "AKVASPORT EOOD"
SHIPPING_TEMPLATE = "Магазин"
HANDLING_TIME = "1 Day"
FULFILLMENT_CHANNEL = "I will ship this item myself"
COUNTRY_OF_ORIGIN = "Mainland China"

CATEGORY_NAMES = {
    "32474": "Sports & Outdoors / Hunting & Fishing / Fishing / Baits & Accessories / Bait Traps",
    "32476": "Sports & Outdoors / Hunting & Fishing / Fishing / Baits & Accessories / Baits & Attractants / Artificial Bait",
    "32477": "Sports & Outdoors / Hunting & Fishing / Fishing / Baits & Accessories / Baits & Attractants / Attractants",
    "32478": "Sports & Outdoors / Hunting & Fishing / Fishing / Baits & Accessories / Baits & Attractants / Eggs",
    "32479": "Sports & Outdoors / Hunting & Fishing / Fishing / Baits & Accessories / Baits & Attractants / Light Attractants",
}

# Template columns. Row 2 contains the user-facing labels.
COL = {
    "category": "E",
    "category_name": "F",
    "product_type": "G",
    "product_name": "L",
    "contribution_goods": "M",
    "contribution_sku": "N",
    "update_or_add": "O",
    "brand": "R",
    "description": "T",
    "bullet_1": "U",
    "bullet_2": "V",
    "bullet_3": "W",
    "bullet_4": "X",
    "bullet_5": "Y",
    "detail_img_start": "AA",
    "major_material_1": "DT",
    "major_material_2": "DU",
    "major_material_3": "DV",
    "power_mode": "EL",
    "variation_theme": "ET",
    "color": "EU",
    "size": "EV",
    "style": "EW",
    "material": "EX",
    "flavor": "EY",
    "capacity": "FA",
    "weight_variant": "FC",
    "items": "FD",
    "quantity_variant": "FE",
    "model": "FF",
    "sku_img_start": "FH",
    "quantity": "FR",
    "base_price": "FS",
    "reference_link": "FT",
    "list_price": "FU",
    "no_list_price": "FV",
    "weight_g": "FW",
    "length_cm": "FX",
    "width_cm": "FY",
    "height_cm": "FZ",
    "sku_type": "GA",
    "individually_packed": "GB",
    "total_packaging_qty": "GC",
    "packaging_unit": "GD",
    "shipping_template": "GL",
    "handling_time": "GM",
    "fulfillment": "GN",
    "country_origin": "GP",
    "product_identification": "IQ",
    "manufacturer": "IR",
    "eu_responsible": "IS",
}

REQUIRED_BY_CATEGORY = {
    cat: {
        "E", "L", "M", "N", "DT", "DU", "DV", "ET", "FH", "FR", "FS",
        "FU", "FW", "FX", "FY", "FZ", "GB", "GC", "GD", "GL", "GP",
        "IQ", "IR", "IS",
    }
    for cat in CATEGORY_NAMES
}
REQUIRED_BY_CATEGORY["32479"] = (
    REQUIRED_BY_CATEGORY["32479"] - {"DT", "DU", "DV"}
) | {"EL"}

logger = logging.getLogger("akvasport")


@dataclass
class Variant:
    sku: str
    option_values: dict[str, str] = field(default_factory=dict)
    price_eur: Decimal = Decimal("0")
    quantity: int = 0
    weight_g: float | None = None
    image_urls: list[str] = field(default_factory=list)


@dataclass
class Product:
    url: str
    title: str
    parent_sku: str
    brand: str
    description: str
    bullets: list[str]
    category_id: str
    detail_images: list[str]
    attributes: dict[str, str]
    list_price_eur: Decimal | None
    variants: list[Variant]


def setup_logging() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.handlers[:] = [stream, file_handler]


def build_session() -> requests.Session:
    retry = Retry(
        total=5,
        connect=5,
        read=5,
        backoff_factor=1.0,
        status_forcelist=(429, 500, 502, 503, 504),
        allowed_methods=frozenset(("GET", "POST")),
        raise_on_status=False,
    )
    session = requests.Session()
    session.headers.update(
        {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/136.0 Safari/537.36"
            ),
            "Accept-Language": "bg-BG,bg;q=0.9,en;q=0.7",
        }
    )
    session.mount("https://", HTTPAdapter(max_retries=retry))
    session.mount("http://", HTTPAdapter(max_retries=retry))
    return session


def fetch(session: requests.Session, url: str, *, method: str = "GET", **kwargs) -> requests.Response:
    time.sleep(max(0.0, REQUEST_DELAY))
    response = session.request(method, url, timeout=45, **kwargs)
    response.raise_for_status()
    response.encoding = response.apparent_encoding or response.encoding or "utf-8"
    return response


def clean_text(value: str | None) -> str:
    if not value:
        return ""
    value = html.unescape(value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def normalized_url(url: str) -> str:
    absolute = urljoin(BASE_URL, url)
    parsed = urlparse(absolute)
    return urlunparse(("https", parsed.netloc.lower(), parsed.path, "", "", ""))


def original_image_url(url: str) -> str:
    url = normalized_url(url)
    replacements = (
        (".thumb.webp", ".webp"),
        (".box.webp", ".webp"),
        (".large.webp", ".webp"),
        (".thumb.jpg", ".jpg"),
        (".box.jpg", ".jpg"),
        (".large.jpg", ".jpg"),
        (".thumb.jpeg", ".jpeg"),
        (".box.jpeg", ".jpeg"),
        (".large.jpeg", ".jpeg"),
        (".thumb.png", ".png"),
        (".box.png", ".png"),
        (".large.png", ".png"),
    )
    for old, new in replacements:
        url = url.replace(old, new)
    return url


def unique(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if not item or item in seen:
            continue
        seen.add(item)
        result.append(item)
    return result


def collect_product_urls(session: requests.Session, category_url: str, limit: int) -> list[str]:
    urls: list[str] = []
    seen_pages: set[str] = set()
    page_url = category_url

    while page_url and page_url not in seen_pages:
        seen_pages.add(page_url)
        logger.info("Category page: %s", page_url)
        soup = BeautifulSoup(fetch(session, page_url).text, "lxml")

        page_products = []
        for anchor in soup.select('a[href*="/product/"]'):
            href = anchor.get("href")
            if not href:
                continue
            url = normalized_url(href)
            if re.search(r"/product/\d+/", url):
                page_products.append(url)
        for url in unique(page_products):
            if url not in urls:
                urls.append(url)
                if limit > 0 and len(urls) >= limit:
                    return urls

        next_link = soup.select_one('link[rel="next"]') or soup.select_one('a[rel="next"]')
        if next_link and next_link.get("href"):
            page_url = normalized_url(next_link["href"])
        else:
            page_url = ""

    return urls


def parse_decimal(value: str | float | int | None) -> Decimal | None:
    if value is None:
        return None
    cleaned = re.sub(r"[^0-9,.-]", "", str(value)).replace(",", ".")
    try:
        return Decimal(cleaned).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def extract_json_assignment(page: str, variable: str) -> dict:
    pattern = rf"{re.escape(variable)}\s*=\s*(\{{.*?\}})\s*;"
    match = re.search(pattern, page, flags=re.DOTALL)
    if not match:
        return {}
    try:
        return json.loads(match.group(1))
    except json.JSONDecodeError:
        logger.warning("Could not decode %s", variable)
        return {}


def parse_attributes(soup: BeautifulSoup) -> dict[str, str]:
    result: dict[str, str] = {}
    for item in soup.select(".c-product-page__attribute-item"):
        name = clean_text(
            item.select_one(".c-product-page__attribute-name").get_text(" ", strip=True)
            if item.select_one(".c-product-page__attribute-name")
            else ""
        ).rstrip(":")
        value_node = item.select_one(".c-product-page__attribute-value")
        value = clean_text(value_node.get_text(" ", strip=True) if value_node else "")
        if name and value:
            result[name] = value
    return result


def extract_description(soup: BeautifulSoup) -> str:
    node = soup.select_one("#product-detailed-description")
    if not node:
        meta = soup.select_one('meta[property="og:description"]')
        return clean_text(meta.get("content") if meta else "")
    clone = BeautifulSoup(str(node), "lxml")
    for bad in clone.select("img, script, style, iframe, video, source, noscript"):
        bad.decompose()
    text = clone.get_text("\n", strip=True)
    lines = unique(clean_text(line) for line in text.splitlines() if clean_text(line))
    return "\n".join(lines)[:5000]


def extract_images(soup: BeautifulSoup) -> tuple[list[str], list[tuple[str, str]]]:
    gallery = soup.select_one("#product-images")
    if not gallery:
        gallery = soup
    images: list[tuple[str, str]] = []
    for anchor in gallery.select("a[href], a[ref]"):
        raw = anchor.get("ref") or anchor.get("href") or ""
        if "/userfiles/productimages/" not in raw:
            continue
        label = clean_text(anchor.get("title"))
        image = anchor.select_one("img")
        if image:
            label = clean_text(f"{label} {image.get('alt', '')}")
        images.append((original_image_url(raw), label))
    for image in gallery.select("img[src]"):
        raw = image.get("data-pinch-zoom-src") or image.get("src") or ""
        if "/userfiles/productimages/" not in raw:
            continue
        images.append((original_image_url(raw), clean_text(image.get("alt"))))
    dedup: list[tuple[str, str]] = []
    seen: set[str] = set()
    for url, label in images:
        if url not in seen:
            seen.add(url)
            dedup.append((url, label))
    return [u for u, _ in dedup], dedup


def option_groups(soup: BeautifulSoup) -> tuple[dict[str, dict[str, str]], list[tuple[str, str]]]:
    groups: dict[str, dict[str, str]] = {}
    ordered_groups: list[tuple[str, str]] = []
    for select in soup.select('select[name^="Options"], select.optionFilter'):
        name = select.get("name", "")
        data_type = select.get("data-type", "")
        group_id = select.get("data-id", "") or select.get("id", "").replace("optionGroup_", "")
        match = re.search(r"Options\]\[([^]]+)\]\[(\d+)\]", name)
        if match:
            data_type, group_id = match.group(1), match.group(2)
        if not group_id:
            continue
        label_node = select.find_previous(class_=re.compile("option-name|filter-title"))
        label = clean_text(label_node.get_text(" ", strip=True) if label_node else "Option").rstrip(":")
        values = {
            str(opt.get("value")): clean_text(opt.get_text(" ", strip=True))
            for opt in select.select("option[value]")
            if str(opt.get("value")) not in {"", "0"}
        }
        if values:
            groups[group_id] = values
            ordered_groups.append((group_id, label))
    return groups, unique_pairs(ordered_groups)


def unique_pairs(items: Iterable[tuple[str, str]]) -> list[tuple[str, str]]:
    seen: set[str] = set()
    out: list[tuple[str, str]] = []
    for key, value in items:
        if key in seen:
            continue
        seen.add(key)
        out.append((key, value))
    return out


def variant_option_values(
    variant_key: str,
    groups: dict[str, dict[str, str]],
    ordered_groups: list[tuple[str, str]],
) -> dict[str, str]:
    result: dict[str, str] = {}
    ids = re.findall(r"\d+", variant_key)
    used: set[str] = set()
    for option_id in ids:
        for group_id, label in ordered_groups:
            if option_id in groups.get(group_id, {}) and option_id not in used:
                result[label] = groups[group_id][option_id]
                used.add(option_id)
                break
    return result


def match_variant_images(
    option_values: dict[str, str],
    image_pairs: list[tuple[str, str]],
    all_images: list[str],
) -> list[str]:
    values = [v.lower() for v in option_values.values() if v]
    matched = [
        url
        for url, label in image_pairs
        if values and any(re.search(rf"(^|\W){re.escape(v)}($|\W)", label.lower()) for v in values)
    ]
    neutral = [
        url
        for url, label in image_pairs
        if not any(re.search(rf"(^|\W){re.escape(v)}($|\W)", label.lower()) for v in values)
    ]
    return unique(matched + neutral + all_images)[:10]


def parse_list_price(soup: BeautifulSoup, base_price: Decimal | None) -> Decimal | None:
    selectors = (
        ".old-price .price",
        ".list-price .price",
        ".u-price__old__value",
        "[data-list-price]",
    )
    for selector in selectors:
        for node in soup.select(selector):
            raw = node.get("data-list-price") or node.get_text(" ", strip=True)
            candidate = parse_decimal(raw)
            if candidate and base_price and candidate > base_price:
                return candidate
    return None


def infer_brand(title: str) -> str:
    known = [
        "Yo-Zuri", "YO-ZURI", "Duel", "DUEL", "Owner", "Ragot", "Xesta", "SeaBuzz",
        "Filstar", "Shimano", "Daiwa", "Rapala", "Westin", "Mepps", "Blue Fox",
        "Storm", "Salmo", "DTD", "Shout", "Williamson", "Bait Breath", "Little Jack",
        "Seafloor Control", "ZetZ", "Mustad", "VMC", "Hayabusa", "Acme", "Goldy",
    ]
    lower = title.lower()
    for brand in known:
        if brand.lower() in lower:
            return brand
    first = re.split(r"\s+", title.strip())[0] if title.strip() else ""
    return first[:80]


def classify_category(title: str, description: str) -> str:
    text = f"{title} {description}".lower()
    if re.search(r"\b(led|light|lamp|светещ|светлин|лампа)\b", text):
        return "32479"
    if re.search(r"\b(egg|eggs|яйц|хайвер|икра)\b", text):
        return "32478"
    if re.search(r"\b(trap|cage|капан|капани|кош за стръв|bait trap)\b", text):
        return "32474"
    if re.search(r"\b(attractant|атрактант|дип|аромат|спрей|scent|booster|flavour|flavor)\b", text):
        return "32477"
    return "32476"


def material_values(title: str, description: str, category_id: str) -> tuple[str, str, str]:
    text = f"{title} {description}".lower()
    if category_id == "32479":
        return "", "", ""
    if "волфрам" in text or "tungsten" in text:
        return "Tungsten Steel", "Stainless Steel", "Carbon Steel"
    if "джиг" in text or "глава" in text or "lead" in text:
        return "Lead", "Carbon Steel", "Stainless Steel"
    if "силикон" in text or "silicone" in text or "soft bait" in text:
        return "Silicone", "Synthetic Rubber", "PVC"
    if "metal" in text or "блесна" in text or "пилкер" in text or "spinner" in text:
        return "Stainless Steel", "Zinc Alloy", "Copper Alloy"
    return "ABS", "Stainless Steel", "PC"


def parse_measure(attributes: dict[str, str], names: Iterable[str]) -> float | None:
    for name, value in attributes.items():
        if any(key.lower() in name.lower() for key in names):
            match = re.search(r"(\d+(?:[.,]\d+)?)", value)
            if match:
                return float(match.group(1).replace(",", "."))
    return None


def bullets_from(product: Product) -> list[str]:
    result: list[str] = []
    for key in ("Тегло", "Дължина", "Размер", "Цвят", "Материал"):
        for attr_name, value in product.attributes.items():
            if key.lower() in attr_name.lower():
                result.append(f"{attr_name}: {value}")
                break
    if product.brand:
        result.insert(0, f"Марка: {product.brand}")
    result.append("Подходяща примамка за любителски и спортен риболов.")
    result.append("Изберете конкретната вариация преди поръчка.")
    return unique(result)[:6]


def parse_product(session: requests.Session, url: str) -> Product:
    logger.info("Product: %s", url)
    page = fetch(session, url).text
    soup = BeautifulSoup(page, "lxml")

    title_node = soup.select_one("h1.c-product-page__product-name")
    title = clean_text(title_node.get_text(" ", strip=True) if title_node else "")
    if not title:
        meta = soup.select_one('meta[property="og:title"]')
        title = clean_text(meta.get("content") if meta else "")
    if not title:
        raise ValueError(f"Missing product title: {url}")

    sku_node = soup.select_one("#ProductCode")
    base_sku = clean_text(sku_node.get_text(" ", strip=True) if sku_node else "")
    product_id_match = re.search(r"productID\s*:\s*(\d+)", page)
    product_id = product_id_match.group(1) if product_id_match else ""
    parent_sku = base_sku or product_id or re.search(r"/product/(\d+)/", url).group(1)

    description = extract_description(soup)
    attributes = parse_attributes(soup)
    all_images, image_pairs = extract_images(soup)
    if not all_images:
        og_image = soup.select_one('meta[property="og:image"]')
        if og_image and og_image.get("content"):
            all_images = [original_image_url(og_image["content"])]
            image_pairs = [(all_images[0], title)]

    groups, ordered_groups = option_groups(soup)
    variants_json = extract_json_assignment(page, "SC.ProductData.productVariants")
    base_price_node = soup.select_one('[itemprop="price"]')
    base_price = parse_decimal(base_price_node.get("content") if base_price_node and base_price_node.get("content") else (base_price_node.get_text(" ", strip=True) if base_price_node else "")) or Decimal("0")
    list_price = parse_list_price(soup, base_price)

    variants: list[Variant] = []
    for key, data in variants_json.items():
        option_values = variant_option_values(key, groups, ordered_groups)
        sku = clean_text(str(data.get("ProductVariantCode") or data.get("ProductVariantID") or parent_sku))
        price = parse_decimal(data.get("ProductVariantPrice")) or base_price
        try:
            quantity = max(0, int(float(data.get("ProductVariantQuantity", 0))))
        except (TypeError, ValueError):
            quantity = 0
        variant_weight = None
        try:
            # ProductWeight is normally in kg on this platform.
            variant_weight = float(data.get("ProductWeight")) * 1000
        except (TypeError, ValueError):
            pass
        variants.append(
            Variant(
                sku=sku,
                option_values=option_values,
                price_eur=price,
                quantity=quantity,
                weight_g=variant_weight,
                image_urls=match_variant_images(option_values, image_pairs, all_images),
            )
        )

    if not variants:
        availability = soup.select_one('meta[itemprop="availability"]')
        in_stock = bool(availability and "InStock" in str(availability.get("content", "")))
        variants = [
            Variant(
                sku=parent_sku,
                option_values={"Model": clean_text(title[-80:]) or "Standard"},
                price_eur=base_price,
                quantity=DEFAULT_IN_STOCK_QTY if in_stock else 0,
                image_urls=all_images[:10],
            )
        ]

    category_id = classify_category(title, description)
    product = Product(
        url=url,
        title=title,
        parent_sku=parent_sku,
        brand=infer_brand(title),
        description=description,
        bullets=[],
        category_id=category_id,
        detail_images=all_images[:100],
        attributes=attributes,
        list_price_eur=list_price,
        variants=variants,
    )
    product.bullets = bullets_from(product)
    return product


def safe_sku(value: str, fallback: str) -> str:
    value = clean_text(value) or fallback
    value = re.sub(r"[^A-Za-z0-9._-]+", "-", value).strip("-")
    return value[:80] or fallback[:80]


def variation_theme_and_cells(variant: Variant) -> tuple[str, dict[str, str]]:
    mapping = {
        "цвят": ("Color", "EU"),
        "color": ("Color", "EU"),
        "размер": ("Size", "EV"),
        "size": ("Size", "EV"),
        "модел": ("Model", "FF"),
        "model": ("Model", "FF"),
        "тегло": ("Weight", "FC"),
        "weight": ("Weight", "FC"),
        "стил": ("Style", "EW"),
        "style": ("Style", "EW"),
        "материал": ("Material", "EX"),
        "material": ("Material", "EX"),
        "аромат": ("Flavors", "EY"),
        "flavor": ("Flavors", "EY"),
        "вкус": ("Flavors", "EY"),
        "capacity": ("Capacity", "FA"),
        "капацитет": ("Capacity", "FA"),
        "quantity": ("Quantity", "FE"),
        "количество": ("Quantity", "FE"),
    }
    components: list[str] = []
    cells: dict[str, str] = {}
    for label, value in variant.option_values.items():
        normalized = label.lower().strip()
        selected = None
        for key, pair in mapping.items():
            if key in normalized:
                selected = pair
                break
        if selected is None:
            selected = ("Model", "FF")
        theme, column = selected
        if theme not in components:
            components.append(theme)
        cells[column] = value
    if not components:
        components = ["Model"]
        cells["FF"] = "Standard"
    return " × ".join(components[:2]), cells


def append_images(ws, row: int, start_col: str, images: list[str], max_images: int) -> None:
    start = ws[start_col + "2"].column
    for offset, url in enumerate(images[:max_images]):
        ws.cell(row=row, column=start + offset, value=url)


def product_dimensions(product: Product, variant: Variant) -> tuple[float, float, float, float]:
    explicit_weight = parse_measure(product.attributes, ("тегло", "weight"))
    explicit_length = parse_measure(product.attributes, ("дължина", "length"))
    explicit_width = parse_measure(product.attributes, ("ширина", "width"))
    explicit_height = parse_measure(product.attributes, ("височина", "height", "дебелина"))

    weight = explicit_weight or variant.weight_g or 10.0
    length = explicit_length or 10.0
    width = explicit_width or max(1.0, min(10.0, length * 0.25))
    height = explicit_height or max(1.0, min(5.0, width * 0.6))
    return round(weight, 2), round(length, 2), round(width, 2), round(height, 2)


def fill_template(products: list[Product]) -> None:
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Template not found: {TEMPLATE_PATH}")
    shutil.copy2(TEMPLATE_PATH, OUTPUT_XLSX)
    wb = load_workbook(OUTPUT_XLSX)
    ws = wb["Template"]

    # Remove old data only. Reserved/header rows 1-4 are preserved.
    for row in ws.iter_rows(min_row=5, max_row=ws.max_row):
        for cell in row:
            cell.value = None

    output_row = 5
    raw_rows: list[dict] = []

    for product in products:
        mat1, mat2, mat3 = material_values(product.title, product.description, product.category_id)
        for variant in product.variants:
            theme, sale_cells = variation_theme_and_cells(variant)
            weight_g, length_cm, width_cm, height_cm = product_dimensions(product, variant)
            parent = safe_sku(product.parent_sku, f"AKVA-{output_row}")
            sku = safe_sku(variant.sku, f"{parent}-{output_row}")
            primary_image = (variant.image_urls or product.detail_images or [""])[0]

            values = {
                "E": product.category_id,
                "F": CATEGORY_NAMES[product.category_id],
                "G": "Normal product",
                "L": product.title[:500],
                "M": parent,
                "N": sku,
                "O": "Add",
                "R": product.brand,
                "T": product.description[:5000],
                "DT": mat1,
                "DU": mat2,
                "DV": mat3,
                "ET": theme,
                "FH": primary_image,
                "FR": int(variant.quantity),
                "FS": float(variant.price_eur),
                "FT": product.url,
                "FW": weight_g,
                "FX": length_cm,
                "FY": width_cm,
                "FZ": height_cm,
                "GA": "Single set",
                "GB": "Yes",
                "GC": 1,
                "GD": "piece",
                "GL": SHIPPING_TEMPLATE,
                "GM": HANDLING_TIME,
                "GN": FULFILLMENT_CHANNEL,
                "GP": COUNTRY_OF_ORIGIN,
                "IQ": sku,
                "IR": MANUFACTURER,
                "IS": EU_RESPONSIBLE_PERSON,
            }
            if product.category_id == "32479":
                values["EL"] = "Without electricity"
            if product.list_price_eur and product.list_price_eur > variant.price_eur:
                values["FU"] = float(product.list_price_eur)
            else:
                # This is Temu's accepted alternative to a genuine list price.
                values["FV"] = "N/A"

            for col, value in sale_cells.items():
                values[col] = value
            for index, bullet in enumerate(product.bullets[:6]):
                values[get_column_letter(ws["U2"].column + index)] = bullet[:700]
            for col, value in values.items():
                ws[f"{col}{output_row}"] = value

            append_images(ws, output_row, "AA", product.detail_images, 100)
            append_images(ws, output_row, "FH", variant.image_urls or product.detail_images, 10)

            raw_rows.append(
                {
                    "product_url": product.url,
                    "product_name": product.title,
                    "parent_sku": parent,
                    "sku": sku,
                    "category_id": product.category_id,
                    "category_name": CATEGORY_NAMES[product.category_id],
                    "brand": product.brand,
                    "variation_theme": theme,
                    "variation_values": json.dumps(variant.option_values, ensure_ascii=False),
                    "quantity": variant.quantity,
                    "price_eur": str(variant.price_eur),
                    "list_price_eur": str(product.list_price_eur or ""),
                    "weight_g": weight_g,
                    "length_cm": length_cm,
                    "width_cm": width_cm,
                    "height_cm": height_cm,
                    "main_image": primary_image,
                    "all_images": " | ".join(variant.image_urls or product.detail_images),
                    "description": product.description,
                    "manufacturer": MANUFACTURER,
                    "eu_responsible_person": EU_RESPONSIBLE_PERSON,
                }
            )
            output_row += 1

    validate_required_rows(ws, 5, output_row - 1)
    wb.save(OUTPUT_XLSX)
    write_raw_csv(raw_rows)
    logger.info("Saved %s (%d SKU rows)", OUTPUT_XLSX, len(raw_rows))


def validate_required_rows(ws, first_row: int, last_row: int) -> None:
    errors: list[str] = []
    for row in range(first_row, last_row + 1):
        category = str(ws[f"E{row}"].value or "")
        required = REQUIRED_BY_CATEGORY.get(category, set())
        for column in required:
            value = ws[f"{column}{row}"].value
            # FU may be blank when FV explicitly contains N/A, as permitted by Temu.
            if column == "FU" and ws[f"FV{row}"].value == "N/A":
                continue
            if value is None or (isinstance(value, str) and not value.strip()):
                errors.append(f"row {row}: empty required field {column} ({ws[f'{column}2'].value})")
        if not ws[f"FH{row}"].value:
            errors.append(f"row {row}: missing SKU image")
        price = ws[f"FS{row}"].value
        if price is None or float(price) <= 0:
            errors.append(f"row {row}: invalid price")
    if errors:
        preview = "\n".join(errors[:50])
        raise ValueError(f"Template validation failed ({len(errors)} errors):\n{preview}")


def write_raw_csv(rows: list[dict]) -> None:
    if not rows:
        return
    with OUTPUT_CSV.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def main() -> int:
    setup_logging()
    logger.info("Start. Category=%s, product limit=%s", CATEGORY_URL, PRODUCT_LIMIT)
    session = build_session()
    urls = collect_product_urls(session, CATEGORY_URL, PRODUCT_LIMIT)
    if not urls:
        raise RuntimeError("No product URLs were found in the selected category")
    logger.info("Found %d product URLs", len(urls))

    products: list[Product] = []
    failed: list[tuple[str, str]] = []
    for index, url in enumerate(urls, start=1):
        try:
            product = parse_product(session, url)
            products.append(product)
            logger.info(
                "Parsed %d/%d: %s (%d variants)",
                index,
                len(urls),
                product.title,
                len(product.variants),
            )
        except Exception as exc:  # keep the run useful if one product is malformed
            logger.exception("Failed product %s", url)
            failed.append((url, str(exc)))

    if not products:
        raise RuntimeError("All product pages failed to parse")
    fill_template(products)
    if failed:
        logger.warning("%d products failed. See log for details.", len(failed))
    logger.info("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
