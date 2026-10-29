# wb_competitors.py
# Python 3.10+
#
# Что делает:
# 1) читает input.csv с колонкой nmID
# 2) по каждому вашему nmID получает свою карточку через официальный WB API
# 3) строит поисковый запрос по title
# 4) ищет кандидатов в публичном поиске WB
# 5) фильтрует по subjectId
# 6) пытается вытащить страну производства конкурента из HTML карточки
# 7) считает score и сохраняет top-N в result.csv
#
# Зависимости:
#   pip install requests
#
# Запуск:
#   python wb_concurents.py --input input.csv --output result.csv --top 5
#   python wb_concurents.py --fetch-all --output result.csv --top 5   # nmID из API
#
# Set WB_TOKEN explicitly for the optional live WB lookup.
#
# Формат input.csv:
#   nmID
#   12345678
#   23456789

from __future__ import annotations

import argparse
import csv
import html
import json
import os
import random
import re
import sys
import time
from dataclasses import dataclass, asdict
from difflib import SequenceMatcher
from typing import Any, Dict, Iterable, List, Optional, Tuple

import requests


CONTENT_API_URL = "https://content-api.wildberries.ru/content/v2/get/cards/list"
SEARCH_ENDPOINTS = [
    "https://u-search.wb.ru/exactmatch/ru/common/v18/search",
    "https://search.wb.ru/exactmatch/ru/common/v18/search",
    "https://search.wb.ru/exactmatch/ru/common/v4/search",
]
DETAIL_API_URL = "https://card.wb.ru/cards/detail"

DEFAULT_DEST = "-1257786"
DEFAULT_TOP = 5
DEFAULT_SEARCH_PAGES = 2
DEFAULT_SEARCH_PAGE_SIZE = 50
DEFAULT_TIMEOUT = 20
DEFAULT_RETRY_WAIT_429 = 90  # сек — ожидание при 429 перед повторной попыткой
MAX_429_RETRIES_BEFORE_EXIT = 2  # после стольких циклов 429 — выход, чтобы избежать блокировки
CARD_DETAIL_BATCH_SIZE = 30  # карточки WB возвращают 404 при батче > ~50

# Счётчик 429 по всему запуску — при превышении лимита скрипт завершается
_rate_limit_429_count: List[int] = [0]

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/131.0.0.0 Safari/537.36"
)


def load_wb_token_from_keys() -> str:
    """Load only an explicitly configured environment value; never read a shared key file."""
    token = os.environ.get("WB_TOKEN", "").strip()
    if not token:
        raise ValueError("Set WB_TOKEN before using the optional live WB lookup")
    if not token.lower().startswith("bearer "):
        token = f"Bearer {token}"
    return token


def get_all_active_nmids(session: requests.Session) -> List[int]:
    """
    Получает все активные nmID магазина через Content API v2 с пагинацией.
    """
    result: List[int] = []
    cursor_data: Optional[Dict[str, Any]] = None

    while True:
        payload: Dict[str, Any] = {
            "settings": {
                "sort": {"ascending": False},
                "filter": {"withPhoto": -1},
                "cursor": {"limit": 100},
            }
        }

        if cursor_data:
            next_cursor = payload["settings"]["cursor"]
            if cursor_data.get("updatedAt"):
                next_cursor["updatedAt"] = cursor_data["updatedAt"]
            if cursor_data.get("nmID") is not None:
                next_cursor["nmID"] = cursor_data["nmID"]

        data = request_json(
            session,
            "POST",
            CONTENT_API_URL,
            params={"locale": "ru"},
            json_body=payload,
        )

        cards = data.get("cards", [])
        for card in cards:
            try:
                nm_id = int(card.get("nmID"))
                result.append(nm_id)
            except (TypeError, ValueError):
                continue

        cursor_data = data.get("cursor", {})
        total = cursor_data.get("total", 0)
        returned = len(cards)

        if returned < 100 or total <= len(result):
            break

        sleep_jitter(0.5)

    return result


@dataclass
class OwnProduct:
    nm_id: int
    subject_id: int
    subject_name: str
    title: str
    brand: str
    description: str
    country: Optional[str]


@dataclass
class Competitor:
    own_nm_id: int
    own_title: str
    own_subject_id: int
    own_country: Optional[str]

    competitor_nm_id: int
    competitor_title: str
    competitor_brand: str
    competitor_supplier: str
    competitor_subject_id: Optional[int]
    competitor_country: Optional[str]
    competitor_price_rub: Optional[float]
    competitor_rating: Optional[float]
    competitor_feedbacks: Optional[int]
    country_match: bool
    title_similarity: float
    score: float
    product_url: str
    query_used: str


def normalize_text(s: Any) -> str:
    if s is None:
        return ""
    s = str(s)
    s = html.unescape(s)
    s = s.replace("\xa0", " ")
    s = re.sub(r"[^\w\s%.,/-]+", " ", s, flags=re.UNICODE)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def normalize_country(s: Any) -> Optional[str]:
    s = normalize_text(s).lower()
    if not s:
        return None
    mapping = {
        "россия": "Россия",
        "рф": "Россия",
        "китай": "Китай",
        "турция": "Турция",
        "узбекистан": "Узбекистан",
        "киргизия": "Киргизия",
        "кыргызстан": "Киргизия",
        "казахстан": "Казахстан",
        "беларусь": "Беларусь",
        "белоруссия": "Беларусь",
        "индия": "Индия",
        "вьетнам": "Вьетнам",
        "пакистан": "Пакистан",
        "италия": "Италия",
        "германия": "Германия",
        "франция": "Франция",
        "южная корея": "Южная Корея",
        "корея": "Южная Корея",
    }
    return mapping.get(s, s[:1].upper() + s[1:])


def tokenize(s: str) -> List[str]:
    s = normalize_text(s).lower()
    tokens = re.findall(r"[a-zA-Zа-яА-Я0-9%]+", s)
    stop = {
        "для", "и", "в", "на", "с", "по", "из", "под", "над", "или", "при",
        "the", "and", "for", "with", "from", "to", "of",
    }
    return [t for t in tokens if len(t) > 1 and t not in stop]


def token_jaccard(a: str, b: str) -> float:
    sa = set(tokenize(a))
    sb = set(tokenize(b))
    if not sa or not sb:
        return 0.0
    return len(sa & sb) / len(sa | sb)


def seq_similarity(a: str, b: str) -> float:
    a = normalize_text(a).lower()
    b = normalize_text(b).lower()
    if not a or not b:
        return 0.0
    return SequenceMatcher(None, a, b).ratio()


def title_similarity(a: str, b: str) -> float:
    return 0.55 * seq_similarity(a, b) + 0.45 * token_jaccard(a, b)


def safe_float_price_u(value: Any) -> Optional[float]:
    try:
        if value is None:
            return None
        return round(float(value) / 100.0, 2)
    except Exception:
        return None


def sleep_jitter(base: float = 0.6) -> None:
    time.sleep(base + random.uniform(0.05, 0.25))


def build_session(token: Optional[str] = None) -> requests.Session:
    s = requests.Session()
    s.headers.update(
        {
            "User-Agent": USER_AGENT,
            "Accept": "application/json, text/plain, */*",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Referer": "https://www.wildberries.ru/",
        }
    )
    if token:
        s.headers["Authorization"] = token
    return s


def request_json(
    session: requests.Session,
    method: str,
    url: str,
    *,
    params: Optional[Dict[str, Any]] = None,
    json_body: Optional[Dict[str, Any]] = None,
    timeout: int = DEFAULT_TIMEOUT,
) -> Dict[str, Any]:
    resp = session.request(method, url, params=params, json=json_body, timeout=timeout)
    resp.raise_for_status()
    return resp.json()


def request_json_with_retry(
    session: requests.Session,
    method: str,
    url: str,
    *,
    params: Optional[Dict[str, Any]] = None,
    json_body: Optional[Dict[str, Any]] = None,
    timeout: int = DEFAULT_TIMEOUT,
    max_retries: int = 4,
    retry_wait: int = DEFAULT_RETRY_WAIT_429,
) -> Dict[str, Any]:
    """Выполняет запрос с повторными попытками при 429 (rate limit)."""
    last_error: Optional[Exception] = None
    for attempt in range(max_retries):
        try:
            resp = session.request(
                method, url, params=params, json=json_body, timeout=timeout
            )
            if resp.status_code == 429:
                _rate_limit_429_count[0] += 1
                if _rate_limit_429_count[0] >= MAX_429_RETRIES_BEFORE_EXIT:
                    print(
                        f"\n[СТОП] Достигнут лимит 429 ({_rate_limit_429_count[0]}). "
                        "Остановка для предотвращения блокировки.",
                        file=sys.stderr,
                        flush=True,
                    )
                    sys.exit(1)
                wait = retry_wait + (attempt * 15)
                print(
                    f"  [429] ({_rate_limit_429_count[0]}/{MAX_429_RETRIES_BEFORE_EXIT}) "
                    f"Ожидание {wait} сек перед повтором...",
                    flush=True,
                )
                time.sleep(wait)
                last_error = requests.HTTPError(
                    "429 Too Many Requests", response=resp
                )
                continue
            resp.raise_for_status()
            return resp.json()
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 429:
                _rate_limit_429_count[0] += 1
                if _rate_limit_429_count[0] >= MAX_429_RETRIES_BEFORE_EXIT:
                    print(
                        f"\n[СТОП] Достигнут лимит 429 ({_rate_limit_429_count[0]}). "
                        "Остановка для предотвращения блокировки.",
                        file=sys.stderr,
                        flush=True,
                    )
                    sys.exit(1)
                last_error = e
                wait = retry_wait + (attempt * 15)
                if attempt < max_retries - 1:
                    print(
                        f"  [429] ({_rate_limit_429_count[0]}/{MAX_429_RETRIES_BEFORE_EXIT}) "
                        f"Ожидание {wait} сек перед повтором...",
                        flush=True,
                    )
                    time.sleep(wait)
                continue
            raise
    raise last_error or RuntimeError("Исчерпаны попытки при 429")


def request_text(
    session: requests.Session,
    url: str,
    *,
    params: Optional[Dict[str, Any]] = None,
    timeout: int = DEFAULT_TIMEOUT,
) -> str:
    resp = session.get(url, params=params, timeout=timeout)
    resp.raise_for_status()
    return resp.text


def extract_own_country_from_characteristics(characteristics: List[Dict[str, Any]]) -> Optional[str]:
    for ch in characteristics or []:
        name = normalize_text(ch.get("name", "")).lower()
        if "страна" in name:
            value = ch.get("value")
            if isinstance(value, list) and value:
                return normalize_country(value[0])
            if value:
                return normalize_country(value)
    return None


def find_exact_card(cards: List[Dict[str, Any]], nm_id: int) -> Optional[Dict[str, Any]]:
    for card in cards:
        try:
            if int(card.get("nmID")) == int(nm_id):
                return card
        except Exception:
            continue
    return None


def get_own_card(session: requests.Session, nm_id: int) -> Optional[OwnProduct]:
    payload = {
        "settings": {
            "sort": {"ascending": False},
            "filter": {
                "textSearch": str(nm_id),
                "withPhoto": -1,
            },
            "cursor": {
                "limit": 100
            },
        }
    }

    data = request_json(
        session,
        "POST",
        CONTENT_API_URL,
        params={"locale": "ru"},
        json_body=payload,
    )

    cards = data.get("cards", [])
    card = find_exact_card(cards, nm_id)
    if not card:
        return None

    characteristics = card.get("characteristics", []) or []
    return OwnProduct(
        nm_id=int(card.get("nmID")),
        subject_id=int(card.get("subjectID")),
        subject_name=normalize_text(card.get("subjectName")),
        title=normalize_text(card.get("title")),
        brand=normalize_text(card.get("brand")),
        description=normalize_text(card.get("description")),
        country=extract_own_country_from_characteristics(characteristics),
    )


def build_search_query(own: OwnProduct) -> str:
    # Берем title без бренда, первые 6-8 токенов
    title = own.title
    if own.brand:
        title = re.sub(re.escape(own.brand), " ", title, flags=re.IGNORECASE)

    tokens = tokenize(title)
    if not tokens:
        tokens = tokenize(own.subject_name)

    # для поиска лучше не делать запрос слишком длинным
    query = " ".join(tokens[:8]).strip()
    if not query:
        query = own.subject_name or own.title or str(own.nm_id)
    return query


def parse_search_products(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    # Иногда products лежит в data.products, иногда в products
    if isinstance(data.get("data"), dict) and isinstance(data["data"].get("products"), list):
        return data["data"]["products"]
    if isinstance(data.get("products"), list):
        return data["products"]
    return []


def search_public_candidates(
    session: requests.Session,
    query: str,
    *,
    dest: str,
    pages: int,
    page_size: int,
    search_cache: Optional[Dict[Tuple[str, int], List[Dict[str, Any]]]] = None,
) -> List[Dict[str, Any]]:
    cache = search_cache if search_cache is not None else {}
    all_products: List[Dict[str, Any]] = []
    seen_ids = set()

    for page in range(1, pages + 1):
        cache_key = (query, page)
        if cache_key in cache:
            page_products = cache[cache_key]
        else:
            params = {
                "appType": 1,
                "curr": "rub",
                "dest": dest,
                "lang": "ru",
                "page": page,
                "query": query,
                "resultset": "catalog",
                "sort": "popular",
                "spp": 30,
                "suppressSpellcheck": "false",
            }

            page_products = None
            last_error = None

            for endpoint in SEARCH_ENDPOINTS:
                try:
                    data = request_json_with_retry(
                        session, "GET", endpoint, params=params
                    )
                    products = parse_search_products(data)
                    if products:
                        page_products = products
                        break
                    page_products = []
                    break
                except Exception as e:
                    last_error = e
                    continue

            if page_products is None:
                raise RuntimeError(
                    f"Не удалось выполнить поиск WB для query='{query}': {last_error}"
                )

            cache[cache_key] = page_products

        for p in page_products:
            pid = p.get("id")
            if pid in seen_ids:
                continue
            seen_ids.add(pid)
            all_products.append(p)

        if len(page_products) < page_size:
            break

        sleep_jitter(0.6)

    return all_products


def parse_detail_products(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    if isinstance(data.get("data"), dict) and isinstance(data["data"].get("products"), list):
        return data["data"]["products"]
    if isinstance(data.get("products"), list):
        return data["products"]
    return []


def chunks(items: List[int], size: int) -> Iterable[List[int]]:
    for i in range(0, len(items), size):
        yield items[i:i + size]


def _fetch_card_batch(
    session: requests.Session,
    nm_ids: List[int],
    dest: str,
) -> Dict[int, Dict[str, Any]]:
    params = {
        "appType": 0,
        "curr": "rub",
        "dest": dest,
        "spp": 30,
        "nm": ";".join(str(x) for x in nm_ids),
    }
    data = request_json_with_retry(
        session, "GET", DETAIL_API_URL, params=params
    )
    result: Dict[int, Dict[str, Any]] = {}
    for p in parse_detail_products(data):
        try:
            result[int(p["id"])] = p
        except Exception:
            continue
    return result


def get_public_details_batch(
    session: requests.Session,
    nm_ids: List[int],
    *,
    dest: str,
) -> Dict[int, Dict[str, Any]]:
    result: Dict[int, Dict[str, Any]] = {}

    for batch in chunks(nm_ids, CARD_DETAIL_BATCH_SIZE):
        try:
            batch_result = _fetch_card_batch(session, batch, dest)
            result.update(batch_result)
        except requests.HTTPError as e:
            if e.response is not None and e.response.status_code == 404:
                pass
            else:
                raise
        sleep_jitter(0.5)

    return result


COUNTRY_PATTERNS = [
    # HTML блоки
    re.compile(r"Страна\s*производства.*?</span>\s*<span[^>]*>([^<]{2,80})<", re.IGNORECASE | re.DOTALL),
    re.compile(r"Страна\s*производства.*?</dt>\s*<dd[^>]*>([^<]{2,80})<", re.IGNORECASE | re.DOTALL),
    re.compile(r"Страна\s*производства.*?aria-hidden=\"true\">([^<]{2,80})<", re.IGNORECASE | re.DOTALL),
    # JSON-подобные куски
    re.compile(r'"country(?:OfOrigin|Origin|Production|Name)?"\s*:\s*"([^"]{2,80})"', re.IGNORECASE),
    re.compile(r'"Страна производства"\s*[:,]\s*"([^"]{2,80})"', re.IGNORECASE),
]


def try_extract_country_from_product_html(session: requests.Session, nm_id: int) -> Optional[str]:
    url = f"https://www.wildberries.ru/catalog/{nm_id}/detail.aspx"
    try:
        text = request_text(session, url, timeout=DEFAULT_TIMEOUT)
    except Exception:
        return None

    text = html.unescape(text)

    for pattern in COUNTRY_PATTERNS:
        m = pattern.search(text)
        if m:
            candidate = normalize_country(m.group(1))
            if candidate:
                return candidate

    return None


def compute_score(own: OwnProduct, cand_title: str, cand_country: Optional[str]) -> Tuple[float, float, bool]:
    sim = title_similarity(own.title, cand_title)
    country_match = False
    country_bonus = 0.0

    own_country = normalize_country(own.country)
    cand_country = normalize_country(cand_country)

    if own_country and cand_country and own_country == cand_country:
        country_match = True
        country_bonus = 0.25

    score = min(1.0, 0.75 * sim + country_bonus)
    return score, sim, country_match


def build_competitors_for_one(
    session_public: requests.Session,
    own: OwnProduct,
    *,
    dest: str,
    top_k: int,
    pages: int,
    page_size: int,
    search_cache: Optional[Dict[Tuple[str, int], List[Dict[str, Any]]]] = None,
) -> List[Competitor]:
    query = build_search_query(own)
    search_products = search_public_candidates(
        session_public,
        query=query,
        dest=dest,
        pages=pages,
        page_size=page_size,
        search_cache=search_cache,
    )

    # Фильтр по subjectId и исключаем собственный товар
    filtered_ids: List[int] = []
    raw_search_map: Dict[int, Dict[str, Any]] = {}

    for p in search_products:
        try:
            pid = int(p.get("id"))
        except Exception:
            continue

        if pid == own.nm_id:
            continue

        subject_id = p.get("subjectId")
        if subject_id is None:
            continue

        try:
            if int(subject_id) != int(own.subject_id):
                continue
        except Exception:
            continue

        raw_search_map[pid] = p
        filtered_ids.append(pid)

    # Ограничим число карточек, по которым тянем detail/html
    filtered_ids = filtered_ids[:25]
    if not filtered_ids:
        return []

    details_map = get_public_details_batch(session_public, filtered_ids, dest=dest)

    rows: List[Competitor] = []
    for pid in filtered_ids:
        src = raw_search_map.get(pid, {})
        det = details_map.get(pid, src)

        title = normalize_text(det.get("name") or src.get("name"))
        brand = normalize_text(det.get("brand") or src.get("brand"))
        supplier = normalize_text(det.get("supplier") or src.get("supplier"))
        subject_id = det.get("subjectId", src.get("subjectId"))
        price_rub = safe_float_price_u(det.get("salePriceU", src.get("salePriceU")))
        rating = det.get("reviewRating", det.get("rating", src.get("reviewRating", src.get("rating"))))
        feedbacks = det.get("feedbacks", src.get("feedbacks"))

        cand_country = try_extract_country_from_product_html(session_public, pid)
        score, sim, country_match = compute_score(own, title, cand_country)

        row = Competitor(
            own_nm_id=own.nm_id,
            own_title=own.title,
            own_subject_id=own.subject_id,
            own_country=own.country,

            competitor_nm_id=pid,
            competitor_title=title,
            competitor_brand=brand,
            competitor_supplier=supplier,
            competitor_subject_id=int(subject_id) if subject_id is not None else None,
            competitor_country=cand_country,
            competitor_price_rub=price_rub,
            competitor_rating=float(rating) if rating is not None else None,
            competitor_feedbacks=int(feedbacks) if feedbacks is not None else None,
            country_match=country_match,
            title_similarity=round(sim, 4),
            score=round(score, 4),
            product_url=f"https://www.wildberries.ru/catalog/{pid}/detail.aspx",
            query_used=query,
        )
        rows.append(row)
        sleep_jitter(0.25)

    rows.sort(
        key=lambda x: (
            x.country_match,
            x.score,
            x.competitor_feedbacks or 0,
            x.competitor_rating or 0.0,
        ),
        reverse=True,
    )

    # дедуп по nmID, top_k
    final_rows: List[Competitor] = []
    seen = set()
    for row in rows:
        if row.competitor_nm_id in seen:
            continue
        seen.add(row.competitor_nm_id)
        final_rows.append(row)
        if len(final_rows) >= top_k:
            break

    return final_rows


def read_input_nmids(path: str) -> List[int]:
    result: List[int] = []
    with open(path, "r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        for row in reader:
            value = row.get("nmID") or row.get("nm_id") or row.get("nmid")
            if not value:
                continue
            try:
                result.append(int(str(value).strip()))
            except ValueError:
                continue
    return result


def write_output(path: str, rows: List[Competitor]) -> None:
    fieldnames = list(asdict(rows[0]).keys()) if rows else [
        "own_nm_id",
        "own_title",
        "own_subject_id",
        "own_country",
        "competitor_nm_id",
        "competitor_title",
        "competitor_brand",
        "competitor_supplier",
        "competitor_subject_id",
        "competitor_country",
        "competitor_price_rub",
        "competitor_rating",
        "competitor_feedbacks",
        "country_match",
        "title_similarity",
        "score",
        "product_url",
        "query_used",
    ]

    with open(path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(asdict(row))


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="input.csv", help="CSV со столбцом nmID (игнорируется при --fetch-all)")
    parser.add_argument("--output", default="result.csv", help="Куда сохранить результат")
    parser.add_argument("--fetch-all", action="store_true", help="Выгрузить все активные nmID из API вместо чтения input.csv")
    parser.add_argument("--top", type=int, default=DEFAULT_TOP, help="Сколько конкурентов вернуть на 1 товар")
    parser.add_argument("--dest", default=DEFAULT_DEST, help="Регион WB, влияет на выдачу")
    parser.add_argument("--pages", type=int, default=DEFAULT_SEARCH_PAGES, help="Сколько страниц поиска смотреть")
    parser.add_argument("--page-size", type=int, default=DEFAULT_SEARCH_PAGE_SIZE, help="Ожидаемый размер страницы")
    args = parser.parse_args()
    _rate_limit_429_count[0] = 0

    try:
        token = load_wb_token_from_keys()
    except (FileNotFoundError, ValueError) as e:
        print(f"Ошибка загрузки токена: {e}", file=sys.stderr)
        return 2

    session_official = build_session(token=token)
    session_public = build_session()

    if args.fetch_all:
        print("Выгрузка всех активных nmID из WB API...")
        try:
            nm_ids = get_all_active_nmids(session_official)
        except Exception as e:
            print(f"Ошибка выгрузки nmID из API: {e}", file=sys.stderr)
            return 2
        print(f"Загружено nmID: {len(nm_ids)}")
    else:
        nm_ids = read_input_nmids(args.input)
        if not nm_ids:
            print("Ошибка: во входном CSV не найдено ни одного nmID", file=sys.stderr)
            return 2

    all_rows: List[Competitor] = []
    search_cache: Dict[Tuple[str, int], List[Dict[str, Any]]] = {}

    for idx, nm_id in enumerate(nm_ids, start=1):
        print(f"[{idx}/{len(nm_ids)}] nmID={nm_id}")

        try:
            own = get_own_card(session_official, nm_id)
        except requests.HTTPError as e:
            print(f"  ! Ошибка official API для {nm_id}: {e}", file=sys.stderr)
            sleep_jitter(0.8)
            continue
        except Exception as e:
            print(f"  ! Ошибка получения своей карточки {nm_id}: {e}", file=sys.stderr)
            sleep_jitter(0.8)
            continue

        if not own:
            print(f"  ! Карточка {nm_id} не найдена в вашем кабинете")
            sleep_jitter(0.8)
            continue

        print(
            f"  subjectID={own.subject_id}, title={own.title[:70]!r}, country={own.country!r}"
        )

        try:
            competitors = build_competitors_for_one(
                session_public,
                own,
                dest=args.dest,
                top_k=args.top,
                pages=args.pages,
                page_size=args.page_size,
                search_cache=search_cache,
            )
        except Exception as e:
            print(f"  ! Ошибка поиска конкурентов для {nm_id}: {e}", file=sys.stderr)
            sleep_jitter(0.8)
            continue

        if not competitors:
            print("  - кандидаты не найдены")
        else:
            for rank, c in enumerate(competitors, start=1):
                print(
                    f"  {rank}. {c.competitor_nm_id} | {c.competitor_title[:60]!r} "
                    f"| country={c.competitor_country!r} | score={c.score}"
                )
            all_rows.extend(competitors)

        # Content API и поиск WB имеют лимиты (429), поэтому пауза между товарами
        sleep_jitter(1.2)

    write_output(args.output, all_rows)
    print(f"\nГотово. Сохранено строк: {len(all_rows)} -> {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())