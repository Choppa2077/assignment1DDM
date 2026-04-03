import requests
from bs4 import BeautifulSoup
import time
import sys
import re
from datetime import datetime, timedelta
import random

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
}
DELAY = 2.0  # seconds between requests

CITY_NAMES = {
    'astana': 'Астана',
    'almaty': 'Алматы',
    'shymkent': 'Шымкент',
    'karaganda': 'Карагандa',
    'aktobe': 'Актобе',
}

_listing_id_counter = [0]

def next_id():
    _listing_id_counter[0] += 1
    return _listing_id_counter[0]


def parse_price(text):
    """Extract integer price from text like '220 000 ₸/мес'"""
    if not text:
        return None
    digits = re.sub(r'[^\d]', '', text)
    return int(digits) if digits else None


def parse_olx_date(text):
    """Convert OLX date like '01 апреля 2026 г.' or 'Сегодня в 12:44' to YYYY-MM-DD"""
    if not text:
        return None
    text = text.strip()

    if 'сегодня' in text.lower():
        return datetime.now().strftime('%Y-%m-%d')
    if 'вчера' in text.lower():
        return (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')

    MONTH_MAP = {
        'январ': 1, 'феврал': 2, 'март': 3, 'апрел': 4,
        'мая': 5, 'май': 5, 'июн': 6, 'июл': 7, 'август': 8,
        'сентябр': 9, 'октябр': 10, 'ноябр': 11, 'декабр': 12,
    }
    m = re.match(r'(\d+)\s+([а-яА-Я]+)\s+(\d{4})', text)
    if m:
        day = int(m.group(1))
        month_str = m.group(2).lower()
        year = int(m.group(3))
        month = None
        for key, val in MONTH_MAP.items():
            if month_str.startswith(key):
                month = val
                break
        if month:
            try:
                return datetime(year, month, day).strftime('%Y-%m-%d')
            except ValueError:
                pass
    return None


def parse_rooms_from_title(title):
    """
    Extract rooms from free-form title (OLX style).
    Handles: '1 комнатная', 'однокомнатную', '3к', '1 ком.', 'студия'
    """
    if not title:
        return None
    t = title.lower()

    # Word-based: однокомнатная, двухкомнатная, etc.
    word_map = {
        r'однокомнат': 1,
        r'двухкомнат|2\s*-?\s*х?\s*комнат': 2,
        r'тр[её]хкомнат|3\s*-?\s*х?\s*комнат': 3,
        r'четыр[её]хкомнат|4\s*-?\s*х?\s*комнат': 4,
        r'пятикомнат|5\s*-?\s*х?\s*комнат': 5,
    }
    for pattern, val in word_map.items():
        if re.search(pattern, t):
            return val

    # Digit-based: "1 комнатная", "2-комнатная"
    m = re.search(r'(\d+)\s*-?\s*комнат', t)
    if m:
        return int(m.group(1))

    # Short forms: "3к", "1 ком."
    m = re.search(r'(\d+)\s*к\b', t)
    if m:
        return int(m.group(1))
    m = re.search(r'(\d+)\s*ком\.', t)
    if m:
        return int(m.group(1))

    if re.search(r'студи', t):
        return 0

    return None


def parse_rooms_area_floor(title):
    """
    Extract rooms, area, floor, total_floors from title like:
    '4-комнатная квартира · 95 м² · 6/9 этаж'
    'Студия · 32 м² · 3/12 этаж'
    """
    rooms = None
    area = None
    floor = None
    total_floors = None

    # Rooms
    m = re.search(r'(\d+)-комнатн', title, re.IGNORECASE)
    if m:
        rooms = int(m.group(1))
    elif re.search(r'студия', title, re.IGNORECASE):
        rooms = 0

    # Area
    m = re.search(r'([\d.,]+)\s*м²', title)
    if m:
        area = float(m.group(1).replace(',', '.'))

    # Floor / total_floors
    m = re.search(r'(\d+)/(\d+)\s*этаж', title)
    if m:
        floor = int(m.group(1))
        total_floors = int(m.group(2))

    return rooms, area, floor, total_floors


def parse_date(text):
    """Convert Krisha date like '19 мар.' to YYYY-MM-DD"""
    MONTH_MAP = {
        'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4,
        'май': 5, 'июн': 6, 'июл': 7, 'авг': 8,
        'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12,
    }
    if not text:
        return None
    text = text.strip().lower().rstrip('.')
    m = re.match(r'(\d+)\s+([а-я]+)', text)
    if m:
        day = int(m.group(1))
        month_str = m.group(2)[:3]
        month = MONTH_MAP.get(month_str)
        if month:
            year = datetime.now().year
            # If month is in the future, it's from last year
            if month > datetime.now().month:
                year -= 1
            try:
                return datetime(year, month, day).strftime('%Y-%m-%d')
            except ValueError:
                pass
    if 'сегодня' in text:
        return datetime.now().strftime('%Y-%m-%d')
    if 'вчера' in text:
        return (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
    return None


def parse_seller_type(card):
    """Extract seller type from card"""
    owner_el = card.select_one('.a-card__owner')
    if not owner_el:
        return 'unknown'
    text = owner_el.get_text(strip=True).lower()
    if 'специалист' in text or 'агент' in text or 'агентств' in text:
        return 'agency'
    elif 'частное' in text or 'владелец' in text or 'собственник' in text or 'хозяин' in text:
        return 'owner'
    return 'unknown'


def parse_krisha_page(soup, city_slug):
    """Parse one page of Krisha.kz listings"""
    results = []
    city_name = CITY_NAMES.get(city_slug, city_slug)

    cards = soup.select('div.a-card')

    for card in cards:
        item = {}
        item['source'] = 'krisha'
        item['city'] = city_name

        # listing_id from data-id attribute
        item['listing_id'] = card.get('data-id', '')

        # Title contains rooms + area + floor info: "4-комнатная квартира · 95 м² · 6/9 этаж"
        title_el = card.select_one('a.a-card__title')
        title = title_el.get_text(strip=True) if title_el else ''
        item['title'] = title

        rooms, area, floor, total_floors = parse_rooms_area_floor(title)
        item['rooms'] = rooms
        item['area_m2'] = area
        item['floor'] = floor
        item['total_floors'] = total_floors

        # Price: in .a-card__price, ignore child spans (currency etc)
        price_el = card.select_one('.a-card__price')
        if price_el:
            # Get only the direct text (price number), not spans
            price_text = price_el.get_text(strip=True)
            item['price_tenge'] = parse_price(price_text)
        else:
            item['price_tenge'] = None

        # District / address: in .a-card__subtitle
        subtitle_el = card.select_one('.a-card__subtitle')
        subtitle = subtitle_el.get_text(strip=True) if subtitle_el else ''
        # Format: "Есильский р-н, Туркестан — Алматы"
        # District is first part before comma
        if subtitle:
            parts = subtitle.split(',')
            item['district'] = parts[0].strip()
            item['address'] = subtitle
        else:
            item['district'] = None
            item['address'] = ''

        # Date: in .a-card__stats-item (second one)
        stats_items = card.select('.a-card__stats-item')
        listing_date = None
        for stat in stats_items:
            text = stat.get_text(strip=True)
            if re.search(r'\d+\s+[а-я]+', text, re.IGNORECASE) or 'сегодня' in text.lower() or 'вчера' in text.lower():
                listing_date = parse_date(text)
                break
        item['listing_date'] = listing_date or datetime.now().strftime('%Y-%m-%d')

        # Seller type
        item['seller_type'] = parse_seller_type(card)

        # URL
        link_el = card.select_one('a.a-card__title') or card.select_one('a[href*="/a/show/"]')
        if link_el:
            href = link_el.get('href', '')
            item['url'] = f"https://krisha.kz{href}" if href.startswith('/') else href
        else:
            item['url'] = ''

        # price_per_m2
        if item['price_tenge'] and item['area_m2'] and item['area_m2'] > 0:
            item['price_per_m2'] = round(item['price_tenge'] / item['area_m2'], 1)
        else:
            item['price_per_m2'] = None

        if item['price_tenge']:
            results.append(item)

    return results


def parse_krisha(city_slug: str, max_pages: int = 50) -> list:
    """
    Parse rental listings from Krisha.kz for given city.
    city_slug: 'astana', 'almaty', 'shymkent', 'karaganda', 'aktobe'
    Returns list of dicts with dataset fields.
    If site is unavailable (status != 200) — calls sys.exit() with error message.
    """
    base_url = f"https://krisha.kz/arenda/kvartiry/{city_slug}/"
    all_listings = []

    # Test first page
    print(f"  Krisha.kz / {city_slug}: connecting...")
    try:
        resp = requests.get(base_url, headers=HEADERS, timeout=15)
        print(f"  Status: {resp.status_code}")
        if resp.status_code != 200:
            print(f"ERROR: Krisha.kz returned status {resp.status_code} for {city_slug}. Stopping.")
            sys.exit(1)
    except requests.RequestException as e:
        print(f"ERROR: Cannot connect to Krisha.kz ({city_slug}): {e}")
        sys.exit(1)

    for page in range(1, max_pages + 1):
        url = base_url if page == 1 else f"{base_url}?page={page}"
        try:
            if page > 1:
                time.sleep(DELAY)
            resp = requests.get(url, headers=HEADERS, timeout=15)
            if resp.status_code != 200:
                print(f"  Page {page}: status {resp.status_code}, stopping city")
                break

            soup = BeautifulSoup(resp.text, 'lxml')
            page_listings = parse_krisha_page(soup, city_slug)

            if not page_listings:
                print(f"  Page {page}: no listings found, stopping city")
                break

            all_listings.extend(page_listings)
            print(f"  Page {page}: +{len(page_listings)} listings (total: {len(all_listings)})")

            # Check if there's a next page
            next_btn = (soup.select_one('a[data-marker="pagination-button/nextPage"]') or
                        soup.select_one('.paginator__page--next') or
                        soup.select_one('[rel="next"]') or
                        soup.select_one('a.paginator__page[class*="next"]'))

            # Also check via page number: if current page is the last
            paginator = soup.select('.paginator__page')
            if paginator:
                page_nums = []
                for p in paginator:
                    try:
                        num = int(p.get_text(strip=True))
                        page_nums.append(num)
                    except ValueError:
                        pass
                if page_nums and page >= max(page_nums):
                    print(f"  Reached last page ({page}), stopping city")
                    break

        except requests.RequestException as e:
            print(f"  Page {page}: request error {e}, stopping city")
            break

    return all_listings


def parse_olx_page(soup, city_slug):
    """Parse one page of OLX.kz listings"""
    results = []
    city_name = CITY_NAMES.get(city_slug, city_slug)

    cards = (soup.select('[data-cy="l-card"]') or
             soup.select('.offer-wrapper') or
             soup.select('article[data-id]'))

    for card in cards:
        item = {}
        item['source'] = 'olx'
        item['city'] = city_name
        item['listing_id'] = card.get('id', '') or card.get('data-id', '')

        # Title — OLX renders title in img alt and in <a> links
        img_el = card.select_one('img[alt]')
        if img_el and img_el.get('alt', '').strip():
            title = img_el['alt'].strip()
        else:
            # Fallback: second <a> tag usually has the title text
            links = card.select('a[href]')
            title = ''
            for link in links:
                text = link.get_text(strip=True)
                if len(text) > 5:
                    title = text
                    break
        item['title'] = title

        # Rooms from free-form OLX title; area/floor usually not available
        item['rooms'] = parse_rooms_from_title(title)
        item['area_m2'] = None
        item['floor'] = None
        item['total_floors'] = None

        # Price
        price_el = (card.select_one('[data-testid="ad-price"]') or
                    card.select_one('.price') or
                    card.select_one('[class*="price"]'))
        if price_el:
            price_text = price_el.get_text(strip=True)
            item['price_tenge'] = parse_price(price_text)
        else:
            item['price_tenge'] = None

        # Location — format: "Астана, Нура  - 01 апреля 2026 г."
        loc_el = (card.select_one('[data-testid="location-date"]') or
                  card.select_one('[class*="location"]'))
        loc_text = loc_el.get_text(strip=True) if loc_el else ''
        # Split by " - " to separate location from date
        loc_parts = loc_text.split(' - ')
        location = loc_parts[0].strip() if loc_parts else ''
        date_part = loc_parts[1].strip() if len(loc_parts) > 1 else ''

        # District: "Астана, Есильский район" → district = "Есильский район"
        addr_parts = location.split(',')
        if len(addr_parts) >= 2:
            item['district'] = addr_parts[1].strip()
        else:
            item['district'] = addr_parts[0].strip() if addr_parts else None
        item['address'] = location

        # Date from location string
        item['listing_date'] = parse_olx_date(date_part) or datetime.now().strftime('%Y-%m-%d')

        item['seller_type'] = 'unknown'

        # URL
        link_el = card.select_one('a[href]')
        if link_el:
            href = link_el.get('href', '')
            item['url'] = f"https://www.olx.kz{href}" if href.startswith('/') else href
        else:
            item['url'] = ''

        if item['price_tenge'] and item['area_m2'] and item['area_m2'] > 0:
            item['price_per_m2'] = round(item['price_tenge'] / item['area_m2'], 1)
        else:
            item['price_per_m2'] = None

        if item['price_tenge']:
            results.append(item)

    return results


def parse_olx(city_slug: str, max_pages: int = 20) -> list:
    """
    Parse rental listings from OLX.kz for given city.
    If site is unavailable — calls sys.exit() with error message.
    """
    olx_city_map = {
        'astana': 'astana',
        'almaty': 'alm',
    }
    olx_city = olx_city_map.get(city_slug, city_slug)
    base_url = f"https://www.olx.kz/nedvizhimost/arenda-kvartiry/{olx_city}/"
    all_listings = []

    print(f"  OLX.kz / {city_slug}: connecting...")
    try:
        resp = requests.get(base_url, headers=HEADERS, timeout=15)
        print(f"  Status: {resp.status_code}")
        if resp.status_code != 200:
            print(f"WARNING: OLX.kz returned status {resp.status_code} for {city_slug}. Skipping OLX.")
            return []
    except requests.RequestException as e:
        print(f"WARNING: Cannot connect to OLX.kz ({city_slug}): {e}. Skipping OLX.")
        return []

    for page in range(1, max_pages + 1):
        url = base_url if page == 1 else f"{base_url}?page={page}"
        try:
            if page > 1:
                time.sleep(DELAY)
            resp = requests.get(url, headers=HEADERS, timeout=15)
            if resp.status_code != 200:
                break

            soup = BeautifulSoup(resp.text, 'lxml')
            page_listings = parse_olx_page(soup, city_slug)

            if not page_listings:
                break

            all_listings.extend(page_listings)
            print(f"  Page {page}: +{len(page_listings)} listings (total: {len(all_listings)})")

        except requests.RequestException as e:
            print(f"  OLX page {page}: error {e}")
            break

    return all_listings


def main():
    all_listings = []

    cities_krisha = ['astana', 'almaty', 'shymkent', 'karaganda', 'aktobe']
    for city in cities_krisha:
        print(f"\n=== Krisha.kz / {city} ===")
        listings = parse_krisha(city, max_pages=50)
        all_listings.extend(listings)
        print(f"Krisha / {city}: {len(listings)} записей")

    cities_olx = ['astana', 'almaty']
    for city in cities_olx:
        print(f"\n=== OLX.kz / {city} ===")
        listings = parse_olx(city, max_pages=20)
        all_listings.extend(listings)
        print(f"OLX / {city}: {len(listings)} записей")

    # Number IDs sequentially
    for i, item in enumerate(all_listings, 1):
        item['id'] = i

    print(f"\nИтого: {len(all_listings)} записей")

    if len(all_listings) < 6000:
        print(f"ВНИМАНИЕ: собрано только {len(all_listings)} записей, требуется 6000+")
        print("Попробуй увеличить max_pages или добавить города")

    return all_listings


if __name__ == "__main__":
    data = main()
