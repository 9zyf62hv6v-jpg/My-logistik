# geocode.py — геокодирование адресов (DaData + Nominatim)
# v4: ключ читается из secret.py (приоритет) или из переменной окружения
import os
import json
import time
import requests

# ------------------------------------------------------------
# Чтение ключа: сначала secret.py, потом переменная окружения
# ------------------------------------------------------------
DADATA_API_KEY = ""

# 1) пробуем secret.py
try:
    from secret import DADATA_API_KEY as _key_from_secret
    if _key_from_secret and str(_key_from_secret).strip():
        DADATA_API_KEY = str(_key_from_secret).strip()
        print("[geocode] Ключ DaData загружен из secret.py")
except ImportError:
    pass
except Exception as e:
    print(f"[geocode] secret.py есть, но не читается: {e}")

# 2) если в secret.py пусто — берём из переменной окружения
if not DADATA_API_KEY:
    DADATA_API_KEY = os.environ.get("DADATA_API_KEY", "").strip()
    if DADATA_API_KEY:
        print("[geocode] Ключ DaData загружен из переменной окружения")

# 3) если совсем ничего — предупреждаем
if not DADATA_API_KEY:
    print("[geocode] ⚠️ Ключ DaData не найден. Создай secret.py рядом с app.py")

# ------------------------------------------------------------
# Константы
# ------------------------------------------------------------
DADATA_SUGGEST_URL = "https://suggestions.dadata.ru/suggestions/api/4_1/rs/suggest/address"
NOMINATIM_URL = "https://nominatim.openstreetmap.org/search"
USER_AGENT = "MyLogistics/1.0 (contact: logistics@moylogistika.ru)"

CITY_MARKERS = [
    'волжский', 'волгоград', 'волгоградская',
    'котово', 'ерзовка', 'средняя ахтуба', 'краснослободск',
]
DEFAULT_CITY = 'Волжский'


def _normalize_query(address, city=None):
    address = (address or "").strip()
    if not address:
        return ""
    if city and city.lower() not in address.lower():
        return f"{city}, {address}"
    return address


def _has_city_marker(address):
    a = (address or "").lower()
    return any(m in a for m in CITY_MARKERS)


def dadata_suggest(address, city=None, count=5, debug=False):
    if not DADATA_API_KEY:
        print("[geocode] DADATA_API_KEY не задан")
        return []

    address = (address or "").strip()
    if not address:
        return []

    if city:
        queries = [_normalize_query(address, city)]
    elif _has_city_marker(address):
        queries = [_normalize_query(address, None)]
    else:
        queries = [f"{DEFAULT_CITY}, {address}", address]

    headers = {
        "Content-Type": "application/json",
        "Accept": "application/json",
        "Authorization": f"Token {DADATA_API_KEY}",
    }

    for q in queries:
        payload = {"query": q, "count": count}
        if debug:
            print(f"[geocode] → POST payload: {json.dumps(payload, ensure_ascii=False)}")
        try:
            r = requests.post(DADATA_SUGGEST_URL, json=payload,
                              headers=headers, timeout=10)
            if debug:
                print(f"[geocode]   HTTP {r.status_code}")
            r.raise_for_status()
            data = r.json()
        except Exception as e:
            print(f"[geocode] DaData ошибка: {e}")
            continue

        suggestions = data.get("suggestions", [])
        if debug:
            print(f"[geocode] ← suggestions: {len(suggestions)}")
            for s in suggestions[:3]:
                d = s.get("data") or {}
                print(f"[geocode]   • {s.get('value')}  ({d.get('geo_lat')}, {d.get('geo_lon')})")

        results = []
        for s in suggestions:
            d = s.get("data") or {}
            lat = d.get("geo_lat")
            lon = d.get("geo_lon")
            if lat and lon:
                results.append({
                    "address": s.get("value") or q,
                    "lat": float(lat),
                    "lon": float(lon),
                    "source": "dadata",
                })
        if results:
            return results

    return []


def nominatim_search(address, city=None):
    query = _normalize_query(address, city)
    if not query:
        return None
    params = {
        "q": query, "format": "json", "limit": 1,
        "countrycodes": "ru", "addressdetails": 1,
    }
    try:
        r = requests.get(NOMINATIM_URL, params=params,
                         headers={"User-Agent": USER_AGENT}, timeout=10)
        if r.status_code == 403:
            return None
        r.raise_for_status()
        data = r.json()
    except Exception as e:
        print(f"[geocode] Nominatim ошибка: {e}")
        return None
    if not data:
        return None
    item = data[0]
    return {
        "address": item.get("display_name") or query,
        "lat": float(item["lat"]),
        "lon": float(item["lon"]),
        "source": "nominatim",
    }


def geocode_address(address, city=None, debug=False):
    address = (address or "").strip()
    if not address:
        return None
    found = dadata_suggest(address, city=city, count=1, debug=debug)
    if found:
        return found[0]
    time.sleep(1.1)
    return nominatim_search(address, city=city)


def geocode_many(items, city=None, pause=0.2, debug=False):
    out = []
    for cid, addr in items:
        res = geocode_address(addr, city=city, debug=debug)
        if res:
            res["id"] = cid
            out.append(res)
        time.sleep(pause)
    return out


# ============================================================
# ТЕСТ: python geocode.py
# ============================================================
if __name__ == "__main__":
    print("=" * 70)
    print("ТЕСТ ГЕОКОДЕРА v4")
    print("=" * 70)

    if not DADATA_API_KEY:
        print("❌ Ключ не найден. Создай secret.py рядом с app.py:")
        print('   DADATA_API_KEY = "твой_ключ"')
    else:
        print(f"✅ Ключ DaData: {DADATA_API_KEY[:8]}...")

    test_addresses = [
        "пр-кт им Ленина, д 21",
        "г Волгоград, ул им. маршала Воронова, д. 22а",
        "Волгоградская область г котово",
        "Ерзовка, СНТ Вымпел, уч. 142ш",
    ]
    for addr in test_addresses:
        print("-" * 70)
        print(f"АДРЕС: {addr}")
        res = geocode_address(addr, city=None, debug=False)
        if res:
            print(f"✅ {res['lat']}, {res['lon']} ({res['source']})")
            print(f"   {res['address']}")
        else:
            print("❌ НЕ НАЙДЕНО")
    print("=" * 70)