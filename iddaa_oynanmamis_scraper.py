import pandas as pd
import cloudscraper
from bs4 import BeautifulSoup
from datetime import datetime, timedelta, timezone
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import json
import logging
import os
import unicodedata
from pathlib import Path

import requests
from data_store import replace_upcoming_matches

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    print("Not: 'tqdm' kütüphanesi yüklü değil. 'pip install tqdm' ile yükleyebilirsiniz.")
    HAS_TQDM = False

scraper = cloudscraper.create_scraper(
    browser={'browser': 'chrome', 'platform': 'windows', 'mobile': False}
)


def _build_logger():
    """Ag ve parse sorunlarini arayuzden bagimsiz olarak kaydeder."""
    logger = logging.getLogger("iddaa_oynanmamis_scraper")
    if logger.handlers:
        return logger

    log_dir = Path(__file__).resolve().parent / "logs"
    log_dir.mkdir(exist_ok=True)
    handler = logging.FileHandler(log_dir / "scraper.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    logger.propagate = False
    return logger


LOGGER = _build_logger()

BASE_HEADERS = {
    "Referer": "https://www.mackolik.com/",
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
    "Accept": "application/json, text/plain, */*",
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/140.0 Safari/537.36",
}

# ────────────────────────────────────────────────
# Market tanımları
# ────────────────────────────────────────────────
HEDEF_MARKETLER = {
    "Maç Sonucu":           ["1", "X", "2"],
    "2,5 Alt/Üst":          ["Alt", "Üst"],
    "Karşılıklı Gol":       ["Var", "Yok"],
    "Çifte Şans":           ["1-X", "1-2", "X-2"],
    "1. Yarı 0,5 Alt/Üst":  ["Alt", "Üst"],
    "1. Yarı 1,5 Alt/Üst":  ["Alt", "Üst"],
}

PERFORM_MARKET_MAP = {
    "1": "Maç Sonucu",
    "2": "2,5 Alt/Üst",
    "3": "Karşılıklı Gol",
    "4": "Çifte Şans",
    "5": "1. Yarı 0,5 Alt/Üst",
    "6": "1. Yarı 1,5 Alt/Üst",
}

# mackolik'in "başlamamış" maç için kullandığı olası status değerleri
BASLAMAMISLIK_STATULERI = {
    "notstarted", "ns", "scheduled", "fixture",
    "tbd", "postponed", "not started", "0"
}

DEBUG = False


# ────────────────────────────────────────────────
# Yardımcı: HTTP GET
# ────────────────────────────────────────────────
def safe_get(url, params=None, headers=None, max_retry=3):
    h = {**BASE_HEADERS, **(headers or {})}
    for attempt in range(max_retry):
        try:
            resp = scraper.get(url, params=params, headers=h, timeout=15)
            if resp.status_code == 200:
                return resp
            LOGGER.warning("HTTP %s: %s", resp.status_code, url)
            if DEBUG:
                print(f"  [HTTP {resp.status_code}] {url}")
        except Exception as e:
            LOGGER.warning("HTTP istegi basarisiz: %s (%s)", url, e)
            if DEBUG:
                print(f"  [HATA] {e}")
        time.sleep(0.4 * (attempt + 1))
    return None


# ────────────────────────────────────────────────
# Maçın başlamamış olup olmadığını kontrol et
# ────────────────────────────────────────────────
def is_baslamamislik(md, simdi_utc_ms):
    """
    True döndürürse maç henüz başlamamış demektir.
    Üç farklı kontrol yapar (herhangi biri yeterliyse kabul eder):
      1. status alanı → "notstarted", "ns" vb.
      2. Maç saati → şu andan sonraysa
      3. Skor tamamen boş/sıfırsa
    """
    # 1. Status kontrolü
    status = str(md.get("status", md.get("matchStatus", ""))).lower().strip()
    if status in BASLAMAMISLIK_STATULERI:
        return True

    # 2. Zaman kontrolü – maç saati şu andan ilerideyse
    mac_utc_ms = md.get("mstUtc") or md.get("startTimeUtc")
    if mac_utc_ms and isinstance(mac_utc_ms, (int, float)):
        if mac_utc_ms > simdi_utc_ms:
            return True

    # 3. Skor boşsa (henüz gol atılmamış + maç ilerlememiş)
    #    Not: 0-0 maçlar live olabilir, bu yüzden sadece
    #    status null/boş VE skor null ise kabul et
    sc = md.get("score") or {}
    home = sc.get("home")
    away = sc.get("away")
    if home is None and away is None and not status:
        return True

    return False


# ────────────────────────────────────────────────
# 1. Aşama: Başlamamış maç listesi çek
# ────────────────────────────────────────────────
def fetch_day_upcoming(tarih_obj):
    tarih_str = tarih_obj.strftime("%Y-%m-%d")
    url = "https://www.mackolik.com/perform/p0/ajax/components/competition/livescores/json"
    resp = safe_get(url, params={"sports[]": "Soccer", "matchDate": tarih_str})

    simdi_utc_ms = datetime.now(timezone.utc).timestamp() * 1000
    day_matches  = []

    if resp:
        try:
            data  = resp.json().get("data", {})
            matches = data.get("matches", {})
            comps   = data.get("competitions", {})

            for m_id, md in matches.items():
                # İddaa kodu yoksa atla
                if not md.get("iddaaCode"):
                    continue

                # Yalnızca başlamamış maçları al
                if not is_baslamamislik(md, simdi_utc_ms):
                    continue

                comp_id = str(md.get("competitionId"))
                lig_adi = comps.get(comp_id, {}).get("name", "Bilinmeyen Lig")

                day_matches.append({
                    "ID":        m_id,
                    "Lig":       lig_adi,
                    "Tarih_Raw": md.get("mstUtc"),
                    "Maç":       md.get("matchName", ""),
                })
        except Exception as e:
            LOGGER.warning("Mac listesi parse hatasi (%s): %s", tarih_str, e)
            if DEBUG:
                print(f"  [Maç listesi parse hatası – {tarih_str}] {e}")

    return day_matches


def fetch_match_result(match_id, date_text):
    """Mackolik gün sonuçlarından bitmiş maç skorunu döndürür; belirsizse None."""
    try:
        date_obj = datetime.strptime(str(date_text), "%d-%m-%Y")
    except ValueError:
        return None
    resp = safe_get(
        "https://www.mackolik.com/perform/p0/ajax/components/competition/livescores/json",
        params={"sports[]": "Soccer", "matchDate": date_obj.strftime("%Y-%m-%d")},
    )
    if not resp:
        return None
    try:
        matches = resp.json().get("data", {}).get("matches", {})
        match = matches.get(str(match_id)) or matches.get(match_id)
        if not match:
            return None
        status = str(match.get("status", match.get("matchStatus", ""))).lower()
        score = match.get("score") or {}
        home, away = score.get("home"), score.get("away")
        if home is None or away is None or status in BASLAMAMISLIK_STATULERI or status in {"live", "inplay", "1h", "2h", "ht"}:
            return None
        return int(home), int(away)
    except (ValueError, TypeError, AttributeError):
        return None


# ────────────────────────────────────────────────
# 2a. Yöntem 1 – Perform JSON API
# ────────────────────────────────────────────────
PERFORM_ODDS_URL = "https://www.mackolik.com/perform/p0/ajax/components/match/iddaa/json"

def get_odds_via_api(match_id):
    resp = safe_get(PERFORM_ODDS_URL, params={"matchId": match_id})
    if not resp:
        return None
    try:
        payload  = resp.json()
        odds_raw = payload.get("data", {}).get("odds", {})
        if not odds_raw:
            return None

        result = {}
        for market_id, market_data in odds_raw.items():
            market_name = PERFORM_MARKET_MAP.get(str(market_id))
            if not market_name:
                continue
            expected = HEDEF_MARKETLER[market_name]
            for outcome in market_data.get("outcomes", []):
                lbl = outcome.get("name", "").strip()
                val = outcome.get("odd", "")
                if lbl in expected and val and val != "-":
                    result[f"{market_name} ({lbl})"] = str(val).replace(".", ",")

        return result if result else None
    except Exception as e:
        LOGGER.warning("Oran API parse hatasi (%s): %s", match_id, e)
        if DEBUG:
            print(f"  [API parse hatası – {match_id}] {e}")
        return None


# ────────────────────────────────────────────────
# 2b. Yöntem 2 – maç sayfası HTML parse
# ────────────────────────────────────────────────
def _normalise_market_text(value):
    """Compare source market names despite punctuation and Turkish characters."""
    text = unicodedata.normalize("NFKD", str(value or "")).encode("ascii", "ignore").decode().casefold()
    text = text.replace(",", ".").replace("gol ", "")
    return " ".join(text.split())


def _target_market_name(title):
    normalised = _normalise_market_text(title)
    for target in HEDEF_MARKETLER:
        if normalised.startswith(_normalise_market_text(target)):
            return target
    return None


def _extract_odds_from_market_html(html):
    """Extract odds only when both market heading and outcome label agree."""
    soup = BeautifulSoup(html, "html.parser")
    result = {}
    blocks = soup.select(".widget-iddaa-markets__market")

    def is_source_market(block):
        item = block.find_parent("li", class_="widget-iddaa-markets__market-item")
        if not item:
            return False
        market_id, market_number = item.get("data-market-id"), item.get("data-market-no")
        # Mackolik's visible source set has a distinct source market number;
        # its partner mirror repeats the same id in both attributes.
        return bool(market_id and market_number and market_id != market_number)

    # Some pages include two rates for a single market. Prefer the source set
    # consistently for every market; use the mirror only if the source set has
    # no value for that market.
    for source_only in (True, False):
        for block in blocks:
            if source_only != is_source_market(block):
                continue
            header = block.select_one(".widget-iddaa-markets__header-text") or block.find(["h2", "h3"])
            target = _target_market_name(header.get_text(" ", strip=True) if header else "")
            if not target:
                continue
            expected = set(HEDEF_MARKETLER[target])
            option_list = next((item for item in block.select("ul.widget-iddaa-markets__option-list")
                                if "hidden" not in (item.get("class") or [])), None)
            if not option_list:
                continue
            for option in option_list.select("li.widget-iddaa-markets__option"):
                label_node = option.select_one(".widget-iddaa-markets__label")
                value_node = option.select_one(".widget-iddaa-markets__value")
                label = label_node.get_text(" ", strip=True) if label_node else ""
                value = value_node.get_text(" ", strip=True) if value_node else ""
                if label in expected and value and value != "-":
                    result.setdefault(f"{target} ({label})", value.replace(".", ","))
    return result


def get_odds_via_html(row):
    url = f"https://www.mackolik.com/mac/x/iddaa/{row['ID']}"
    headers = {**BASE_HEADERS, "Accept": "text/html,application/xhtml+xml,*/*"}
    try:
        response = requests.get(url, headers=headers, timeout=20)
        html = response.text if response.status_code == 200 else ""
    except requests.RequestException:
        html = ""
    if not html:
        response = safe_get(url, headers=headers)
        html = response.text if response else ""
    if not html:
        return None
    if DEBUG:
        os.makedirs("debug_html", exist_ok=True)
        with open(f"debug_html/{row['ID']}.html", "w", encoding="utf-8") as f:
            f.write(html)
    try:
        result = _extract_odds_from_market_html(html)
        return result if result else None
    except Exception as exc:
        LOGGER.warning("HTML oran parse hatasi (%s): %s", row["ID"], exc)
        return None


def _extract_from_json_obj(obj, depth=0):
    if depth > 8 or not isinstance(obj, (dict, list)):
        return {}
    result = {}
    if isinstance(obj, list):
        for item in obj:
            result.update(_extract_from_json_obj(item, depth + 1))
    else:
        name = obj.get("name", "")
        odd  = obj.get("odd") or obj.get("value") or obj.get("price")
        if name and odd:
            for market_name, labels in HEDEF_MARKETLER.items():
                if name in labels:
                    result[f"_raw_{name}"] = str(odd).replace(".", ",")
        for v in obj.values():
            if isinstance(v, (dict, list)):
                result.update(_extract_from_json_obj(v, depth + 1))
    return result


# ────────────────────────────────────────────────
# Ana oran çekme (API → HTML fallback)
# ────────────────────────────────────────────────
def get_odds_for_match(row):
    base = {
        "_mackolik_id": row["ID"],
        "Tarih": row["Tarih"],
        "Saat":  row["Saat"],
        "Lig":   row["Lig"],
        "Maç":   row["Maç"],
    }

    # The legacy JSON endpoint returns 404 intermittently and exposes only
    # unstable numeric market ids. The match page includes the market name and
    # outcome label, so it is the authoritative source.
    odds = get_odds_via_html(row)
    if not odds:
        return None

    return {**base, **odds}


# ────────────────────────────────────────────────
# Ana fonksiyon
# ────────────────────────────────────────────────
def fetch_today_unplayed_odds(max_workers=4):
    """Bugunun baslamamis maclarini ve mevcut oranlarini DataFrame olarak dondurur.

    Tahmin arayuzu bu fonksiyonu kullanir. ``main`` Excel ciktili komut satiri
    akisini korurken, arayuzun veriyi bellekte isleyebilmesi icin ayni scraper
    sonucunu tablo olarak sunar.
    """
    today = datetime.now().astimezone().date()
    matches = fetch_day_upcoming(today)
    if not matches:
        LOGGER.info("Bugun icin baslamamis iddaali mac bulunamadi.")
        return pd.DataFrame()

    matches_df = pd.DataFrame(matches)
    timestamps = pd.to_datetime(matches_df["Tarih_Raw"], unit="ms", utc=True)
    matches_df["Tarih"] = timestamps.dt.tz_convert("Europe/Istanbul").dt.strftime("%d-%m-%Y")
    matches_df["Saat"] = timestamps.dt.tz_convert("Europe/Istanbul").dt.strftime("%H:%M")
    rows = matches_df.sort_values("Tarih_Raw").to_dict("records")

    final_rows = []
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [executor.submit(get_odds_for_match, row) for row in rows]
        for future in as_completed(futures):
            try:
                row = future.result()
            except Exception:
                row = None
            if row:
                final_rows.append(row)

    if not final_rows:
        LOGGER.warning("%s mac bulundu ancak hicbirinin orani cekilemedi.", len(rows))
        return pd.DataFrame()

    result = pd.DataFrame(final_rows)
    base_columns = ["Tarih", "Saat", "Lig", "Maç"]
    other_columns = sorted(column for column in result.columns if column not in base_columns)
    return result[base_columns + other_columns]


def main(basla, son, cikti="excel/iddaa_fiktur.xlsx", database_path="data/iddaa.sqlite"):
    print(f"\n{'='*50}\nİDDAA FİKSTÜR ORAN SİSTEMİ BAŞLATILDI\n{'='*50}")

    start_t  = datetime.strptime(basla, "%Y-%m-%d")
    end_t    = datetime.strptime(son,   "%Y-%m-%d")
    tarihler = [start_t + timedelta(days=x) for x in range((end_t - start_t).days + 1)]

    # ── Aşama 1: Başlamamış Maç Listesi ──
    print(f"[*] Aşama 1: {len(tarihler)} günlük fikstür taranıyor...")
    all_basics = []
    with ThreadPoolExecutor(max_workers=8) as ex:
        results = list(ex.map(fetch_day_upcoming, tarihler))
    for r in results:
        all_basics.extend(r)

    if not all_basics:
        print("[!] HATA: Belirtilen tarihlerde başlamamış iddaalı maç bulunamadı.")
        print("    → Tarihlerin gelecekte olduğundan emin olun.")
        return

    df_m = pd.DataFrame(all_basics)
    df_m["Tarih_Raw"] = pd.to_datetime(df_m["Tarih_Raw"], unit="ms", utc=True).dt.tz_convert("Europe/Istanbul")
    df_m["Saat"]      = df_m["Tarih_Raw"].dt.strftime("%H:%M")
    df_m["Tarih"]     = df_m["Tarih_Raw"].dt.strftime("%d-%m-%Y")

    # Tarihe ve saate göre sırala
    df_m = df_m.sort_values(["Tarih_Raw"]).reset_index(drop=True)

    match_dicts   = df_m.to_dict("records")
    total_matches = len(match_dicts)
    print(f"[+] Toplam {total_matches} başlamamış maç bulundu. Oranlar çekilecek.\n")

    # ── Aşama 2: Oranlar ──
    final_data = []

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(get_odds_for_match, m): m for m in match_dicts}

        if HAS_TQDM:
            pbar = tqdm(total=total_matches, desc="İŞLENEN", unit="maç", colour="cyan")

        for future in as_completed(futures):
            try:
                res = future.result()
                if res:
                    final_data.append(res)
            except Exception:
                pass
            finally:
                if HAS_TQDM:
                    pbar.update(1)

        if HAS_TQDM:
            pbar.close()

    # ── Aşama 3: Excel ──
    if final_data:
        print(f"\n[*] Aşama 3: Veriler Excel'e aktarılıyor...")
        final_df = pd.DataFrame(final_data)

        # Sütun sırası: önce kimlik bilgileri, sonra marketler
        base_cols  = ["Tarih", "Saat", "Lig", "Maç"]
        other_cols = sorted([c for c in final_df.columns if c not in base_cols])

        # Tarihe ve saate göre sırala
        final_df = final_df.sort_values(["Tarih", "Saat"]).reset_index(drop=True)
        final_df[base_cols + other_cols].to_excel(cikti, index=False)
        if database_path:
            replace_upcoming_matches(final_df.to_dict("records"), database_path)

        print(f"{'='*50}")
        print(f"BAŞARILI: {len(final_df)} maç '{cikti}' dosyasına yazıldı.")
        print(f"Fikstür SQLite deposuna güncellendi: {database_path}")
        print(f"{'='*50}")
    else:
        print("\n[!] HATA: Hiçbir maçın oranı çekilemedi.")
        print("    → DEBUG için 'DEBUG = True' yapıp tekrar deneyin.")


# ────────────────────────────────────────────────
if __name__ == "__main__":
    bugun = datetime.now().strftime("%Y-%m-%d")
    cikti = "excel/iddaa_analiz.xlsx"
    main(bugun, bugun, cikti)
