import pandas as pd
import cloudscraper
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import argparse
import json
import os
import re
from pathlib import Path

from data_store import clear_matches, load_matches, save_matches

try:
    from tqdm import tqdm
    HAS_TQDM = True
except ImportError:
    print("Not: 'tqdm' kütüphanesi yüklü değil. 'pip install tqdm' ile yükleyebilirsiniz.")
    HAS_TQDM = False

scraper = cloudscraper.create_scraper(
    browser={'browser': 'chrome', 'platform': 'windows', 'mobile': False}
)

BASE_HEADERS = {
    "Referer": "https://www.mackolik.com/",
    "Accept-Language": "tr-TR,tr;q=0.9,en;q=0.8",
    "Accept": "application/json, text/plain, */*",
}

# ────────────────────────────────────────────────
# Market tanımları  (API key → sütun adı, seçenek listesi)
# ────────────────────────────────────────────────
HEDEF_MARKETLER = {
    "Maç Sonucu":           ["1", "X", "2"],
    "2,5 Alt/Üst":          ["Alt", "Üst"],
    "Karşılıklı Gol":       ["Var", "Yok"],
    "Çifte Şans":           ["1-X", "1-2", "X-2"],
    "1. Yarı 0,5 Alt/Üst":  ["Alt", "Üst"],
    "1. Yarı 1,5 Alt/Üst":  ["Alt", "Üst"],
}

# mackolik perform API'deki market id → isim eşleştirmesi
# (Tarayıcı geliştirici araçları / Network sekmesiyle bulunabilir)
PERFORM_MARKET_MAP = {
    "1":   "Maç Sonucu",      # Maç Sonucu
    "2":   "2,5 Alt/Üst",     # 2.5 Alt/Üst
    "3":   "Karşılıklı Gol",  # KG
    "4":   "Çifte Şans",
    "5":   "1. Yarı 0,5 Alt/Üst",
    "6":   "1. Yarı 1,5 Alt/Üst",
}

DEBUG = False          # --debug bayrağıyla açılır, HTML kaydeder
BASE_COLS = ["Tarih", "Saat", "Lig", "Maç", "İY Skor", "Skor"]
MATCH_KEY_COLS = ["Tarih", "Saat", "Lig", "Maç"]


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
            if DEBUG:
                print(f"  [HTTP {resp.status_code}] {url}")
        except Exception as e:
            if DEBUG:
                print(f"  [HATA] {e}")
        time.sleep(0.4 * (attempt + 1))
    return None


# ────────────────────────────────────────────────
# 1. Aşama: Maç listesi çek
# ────────────────────────────────────────────────
def fetch_day_matches(tarih_obj):
    tarih_str = tarih_obj.strftime("%Y-%m-%d")
    url = "https://www.mackolik.com/perform/p0/ajax/components/competition/livescores/json"
    resp = safe_get(url, params={"sports[]": "Soccer", "matchDate": tarih_str})
    day_matches = []
    if resp:
        try:
            data = resp.json().get("data", {})
            matches = data.get("matches", {})
            comps   = data.get("competitions", {})
            for m_id, md in matches.items():
                if not isinstance(md, dict):
                    continue
                if not md.get("iddaaCode"):
                    continue
                comp_id  = str(md.get("competitionId"))
                lig_adi  = comps.get(comp_id, {}).get("name", "Bilinmeyen Lig")
                sc = md.get("score", {}) or {}
                ht = sc.get("ht") or {}
                day_matches.append({
                    "ID":       m_id,
                    "Lig":      lig_adi,
                    "Tarih_Raw": md.get("mstUtc"),
                    "Maç":      md.get("matchName", ""),
                    "İY Skor":  f"{ht.get('home','')}-{ht.get('away','')}",
                    "Skor":     f"{sc.get('home','')}-{sc.get('away','')}",
                })
        except Exception as e:
            if DEBUG:
                print(f"  [Maç listesi parse hatası] {e}")
    return day_matches


# ────────────────────────────────────────────────
# 2a. Yöntem 1 – Perform JSON API (tercih edilen)
# ────────────────────────────────────────────────
PERFORM_ODDS_URL = "https://www.mackolik.com/perform/p0/ajax/components/match/iddaa/json"

def get_odds_via_api(match_id):
    """
    mackolik'in kendi AJAX endpoint'ini dener.
    Başarılı olursa dict, değilse None döner.
    """
    resp = safe_get(PERFORM_ODDS_URL, params={"matchId": match_id})
    if not resp:
        return None
    try:
        payload = resp.json()
        odds_raw = payload.get("data", {}).get("odds", {})
        if not odds_raw:
            return None

        result = _extract_odds_from_payload(odds_raw)
        return result if result else None
    except Exception as e:
        if DEBUG:
            print(f"  [API parse hatası – {match_id}] {e}")
        return None


# ────────────────────────────────────────────────
# 2b. Yöntem 2 – HTML parse (yedek)
#     Sorun: mackolik sayfası JS ile render ediliyor;
#     BeautifulSoup statik HTML alır → oranlar çoğu zaman eksik/yanlış.
#     Bu yüzden __json__ script tag içindeki gömülü veriyi önce dener,
#     sonra DOM'a bakar.
# ────────────────────────────────────────────────
def get_odds_via_html(row):
    url  = f"https://www.mackolik.com/mac/x/iddaa/{row['ID']}"
    resp = safe_get(url, headers={"Accept": "text/html,application/xhtml+xml,*/*"})
    if not resp:
        return None

    html = resp.text

    if DEBUG:
        os.makedirs("debug_html", exist_ok=True)
        with open(f"debug_html/{row['ID']}.html", "w", encoding="utf-8") as f:
            f.write(html)

    result = {}

    # ── Yöntem 2a: Sayfaya gömülü JSON (window.__data__ veya benzeri) ──
    try:
        # Yaygın SSR pattern: __NUXT__= veya window.__data__= veya "odds":{...}
        patterns = [
            r'window\.__(?:NUXT|data|state)__\s*=\s*(\{.*?\});',
            r'"iddaaOdds"\s*:\s*(\{[^}]+\})',
            r'"markets"\s*:\s*(\[[^\]]+\])',
        ]
        for pat in patterns:
            m = re.search(pat, html, re.DOTALL)
            if m:
                chunk = m.group(1)
                try:
                    obj = json.loads(chunk)
                    # İçerisinde market verisi aramak için özyinelemeli tarama
                    embedded = _extract_from_json_obj(obj)
                    if embedded:
                        result.update(embedded)
                        break
                except Exception:
                    pass
    except Exception:
        pass

    # ── Yöntem 2b: BeautifulSoup DOM parse ──
    if not result:
        try:
            soup = BeautifulSoup(html, "html.parser")

            # Önce class adı farklı olabilir mi diye tüm alternatif seçicileri dene
            selectors = [
                ("div", "widget-iddaa-markets__market"),
                ("div", "iddaa-market"),
                ("div", "market-block"),
                ("section", "market"),
            ]

            market_blocks = []
            for tag, cls in selectors:
                blocks = soup.find_all(tag, class_=cls)
                if blocks:
                    market_blocks = blocks
                    break

            if not market_blocks:
                # Son çare: h2 içeren tüm bölümleri tara
                market_blocks = []
                for h2 in soup.find_all("h2"):
                    parent = h2.find_parent(["div", "section", "article"])
                    if parent:
                        market_blocks.append(parent)

            for block in market_blocks:
                header = block.find(["h2", "h3", "span"])
                if not header:
                    continue
                raw_title = header.get_text(strip=True)

                target_key = None
                for key in HEDEF_MARKETLER:
                    if raw_title.startswith(key):
                        target_key = key
                        break
                if not target_key:
                    continue

                expected = HEDEF_MARKETLER[target_key]

                for item in block.find_all("li"):
                    spans = item.find_all("span")
                    if len(spans) < 2:
                        continue

                    # İki olası sıralama: [label, değer] veya [değer, label]
                    lbl_a, val_a = spans[0].get_text(strip=True), spans[1].get_text(strip=True)
                    lbl_b, val_b = spans[-1].get_text(strip=True), spans[-2].get_text(strip=True)

                    # Hangi sıralamada label expected listesindeyse onu kullan
                    if lbl_a in expected and val_a not in expected and val_a != "-":
                        result[f"{target_key} ({lbl_a})"] = val_a.replace(".", ",")
                    elif lbl_b in expected and val_b not in expected and val_b != "-":
                        result[f"{target_key} ({lbl_b})"] = val_b.replace(".", ",")

        except Exception as e:
            if DEBUG:
                print(f"  [HTML parse hatası – {row['ID']}] {e}")

    return result if result else None


def _market_name(value):
    """API/HTML verisindeki market kimliğini hedef market ismine çevirir."""
    if value is None:
        return None
    text = str(value).strip()
    if text in PERFORM_MARKET_MAP:
        return PERFORM_MARKET_MAP[text]
    for market in HEDEF_MARKETLER:
        if text.casefold() == market.casefold():
            return market
    return None


def _add_odd(result, market_name, label, value):
    """Yalnızca marketi ve seçeneği kesin olarak bilinen oranı ekler."""
    if not market_name or market_name not in HEDEF_MARKETLER:
        return
    label = str(label or "").strip()
    value = str(value or "").strip()
    if label in HEDEF_MARKETLER[market_name] and value and value != "-":
        result[f"{market_name} ({label})"] = value.replace(".", ",")


def _extract_odds_from_payload(odds_raw):
    """Farklı API şemalarındaki market/seçim listelerinden güvenli oran çıkarır."""
    result = {}
    markets = odds_raw.items() if isinstance(odds_raw, dict) else enumerate(odds_raw or [])
    for market_id, market_data in markets:
        if not isinstance(market_data, dict):
            continue
        market_name = (
            _market_name(market_id)
            or _market_name(market_data.get("id"))
            or _market_name(market_data.get("marketId"))
            or _market_name(market_data.get("name"))
            or _market_name(market_data.get("title"))
        )
        outcomes = market_data.get("outcomes") or market_data.get("selections") or market_data.get("options") or []
        if isinstance(outcomes, dict):
            outcomes = outcomes.values()
        for outcome in outcomes:
            if not isinstance(outcome, dict):
                continue
            _add_odd(
                result,
                market_name,
                outcome.get("name") or outcome.get("label") or outcome.get("title"),
                outcome.get("odd") or outcome.get("value") or outcome.get("price"),
            )
    return result


def _extract_from_json_obj(obj, market_name=None, depth=0):
    """Gömülü JSON'dan, market bağlamını koruyarak doğrulanmış oranları çıkarır."""
    if depth > 8 or not isinstance(obj, (dict, list)):
        return {}
    result = {}
    if isinstance(obj, list):
        for item in obj:
            result.update(_extract_from_json_obj(item, market_name, depth + 1))
    else:
        has_outcome_group = any(key in obj for key in ("outcomes", "selections", "options"))
        current_market = (
            _market_name(obj.get("marketId"))
            or _market_name(obj.get("market_id"))
            or _market_name(obj.get("id") if has_outcome_group else None)
            or _market_name(obj.get("marketName"))
            or _market_name(obj.get("title") if has_outcome_group else None)
            or _market_name(obj.get("name") if has_outcome_group else None)
            or market_name
        )
        _add_odd(
            result,
            current_market,
            obj.get("name") or obj.get("label"),
            obj.get("odd") or obj.get("value") or obj.get("price"),
        )
        for v in obj.values():
            if isinstance(v, (dict, list)):
                result.update(_extract_from_json_obj(v, current_market, depth + 1))
    return result


def _match_key(row):
    return tuple(str(row.get(column, "")).strip() for column in MATCH_KEY_COLS)


def _load_existing_output(output_path):
    if not output_path.exists():
        return pd.DataFrame(), set()
    existing = load_matches(output_path) if output_path.suffix.lower() == ".sqlite" else pd.read_excel(output_path)
    if existing.empty:
        return existing, set()
    missing = [column for column in MATCH_KEY_COLS if column not in existing.columns]
    if missing:
        raise ValueError(f"Mevcut çıktı dosyasında zorunlu kolon eksik: {', '.join(missing)}")
    existing = existing.drop_duplicates(subset=MATCH_KEY_COLS, keep="last").reset_index(drop=True)
    return existing, {_match_key(row) for row in existing.to_dict("records")}


def _write_output(rows, output_path):
    if not rows:
        return 0
    if output_path.suffix.lower() == ".sqlite":
        return save_matches(rows, output_path)
    final_df = pd.DataFrame(rows)
    base_cols = [column for column in BASE_COLS if column in final_df.columns]
    other_cols = sorted(column for column in final_df.columns if column not in base_cols)
    final_df = final_df.drop_duplicates(subset=MATCH_KEY_COLS, keep="last")
    final_df = final_df.sort_values(["Tarih", "Saat", "Lig", "Maç"]).reset_index(drop=True)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    final_df[base_cols + other_cols].to_excel(output_path, index=False)
    return len(final_df)


# ────────────────────────────────────────────────
# 2. Ana oran çekme fonksiyonu (API → HTML fallback)
# ────────────────────────────────────────────────
def get_odds_for_match(row):
    base = {
        "Tarih":   row["Tarih"],
        "Saat":    row["Saat"],
        "Lig":     row["Lig"],
        "Maç":     row["Maç"],
        "İY Skor": row["İY Skor"],
        "Skor":    row["Skor"],
    }

    # Önce JSON API dene
    odds = get_odds_via_api(row["ID"])

    # Başarısız olursa HTML parse
    if not odds:
        odds = get_odds_via_html(row)

    if not odds:
        return None

    return {**base, **odds}


# ────────────────────────────────────────────────
# DEBUG aracı: Tek bir maç için HTML kaydet
# ────────────────────────────────────────────────
def debug_single_match(match_id):
    """
    Kullanım: python iddaa_scraperclaude.py --debug-match 123456
    Verilen maç için hem API yanıtını hem HTML'i kaydeder.
    """
    print(f"\n[DEBUG] Maç {match_id} için veri çekiliyor...")

    # API
    api_resp = safe_get(PERFORM_ODDS_URL, params={"matchId": match_id})
    if api_resp:
        with open(f"debug_{match_id}_api.json", "w", encoding="utf-8") as f:
            json.dump(api_resp.json(), f, ensure_ascii=False, indent=2)
        print(f"  → debug_{match_id}_api.json kaydedildi")
    else:
        print("  → API yanıt vermedi")

    # HTML
    html_resp = safe_get(
        f"https://www.mackolik.com/mac/x/iddaa/{match_id}",
        headers={"Accept": "text/html,application/xhtml+xml,*/*"}
    )
    if html_resp:
        with open(f"debug_{match_id}.html", "w", encoding="utf-8") as f:
            f.write(html_resp.text)
        print(f"  → debug_{match_id}.html kaydedildi")

        # Sayfada kaç tane widget-iddaa-markets__market var?
        soup = BeautifulSoup(html_resp.text, "html.parser")
        blocks = soup.find_all("div", class_="widget-iddaa-markets__market")
        print(f"  → 'widget-iddaa-markets__market' div sayısı: {len(blocks)}")
        for i, b in enumerate(blocks):
            h = b.find(["h2", "h3"])
            print(f"     Block {i}: {h.get_text(strip=True) if h else '(başlık yok)'}")
    else:
        print("  → HTML alınamadı")


# ────────────────────────────────────────────────
# Ana fonksiyon
# ────────────────────────────────────────────────
def main(basla, son, cikti="iddaa_analiz.xlsx", workers=8, checkpoint_every=250, resume=True):
    print(f"\n{'='*50}\nİDDAA ANALİZ SİSTEMİ BAŞLATILDI\n{'='*50}")
    output_path = Path(cikti)

    start_t  = datetime.strptime(basla, "%Y-%m-%d")
    end_t    = datetime.strptime(son,   "%Y-%m-%d")
    if end_t < start_t:
        raise ValueError("Bitiş tarihi başlangıç tarihinden önce olamaz.")
    tarihler = [start_t + timedelta(days=x) for x in range((end_t - start_t).days + 1)]

    # ── Aşama 1: Maç Listesi ──
    print(f"[*] Aşama 1: {len(tarihler)} günlük maç bülteni indiriliyor...")
    all_basics = []
    with ThreadPoolExecutor(max_workers=workers) as ex:
        results = list(ex.map(fetch_day_matches, tarihler))
    for r in results:
        all_basics.extend(r)

    if not all_basics:
        print("[!] HATA: Maç listesi boş döndü!")
        return

    df_m = pd.DataFrame(all_basics)
    df_m["Tarih_Raw"] = pd.to_datetime(df_m["Tarih_Raw"], unit="ms", utc=True).dt.tz_convert("Europe/Istanbul")
    df_m["Saat"]      = df_m["Tarih_Raw"].dt.strftime("%H:%M")
    df_m["Tarih"]     = df_m["Tarih_Raw"].dt.strftime("%d-%m-%Y")

    match_dicts = df_m.to_dict("records")
    total_matches = len(match_dicts)

    existing_df = pd.DataFrame()
    completed_keys = set()
    if not resume and output_path.suffix.lower() == ".sqlite":
        clear_matches(output_path)
    if resume:
        existing_df, completed_keys = _load_existing_output(output_path)
    rows_to_fetch = [match for match in match_dicts if _match_key(match) not in completed_keys]

    if resume and not existing_df.empty:
        print(f"[=] Mevcut çıktı bulundu: {len(existing_df)} maç zaten kayıtlı, bunlar atlanacak.")
    print(f"[+] Toplam {total_matches} maç bulundu. Oranlar çekilecek: {len(rows_to_fetch)}\n")

    if not rows_to_fetch:
        print("[OK] İşlenecek yeni maç yok. Çıktı dosyası zaten güncel görünüyor.")
        return

    # ── Aşama 2: Oranlar ──
    final_data = existing_df.to_dict("records") if not existing_df.empty else []
    processed_since_checkpoint = 0
    success_count = 0
    with ThreadPoolExecutor(max_workers=workers) as executor:
        futures = {executor.submit(get_odds_for_match, m): m for m in rows_to_fetch}

        if HAS_TQDM:
            pbar = tqdm(total=len(rows_to_fetch), desc="İŞLENEN", unit="maç", colour="green")

        for future in as_completed(futures):
            try:
                res = future.result()
                if res:
                    final_data.append(res)
                    completed_keys.add(_match_key(res))
                    processed_since_checkpoint += 1
                    success_count += 1
            except Exception as exc:
                if DEBUG:
                    match = futures[future]
                    print(f"  [Oran çekme hatası] {match.get('ID')} {match.get('Maç', '')}: {exc}")
            finally:
                if HAS_TQDM:
                    pbar.update(1)

            if checkpoint_every and processed_since_checkpoint >= checkpoint_every:
                saved_rows = _write_output(final_data, output_path)
                print(f"\n[ARA KAYIT] {saved_rows} maç '{output_path}' dosyasına yazıldı.")
                processed_since_checkpoint = 0

        if HAS_TQDM:
            pbar.close()

    # ── Aşama 3: Excel ──
    if final_data:
        print(f"\n[*] Aşama 3: Veriler Excel'e aktarılıyor...")
        saved_rows = _write_output(final_data, output_path)
        print(f"{'='*50}\nBAŞARILI: {success_count} yeni maç çekildi, toplam {saved_rows} maç '{output_path}' dosyasına yazıldı.\n{'='*50}")
    else:
        print("\n[!] HATA: Hiçbir maçın oranı çekilemedi.")
        print("    → '--debug-match <ID>' ile tek bir maçı test edin.")


# ────────────────────────────────────────────────
def ask_date(prompt):
    """Geçerli bir YYYY-MM-DD tarihi girilene kadar sorar."""
    while True:
        val = input(prompt).strip()
        try:
            datetime.strptime(val, "%Y-%m-%d")
            return val
        except ValueError:
            print("  [!] Geçersiz format. Lütfen YYYY-MM-DD şeklinde girin (örn: 2026-09-01)")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Mackolik'ten geçmiş iddaa oranlarını Excel'e aktarır.")
    parser.add_argument("basla", nargs="?", help="Başlangıç tarihi (YYYY-MM-DD)")
    parser.add_argument("son", nargs="?", help="Bitiş tarihi (YYYY-MM-DD)")
    parser.add_argument("-o", "--output", default="data/iddaa.sqlite", help="Çıktı deposu (.sqlite) veya Excel dosyası")
    parser.add_argument("--workers", type=int, default=8, help="Eşzamanlı istek sayısı (varsayılan: 8)")
    parser.add_argument("--checkpoint-every", type=int, default=250, help="Her N başarılı maçta ara kayıt alır (0 kapatır)")
    parser.add_argument("--fresh", action="store_true", help="Mevcut çıktı varsa devam etmek yerine sıfırdan yazar")
    parser.add_argument("--debug", action="store_true", help="HTTP ve ayrıştırma hatalarını gösterir")
    parser.add_argument("--debug-match", metavar="ID", help="Tek bir maçın API/HTML yanıtlarını kaydeder")
    args = parser.parse_args()

    DEBUG = args.debug
    if args.debug_match:
        debug_single_match(args.debug_match)
    else:
        if bool(args.basla) != bool(args.son):
            parser.error("Başlangıç ve bitiş tarihlerini birlikte verin.")
        basla = args.basla or ask_date("Başlangıç tarihi (YYYY-MM-DD): ")
        son = args.son or ask_date("Bitiş tarihi   (YYYY-MM-DD): ")
        if args.workers < 1:
            parser.error("--workers en az 1 olmalıdır.")
        if args.checkpoint_every < 0:
            parser.error("--checkpoint-every negatif olamaz.")
        try:
            main(basla, son, args.output, args.workers, args.checkpoint_every, not args.fresh)
        except ValueError as exc:
            parser.error(str(exc))
