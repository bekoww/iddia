from __future__ import annotations

from datetime import datetime
from itertools import combinations
from pathlib import Path
import unicodedata
import math

import pandas as pd

from data_store import load_matches


DATA_SCOPE_ALL = "Tum Veri"
DATA_SCOPE_2026 = "Sadece 2026"
DATA_SCOPE_365 = "Son 365 Gun"
DATA_SCOPE_180 = "Son 180 Gun"
DATA_SCOPES = [DATA_SCOPE_ALL, DATA_SCOPE_2026, DATA_SCOPE_365, DATA_SCOPE_180]
MATCH_SCOPE_SMART = "Lig/Ulke Akilli"
MATCH_SCOPE_ALL = "Genel"
MATCH_SCOPE_COUNTRY = "Ayni Ulke"
MATCH_SCOPE_LEAGUE = "Ayni Lig"
MATCH_SCOPES = [MATCH_SCOPE_SMART, MATCH_SCOPE_ALL, MATCH_SCOPE_COUNTRY, MATCH_SCOPE_LEAGUE]
TODAY_SORT_CONFIDENCE = "Guven"
TODAY_SORT_VALUE = "Deger"
TODAY_SORT_ODD = "Oran"
TODAY_SORT_QUALITY = "Kalite"
TODAY_SORT_MATCHES = "Eslesme"
TODAY_SORT_TIME = "Saat"
TODAY_SORT_OPTIONS = [
    TODAY_SORT_CONFIDENCE,
    TODAY_SORT_VALUE,
    TODAY_SORT_ODD,
    TODAY_SORT_QUALITY,
    TODAY_SORT_MATCHES,
    TODAY_SORT_TIME,
]
FIXED_TOLERANCE = 8.0
FIXED_MINIMUM_MATCHES = 5

C = {
    "bg0": "#070B14", "bg1": "#0C1220", "bg2": "#111A2C", "bg3": "#182540",
    "bg4": "#1F2E4D", "bg5": "#28395E",
    "border": "#23324F", "border2": "#2E4066",
    "text1": "#EAF1FF", "text2": "#8FA0C4", "text3": "#516086",
    "green": "#2ECC71", "green_dim": "#123C25",
    "red": "#EF5D63", "red_dim": "#3D1A1D",
    "yellow": "#F2C94C", "yellow_dim": "#3D3312",
    "blue": "#4F8EFF", "blue_dim": "#152B52",
    "cyan": "#22D3EE", "purple": "#B183FF", "purple_dim": "#2B1C4D",
    "orange": "#FFA94D", "teal": "#2BD9C9",
    "accent": "#4F8EFF",
}

COUNTRY_PATTERNS = [
    ("turkiye", "Türkiye"),
    ("trendyol", "Türkiye"),
    ("nesine", "Türkiye"),
    ("super lig", "Türkiye"),
    ("1. lig", "Türkiye"),
    ("2. lig", "Türkiye"),
    ("3. lig", "Türkiye"),
    ("premier lig", "İngiltere"),
    ("championship", "İngiltere"),
    ("fa cup", "İngiltere"),
    ("ingiltere", "İngiltere"),
    ("laliga", "İspanya"),
    ("la liga", "İspanya"),
    ("ispanya", "İspanya"),
    ("serie a", "İtalya"),
    ("serie b", "İtalya"),
    ("serie c", "İtalya"),
    ("italya", "İtalya"),
    ("bundesliga", "Almanya"),
    ("almanya", "Almanya"),
    ("ligue 1", "Fransa"),
    ("ligue 2", "Fransa"),
    ("fransa", "Fransa"),
    ("eredivisie", "Hollanda"),
    ("eerste divisie", "Hollanda"),
    ("hollanda", "Hollanda"),
    ("mls", "ABD"),
    ("liga mx", "Meksika"),
    ("ascenso mx", "Meksika"),
    ("meksika", "Meksika"),
    ("eliteserien", "Norveç"),
    ("norvec", "Norveç"),
    ("allsvenskan", "İsveç"),
    ("superettan", "İsveç"),
    ("isvec", "İsveç"),
    ("ekstraklasa", "Polonya"),
    ("polonya", "Polonya"),
    ("k-lig", "Güney Kore"),
    ("j1 ", "Japonya"),
    ("j2/", "Japonya"),
    ("japonya", "Japonya"),
    ("veikkausliiga", "Finlandiya"),
    ("ykkosliiga", "Finlandiya"),
    ("finlandiya", "Finlandiya"),
    ("1. hnl", "Hırvatistan"),
    ("hirvatistan", "Hırvatistan"),
    ("1. snl", "Slovenya"),
    ("2. snl", "Slovenya"),
    ("slovenya", "Slovenya"),
    ("nb i", "Macaristan"),
    ("macaristan", "Macaristan"),
    ("urvalsdeild", "İzlanda"),
    ("izlanda", "İzlanda"),
    ("botola", "Fas"),
    ("fas", "Fas"),
    ("libertadores", "Güney Amerika"),
    ("sudamericana", "Güney Amerika"),
    ("concacaf", "Kuzey Amerika"),
    ("afc", "Asya"),
    ("avrupa", "Avrupa"),
    ("uefa", "Avrupa"),
    ("dunya", "Dünya"),
]

ODDS_FIELDS = [
    {"key": "ms1",   "label": "MS 1",       "hint": "Ev Kazanır",  "col": "Maç Sonucu (1)",              "color": C["green"]},
    {"key": "msx",   "label": "MS X",       "hint": "Beraberlik",  "col": "Maç Sonucu (X)",              "color": C["yellow"]},
    {"key": "ms2",   "label": "MS 2",       "hint": "Dep Kazanır", "col": "Maç Sonucu (2)",              "color": C["red"]},
    {"key": "au25",  "label": "2.5 Alt",    "hint": "",            "col": "2,5 Alt/Üst (Alt)",           "color": C["blue"]},
    {"key": "uu25",  "label": "2.5 Üst",    "hint": "",            "col": "2,5 Alt/Üst (Üst)",           "color": C["cyan"]},
    {"key": "kgvar", "label": "KG Var",     "hint": "",            "col": "Karşılıklı Gol (Var)",        "color": C["purple"]},
    {"key": "kgyok", "label": "KG Yok",     "hint": "",            "col": "Karşılıklı Gol (Yok)",        "color": C["orange"]},
    {"key": "cs1x",  "label": "Çift 1-X",   "hint": "",            "col": "Çifte Şans (1-X)",            "color": C["teal"]},
    {"key": "cs12",  "label": "Çift 1-2",   "hint": "",            "col": "Çifte Şans (1-2)",            "color": C["teal"]},
    {"key": "csx2",  "label": "Çift X-2",   "hint": "",            "col": "Çifte Şans (X-2)",            "color": C["teal"]},
    {"key": "iy05a", "label": "İY 0.5 Alt", "hint": "",            "col": "1. Yarı 0,5 Alt/Üst (Alt)",   "color": C["orange"]},
    {"key": "iy05u", "label": "İY 0.5 Üst", "hint": "",            "col": "1. Yarı 0,5 Alt/Üst (Üst)",   "color": C["orange"]},
    {"key": "iy15a", "label": "İY 1.5 Alt", "hint": "",            "col": "1. Yarı 1,5 Alt/Üst (Alt)",   "color": C["purple"]},
    {"key": "iy15u", "label": "İY 1.5 Üst", "hint": "",            "col": "1. Yarı 1,5 Alt/Üst (Üst)",   "color": C["purple"]},
]
FIELD_BY_KEY = {f["key"]: f for f in ODDS_FIELDS}

PREDICTION_ODDS_COLUMNS = {
    "MS 1": "Maç Sonucu (1)",
    "MS X": "Maç Sonucu (X)",
    "MS 2": "Maç Sonucu (2)",
    "2.5 Alt": "2,5 Alt/Üst (Alt)",
    "2.5 Üst": "2,5 Alt/Üst (Üst)",
    "KG Var": "Karşılıklı Gol (Var)",
    "KG Yok": "Karşılıklı Gol (Yok)",
}
CORE_ODDS_COLUMNS = ["Maç Sonucu (1)", "Maç Sonucu (X)", "Maç Sonucu (2)"]
COUPON_RISK_PROFILES = {
    # Güven değeri kalibre edilerek en fazla %75'e çıkar. Eski %80 eşiği
    # bu nedenle hiçbir zaman kupon üretemiyordu.
    "Guvenli": {"min_confidence": 65, "max_selections": 3, "min_gap": 90},
    "Dengeli": {"min_confidence": 55, "max_selections": 5, "min_gap": 60},
    "Agresif": {"min_confidence": 50, "max_selections": 7, "min_gap": 30},
    "Ozel": {},
}
# "İncele" kararı için minimum kanıt eşiği. Bu bir bahis emri değil,
# kullanıcının incelemesi gereken model/piyasa uyumsuzluğudur.
ACTIONABLE_MIN_SAMPLES = 12
ACTIONABLE_MIN_CONFIDENCE = 55.0
PREDICTION_TIERS = ("Banko", "Riskli", "Ya Tutarsa")


def to_float(value):
    if value is None:
        return None
    s = str(value).strip().replace("%", "").replace(",", ".")
    if s in ("", "-", "nan", "None", "NaN"):
        return None
    try:
        return float(s)
    except ValueError:
        return None


def parse_score(value):
    s = str(value) if value is not None else ""
    if "-" not in s:
        return None
    try:
        h, a = s.split("-", 1)
        return int(h.strip()), int(a.strip())
    except Exception:
        return None


def normalize_text(value):
    text = str(value or "").strip().casefold()
    text = unicodedata.normalize("NFKD", text)
    text = "".join(char for char in text if not unicodedata.combining(char))
    return text.replace("ı", "i")


def infer_country_from_league(league):
    text = normalize_text(league)
    if not text:
        return ""
    for pattern, country in COUNTRY_PATTERNS:
        if pattern in text:
            return country
    return ""


def load_odds_data(path):
    path = Path(path)
    df = load_matches(path) if path.suffix.lower() == ".sqlite" else pd.read_excel(path)
    duplicate_columns = [column for column in ("Tarih", "Saat", "Lig", "Maç") if column in df.columns]
    duplicates_removed = 0
    if len(duplicate_columns) == 4:
        original_count = len(df)
        df = df.drop_duplicates(subset=duplicate_columns, keep="last").reset_index(drop=True)
        duplicates_removed = original_count - len(df)
    for f in ODDS_FIELDS:
        if f["col"] in df.columns:
            df[f["col"]] = df[f["col"]].apply(to_float)
    if "Maç" in df.columns:
        split = df["Maç"].astype(str).str.split(" vs ", n=1, expand=True)
        df["_ev"] = split[0].str.strip()
        df["_dep"] = split[1].str.strip() if split.shape[1] > 1 else ""
    else:
        df["_ev"] = ""
        df["_dep"] = ""
    country_column = next((column for column in ("Ülke", "Ulke", "Country") if column in df.columns), None)
    if country_column:
        df["_ulke"] = df[country_column].fillna("").astype(str).str.strip()
    else:
        leagues = df["Lig"] if "Lig" in df.columns else pd.Series("", index=df.index)
        df["_ulke"] = leagues.apply(infer_country_from_league)
    if "Tarih" in df.columns:
        df["_tarih_dt"] = pd.to_datetime(df["Tarih"], format="%d-%m-%Y", errors="coerce")
        df = df.loc[df["_tarih_dt"].notna()].copy()
    df["_search"] = (
        df.get("Lig", "").astype(str) + " " +
        df.get("Maç", "").astype(str) + " " +
        df.get("_ulke", "").astype(str)
    ).str.lower()
    df.attrs["duplicates_removed"] = duplicates_removed
    return df


def filter_data_by_scope(df, scope, today=None):
    if df.empty or scope == DATA_SCOPE_ALL or "_tarih_dt" not in df.columns:
        return df.copy()
    dates = pd.to_datetime(df["_tarih_dt"], dayfirst=True, errors="coerce")
    if scope == DATA_SCOPE_2026:
        return df.loc[dates.dt.year == 2026].copy()
    day_count_by_scope = {DATA_SCOPE_365: 365, DATA_SCOPE_180: 180}
    day_count = day_count_by_scope.get(scope)
    if not day_count:
        return df.copy()
    today_ts = pd.Timestamp(today or datetime.now().date())
    start = today_ts - pd.Timedelta(days=day_count)
    return df.loc[dates.notna() & (dates >= start) & (dates <= today_ts)].copy()


def filter_history_by_match_scope(df, league, scope=MATCH_SCOPE_SMART, min_rows=0):
    if df.empty or scope == MATCH_SCOPE_ALL:
        return df.copy(), MATCH_SCOPE_ALL

    target_league = str(league or "").strip()
    target_country = infer_country_from_league(target_league)

    def by_league():
        if not target_league or "Lig" not in df.columns:
            return pd.DataFrame()
        mask = df["Lig"].astype(str).map(normalize_text) == normalize_text(target_league)
        return df.loc[mask].copy()

    def by_country():
        if not target_country:
            return pd.DataFrame()
        leagues = df["Lig"] if "Lig" in df.columns else pd.Series("", index=df.index)
        countries = df["_ulke"] if "_ulke" in df.columns else leagues.apply(infer_country_from_league)
        mask = countries.astype(str).str.casefold() == target_country.casefold()
        return df.loc[mask].copy()

    if scope == MATCH_SCOPE_LEAGUE:
        filtered = by_league()
        return (filtered, f"Lig: {target_league}") if not filtered.empty else (df.copy(), MATCH_SCOPE_ALL)
    if scope == MATCH_SCOPE_COUNTRY:
        filtered = by_country()
        return (filtered, f"Ulke: {target_country}") if not filtered.empty else (df.copy(), MATCH_SCOPE_ALL)

    league_filtered = by_league()
    if len(league_filtered) >= min_rows:
        return league_filtered, f"Lig: {target_league}"
    country_filtered = by_country()
    if len(country_filtered) >= min_rows:
        return country_filtered, f"Ulke: {target_country}"
    return df.copy(), MATCH_SCOPE_ALL


def latest_data_date(df):
    if df.empty or "_tarih_dt" not in df.columns:
        return None
    dates = pd.to_datetime(df["_tarih_dt"], dayfirst=True, errors="coerce").dropna()
    if dates.empty:
        return None
    return dates.max().date()


def missing_update_range(df, today=None):
    last_date = latest_data_date(df)
    today_date = pd.Timestamp(today or datetime.now().date()).date()
    if not last_date:
        return None
    start = last_date + pd.Timedelta(days=1)
    end = today_date - pd.Timedelta(days=1)
    if start > end:
        return None
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def inclusive_day_count(start, end):
    return int((pd.Timestamp(end) - pd.Timestamp(start)).days) + 1


def data_quality_summary(df, active_df=None):
    active_df = df if active_df is None else active_df
    summary = {
        "total_rows": len(df),
        "active_rows": len(active_df),
        "duplicate_rows": 0,
        "missing_score_rows": 0,
        "missing_core_odds_rows": 0,
        "league_count": 0,
        "date_min": "",
        "date_max": "",
        "top_leagues": [],
    }
    if df.empty:
        return summary
    key_columns = ["Tarih", "Saat", "Lig", "Maç"]
    if all(column in df.columns for column in key_columns):
        summary["duplicate_rows"] = int(df.duplicated(subset=key_columns).sum())
    if "Skor" in df.columns:
        summary["missing_score_rows"] = int(df["Skor"].isna().sum() + (df["Skor"].astype(str).str.strip() == "-").sum())
    odds_columns = [column for column in CORE_ODDS_COLUMNS if column in df.columns]
    if odds_columns:
        summary["missing_core_odds_rows"] = int(df[odds_columns].isna().all(axis=1).sum())
    if "Lig" in df.columns:
        leagues = df["Lig"].dropna().astype(str).str.strip()
        leagues = leagues[leagues != ""]
        summary["league_count"] = int(leagues.nunique())
        summary["top_leagues"] = list(leagues.value_counts().head(8).items())
    dates = pd.to_datetime(df.get("_tarih_dt"), dayfirst=True, errors="coerce") if "_tarih_dt" in df else pd.Series(dtype="datetime64[ns]")
    valid_dates = dates.dropna()
    if not valid_dates.empty:
        summary["date_min"] = valid_dates.min().strftime("%d-%m-%Y")
        summary["date_max"] = valid_dates.max().strftime("%d-%m-%Y")
    return summary


def outcome_stats(df, target_league=None, weight_recent=False):
    home = draw = away = over15 = over25 = over35 = btts = n = goals = 0
    home_over05 = home_over15 = away_over05 = away_over15 = 0
    weight_total = 0.0
    ft_scores, ht_scores = {}, {}
    ft_score_counts, ht_score_counts = {}, {}
    league_key = str(target_league or "").strip().casefold()
    dates = pd.to_datetime(df.get("_tarih_dt"), errors="coerce") if "_tarih_dt" in df else pd.Series(dtype="datetime64[ns]")
    oldest_date = dates.min() if not dates.empty else pd.NaT
    newest_date = dates.max() if not dates.empty else pd.NaT
    for _, row in df.iterrows():
        p = parse_score(row.get("Skor"))
        if not p:
            continue
        h, a = p
        n += 1
        weight = 1.0
        if league_key and str(row.get("Lig", "")).strip().casefold() == league_key:
            weight *= 1.25
        row_date = row.get("_tarih_dt")
        if weight_recent and pd.notna(row_date) and pd.notna(oldest_date) and newest_date > oldest_date:
            recency = (row_date - oldest_date) / (newest_date - oldest_date)
            weight *= 0.85 + (0.30 * float(recency))
        weight_total += weight
        goals += (h + a) * weight
        if h > a:
            home += weight
        elif h < a:
            away += weight
        else:
            draw += weight
        total_goals = h + a
        if total_goals > 1:
            over15 += weight
        if total_goals > 2:
            over25 += weight
        if total_goals > 3:
            over35 += weight
        if h > 0 and a > 0:
            btts += weight
        if h > 0:
            home_over05 += weight
        if h > 1:
            home_over15 += weight
        if a > 0:
            away_over05 += weight
        if a > 1:
            away_over15 += weight
        k = f"{h}-{a}"
        ft_scores[k] = ft_scores.get(k, 0.0) + weight
        ft_score_counts[k] = ft_score_counts.get(k, 0) + 1
        hp = parse_score(row.get("İY Skor"))
        if hp:
            hk = f"{hp[0]}-{hp[1]}"
            ht_scores[hk] = ht_scores.get(hk, 0.0) + weight
            ht_score_counts[hk] = ht_score_counts.get(hk, 0) + 1
    v = weight_total or 1
    best_ft = max(ft_scores, key=ft_scores.get) if ft_scores else "—"
    best_ht = max(ht_scores, key=ht_scores.get) if ht_scores else "—"
    top_ft_scores = [
        {"score": score, "count": ft_score_counts.get(score, 0), "pct": ft_scores[score] / v * 100}
        for score in sorted(ft_scores, key=ft_scores.get, reverse=True)[:3]
    ]
    top_ht_scores = [
        {"score": score, "count": ht_score_counts.get(score, 0), "pct": ht_scores[score] / v * 100}
        for score in sorted(ht_scores, key=ht_scores.get, reverse=True)[:3]
    ]
    return {
        "n": n, "total": len(df),
        "home": home, "draw": draw, "away": away,
        "home_pct": home / v * 100, "draw_pct": draw / v * 100, "away_pct": away / v * 100,
        "over15_pct": over15 / v * 100, "under15_pct": 100 - (over15 / v * 100),
        "over25_pct": over25 / v * 100, "under25_pct": 100 - (over25 / v * 100),
        "over35_pct": over35 / v * 100, "under35_pct": 100 - (over35 / v * 100),
        "btts_pct": btts / v * 100, "btts_no_pct": 100 - (btts / v * 100),
        "home_over05_pct": home_over05 / v * 100, "home_over15_pct": home_over15 / v * 100,
        "away_over05_pct": away_over05 / v * 100, "away_over15_pct": away_over15 / v * 100,
        "avg_goals": goals / v,
        "best_ft": best_ft, "best_ft_n": ft_score_counts.get(best_ft, 0),
        "best_ht": best_ht, "best_ht_n": ht_score_counts.get(best_ht, 0),
        "top_ft_scores": top_ft_scores,
        "top_ht_scores": top_ht_scores,
    }


def filter_by_fields(df, entered, tol_pct):
    m = df
    for key, val in entered.items():
        col = FIELD_BY_KEY[key]["col"]
        if col not in df.columns:
            return df.iloc[0:0]
        lo, hi = val * (1 - tol_pct / 100.0), val * (1 + tol_pct / 100.0)
        # SQLite/Excel sources can carry odds as strings with either decimal
        # separator. Convert at comparison time so mixed source types cannot
        # make every match appear unavailable.
        values = m[col] if pd.api.types.is_numeric_dtype(m[col]) else pd.to_numeric(
            m[col].astype("string").str.replace(",", ".", regex=False),
            errors="coerce",
        )
        m = m[values.notna() & (values >= lo) & (values <= hi)]
    return m


def best_prediction(stats, row=None):
    """En güçlü geçmiş eğilimi döndürür.

    ``row`` sağlanmışsa sadece kullanıcının oranını girdiği pazarlardan seçim
    yapılır. Böylece fiyatı bilinmeyen bir takım golü pazarı, öneri gibi
    gösterilmez.
    """
    if not stats["n"]:
        return {"label": "Veri yok", "confidence": 0.0, "color": C["text3"]}
    candidates = [
        ("MS 1", stats["home_pct"], C["green"]),
        ("MS X", stats["draw_pct"], C["yellow"]),
        ("MS 2", stats["away_pct"], C["red"]),
        ("1.5 Üst", stats["over15_pct"], C["cyan"]),
        ("1.5 Alt", stats["under15_pct"], C["blue"]),
        ("2.5 Üst", stats["over25_pct"], C["cyan"]),
        ("2.5 Alt", stats["under25_pct"], C["blue"]),
        ("3.5 Üst", stats["over35_pct"], C["cyan"]),
        ("3.5 Alt", stats["under35_pct"], C["blue"]),
        ("KG Var", stats["btts_pct"], C["purple"]),
        ("KG Yok", stats["btts_no_pct"], C["orange"]),
        ("Ev 0.5 Üst", stats["home_over05_pct"], C["green"]),
        ("Ev 1.5 Üst", stats["home_over15_pct"], C["green"]),
        ("Dep 0.5 Üst", stats["away_over05_pct"], C["red"]),
        ("Dep 1.5 Üst", stats["away_over15_pct"], C["red"]),
    ]
    if row is not None:
        priced = [candidate for candidate in candidates if prediction_odd(row, candidate[0]) is not None]
        if priced:
            candidates = priced
    label, raw_confidence, color = max(candidates, key=lambda item: item[1])
    confidence = calibrated_confidence(raw_confidence, stats["n"])
    return {"label": label, "confidence": round(confidence, 1), "raw_confidence": raw_confidence,
            "sample_size": stats["n"], "color": color}


def calibrated_confidence(raw_confidence, sample_size):
    """Geçmiş oranın %95 Wilson alt sınırını, en fazla %75 olarak verir.

    Bu değer bir kazanma garantisi veya saf olasılık değil, geçmiş verinin
    temkinli istatistik desteğidir. Alt sınır, örneğin 64 maçtaki %86'lık ham
    oranı doğrudan %86 göstermek yerine örneklem belirsizliğini hesaba katar.
    """
    n = max(0, int(sample_size or 0))
    if not n:
        return 0.0
    p = max(0.0, min(1.0, float(raw_confidence or 0.0) / 100.0))
    z = 1.96  # %95 güven düzeyi
    denominator = 1.0 + z * z / n
    centre = p + z * z / (2.0 * n)
    margin = z * math.sqrt((p * (1.0 - p) + z * z / (4.0 * n)) / n)
    lower_bound = max(0.0, (centre - margin) / denominator) * 100.0
    return round(min(75.0, lower_bound), 1)


def dominant_pct(stats):
    return best_prediction(stats)["confidence"]


def prediction_odd(row, prediction):
    column = PREDICTION_ODDS_COLUMNS.get(str(prediction or ""))
    if not column:
        return None
    value = to_float(row.get(column))
    return value if value and value > 1.0 else None


MARKET_PAIRS = {
    "MS 1": ("MS 1", "MS X", "MS 2"), "MS X": ("MS 1", "MS X", "MS 2"), "MS 2": ("MS 1", "MS X", "MS 2"),
    "2.5 Alt": ("2.5 Alt", "2.5 Üst"), "2.5 Üst": ("2.5 Alt", "2.5 Üst"),
    "KG Var": ("KG Var", "KG Yok"), "KG Yok": ("KG Var", "KG Yok"),
}


def implied_probability(row, prediction):
    """Aynı pazardaki oranları normalize ederek marjdan arındırılmış piyasa olasılığı verir."""
    market = MARKET_PAIRS.get(prediction)
    if not market:
        return None
    odds = [prediction_odd(row, item) for item in market]
    if any(odd is None for odd in odds):
        return None
    inverse = [1 / odd for odd in odds]
    return inverse[market.index(prediction)] / sum(inverse) * 100


def model_market_predictions(stats, row):
    """Her maç için Banko, Riskli ve Ya Tutarsa olmak üzere üç ayrı seçim üretir."""
    available = [
        ("MS 1", stats["home_pct"]), ("MS X", stats["draw_pct"]), ("MS 2", stats["away_pct"]),
        ("2.5 Alt", stats["under25_pct"]), ("2.5 Üst", stats["over25_pct"]),
        ("KG Var", stats["btts_pct"]), ("KG Yok", stats["btts_no_pct"]),
    ]
    available = [(label, pct) for label, pct in available if implied_probability(row, label) is not None]
    if not available:
        return [{"tier": tier, "prediction": "Piyasa verisi yok", "model_pct": 0.0,
                 "market_pct": None, "price": None, "actionable": False,
                 "decision": "Pas geç", "sample_size": stats["n"]} for tier in PREDICTION_TIERS]

    candidates = []
    for prediction, raw_confidence in available:
        price = prediction_odd(row, prediction)
        model_pct = calibrated_confidence(raw_confidence, stats["n"])
        candidates.append({
            "prediction": prediction,
            "model_pct": model_pct,
            "market_pct": implied_probability(row, prediction),
            "price": price,
            "actionable": bool(stats["n"] >= ACTIONABLE_MIN_SAMPLES and price and model_pct >= ACTIONABLE_MIN_CONFIDENCE),
            "sample_size": stats["n"],
        })

    # Banko: en yüksek kalibre edilmiş destek. Riskli: orta oranlı, ayrı bir
    # pazar. Ya Tutarsa: kalanlardaki yüksek oranlı seçim. Aynı pazar üç kez
    # önerilmez.
    remaining = list(candidates)
    banko = max(remaining, key=lambda item: (item["model_pct"], -item["price"]))
    remaining.remove(banko)
    middle = [item for item in remaining if 1.50 <= item["price"] <= 2.75] or remaining
    riskli = max(middle, key=lambda item: (item["model_pct"], -abs(item["price"] - 2.10))) if middle else None
    if riskli is not None:
        remaining.remove(riskli)
    long_shots = [item for item in remaining if item["price"] >= 2.75] or remaining
    ya_tutarsa = max(long_shots, key=lambda item: (item["price"], item["model_pct"])) if long_shots else None

    result = []
    for tier, candidate in zip(PREDICTION_TIERS, (banko, riskli, ya_tutarsa)):
        if candidate is None:
            continue
        result.append({**candidate, "tier": tier, "decision": "İncele" if candidate["actionable"] else "Pas geç"})
    return result


def model_market_signal(stats, row):
    """Eski tek-tahmin kullanan çağrılar için Banko seçimini döndürür."""
    return model_market_predictions(stats, row)[0]


def format_odd(value):
    odd = to_float(value)
    return "" if odd is None else f"{odd:.2f}"


def add_prediction_odds(data):
    if data.empty or "Tahmin" not in data.columns:
        return data.copy()
    result = data.copy()
    result["Oran"] = [format_odd(prediction_odd(row, row.get("Tahmin"))) for _, row in result.iterrows()]
    return result


def parse_percent(value):
    text = str(value or "").strip().replace("%", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def prediction_quality_score(confidence, match_count, key_count=1, tolerance=8.0, scoped=False):
    # ``confidence`` geriye dönük API uyumluluğu için alınır; veri kalitesi
    # ile tahmin olasılığını birbirine karıştırmamak için puana katılmaz.
    # Örneğin 530 eşleşmeli, düşük olasılıklı bir pazar "Sınırlı veri" değildir.
    _ = max(0.0, min(100.0, float(confidence or 0.0)))
    match_count = max(0, int(match_count or 0))
    key_count = max(1, int(key_count or 1))
    tolerance = max(0.0, float(tolerance or 0.0))
    # Logaritmik örneklem ölçeği: her yeni veri miktarı faydalıdır ancak
    # 50'den 100'e çıkışla 500'den 550'ye çıkış aynı etkiye sahip değildir.
    score = min(88.0, math.log2(match_count + 1) * 10.0)
    score += min(key_count, 6) * 2.0
    score += 2.0 if scoped else 0.0
    score -= max(0.0, tolerance - 8.0) * 0.5
    score = max(0.0, min(100.0, score))
    if score >= 80:
        label = "Guclu veri"
    elif score >= 55:
        label = "Yeterli veri"
    elif score >= 35:
        label = "Sinirli veri"
    else:
        label = "Az veri"
    return {"score": round(score, 1), "label": label}


def sort_today_data(data, sort_mode=TODAY_SORT_CONFIDENCE):
    if data.empty:
        return data.copy()
    result = data.copy()
    result["_confidence_num"] = result.get("Güven", pd.Series(index=result.index, dtype="object")).apply(parse_percent)
    result["_odd_num"] = pd.to_numeric(
        result.get("Oran", pd.Series(index=result.index, dtype="object")).astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    ).fillna(0.0)
    result["_quality_num"] = pd.to_numeric(
        result.get("_quality_score", pd.Series(index=result.index, dtype="object")),
        errors="coerce",
    ).fillna(0.0)
    result["_match_count"] = pd.to_numeric(
        result.get("Eşleşme", pd.Series(index=result.index, dtype="object")),
        errors="coerce",
    ).fillna(0.0)
    result["_value_score"] = result["_confidence_num"] * result["_odd_num"]
    dates = result["Tarih"] if "Tarih" in result else pd.Series("", index=result.index)
    times = result["Saat"] if "Saat" in result else pd.Series("", index=result.index)
    result["_kickoff"] = pd.to_datetime(
        dates.astype(str) + " " + times.astype(str),
        format="%d-%m-%Y %H:%M",
        errors="coerce",
    )
    sort_columns = {
        TODAY_SORT_VALUE: ["_value_score", "_quality_num", "_confidence_num"],
        TODAY_SORT_ODD: ["_odd_num", "_confidence_num", "_quality_num"],
        TODAY_SORT_QUALITY: ["_quality_num", "_confidence_num", "_match_count"],
        TODAY_SORT_MATCHES: ["_match_count", "_confidence_num", "_quality_num"],
        TODAY_SORT_TIME: ["_kickoff", "_confidence_num"],
    }
    columns = sort_columns.get(sort_mode, ["_confidence_num", "_quality_num", "_match_count"])
    ascending = [sort_mode == TODAY_SORT_TIME] + [False] * (len(columns) - 1)
    return result.sort_values(columns, ascending=ascending).drop(
        columns=["_confidence_num", "_odd_num", "_quality_num", "_match_count", "_value_score", "_kickoff"],
        errors="ignore",
    )


def coupon_simulation(total_odds, stake):
    total_odds = to_float(total_odds)
    stake = to_float(stake)
    if total_odds is None or total_odds <= 1.0 or stake is None or stake <= 0:
        return {"stake": 0.0, "total_odds": total_odds or 0.0, "payout": 0.0, "profit": 0.0}
    payout = stake * total_odds
    return {"stake": stake, "total_odds": total_odds, "payout": payout, "profit": payout - stake}


def total_coupon_odds(data):
    if data.empty or "Oran" not in data.columns:
        return None
    total = 1.0
    used = 0
    for value in data["Oran"]:
        odd = to_float(value)
        if odd is None or odd <= 1.0:
            continue
        total *= odd
        used += 1
    return total if used else None


def select_coupon_candidates(
    data,
    min_confidence,
    max_selections,
    min_kickoff_gap_minutes=60,
    diversify_leagues=True,
):
    """Geçerli tahminlerden, kullanıcının kupon kurallarına uygun seçimleri döndürür."""
    if data.empty:
        return pd.DataFrame()
    candidates = data.copy()
    candidates["_confidence"] = pd.to_numeric(
        candidates.get("Güven", pd.Series(index=candidates.index, dtype="object"))
        .astype(str).str.replace("%", "", regex=False).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    match_counts = candidates["Eşleşme"] if "Eşleşme" in candidates else pd.Series(0, index=candidates.index)
    candidates["_match_count"] = pd.to_numeric(match_counts, errors="coerce").fillna(0)
    candidates["_quality_score_sort"] = pd.to_numeric(
        candidates.get("_quality_score", pd.Series(0, index=candidates.index)),
        errors="coerce",
    ).fillna(0)
    candidates["_odd"] = pd.to_numeric(
        candidates.get("Oran", pd.Series(index=candidates.index, dtype="object"))
        .astype(str).str.replace(",", ".", regex=False),
        errors="coerce",
    )
    invalid_predictions = {"", "Oran yok", "Veri yok", "Piyasa verisi yok", "Pas geç"}
    candidates = candidates[
        (candidates["_confidence"] >= max(0.0, float(min_confidence or 0.0)))
        & (candidates["_odd"] > 1.0)
        & (~candidates.get("Tahmin", pd.Series("", index=candidates.index)).astype(str).isin(invalid_predictions))
    ]
    key_columns = [column for column in ("Tarih", "Saat", "Lig", "Maç") if column in candidates.columns]
    if key_columns:
        candidates = candidates.drop_duplicates(subset=key_columns, keep="first")
    candidates["_kickoff"] = pd.to_datetime(
        candidates.get("Tarih", pd.Series("", index=candidates.index)).astype(str) + " " +
        candidates.get("Saat", pd.Series("", index=candidates.index)).astype(str),
        dayfirst=True,
        errors="coerce",
    )
    candidates = candidates.sort_values(
        ["_confidence", "_quality_score_sort", "_match_count"],
        ascending=[False, False, False],
    )
    selected = []
    used_leagues = set()
    kickoff_times = []
    for _, row in candidates.iterrows():
        league = str(row.get("Lig", "")).strip().casefold()
        kickoff = row.get("_kickoff")
        duplicate_league = bool(diversify_leagues and league and league in used_leagues)
        nearby_kickoff = (
            pd.notna(kickoff)
            and any(abs((kickoff - earlier).total_seconds()) < min_kickoff_gap_minutes * 60 for earlier in kickoff_times)
        )
        if duplicate_league or nearby_kickoff:
            continue
        selected.append(row)
        if league:
            used_leagues.add(league)
        if pd.notna(kickoff):
            kickoff_times.append(kickoff)
        if len(selected) == max_selections:
            break
    if not selected:
        return pd.DataFrame()
    return pd.DataFrame(selected).drop(
        columns=["_kickoff", "_quality_score_sort", "_confidence", "_match_count", "_odd"],
        errors="ignore",
    ).reset_index(drop=True)


def best_combo_search(df, entered, tol_base, min_matches, logger=None):
    def log(msg):
        if logger:
            logger(msg)
    keys = list(entered.keys())
    if not keys:
        return None
    # Zero means an empty match set qualifies immediately, preventing the
    # fallback combination/tolerance search from ever running.
    try:
        min_matches = max(FIXED_MINIMUM_MATCHES, int(min_matches))
    except (TypeError, ValueError):
        min_matches = FIXED_MINIMUM_MATCHES
    # Normalize the relevant odds once. The search evaluates many field
    # combinations, so coercing each column inside every candidate would make
    # a larger history unnecessarily slow.
    df = df.copy()
    for key in keys:
        column = FIELD_BY_KEY[key]["col"]
        if column in df.columns and not pd.api.types.is_numeric_dtype(df[column]):
            df[column] = pd.to_numeric(
                df[column].astype("string").str.replace(",", ".", regex=False),
                errors="coerce",
            )
    for mult in (1.0, 1.5, 2.0, 3.0):
        tol = round(tol_base * mult, 2)
        full = filter_by_fields(df, entered, tol)
        log(f"Tolerans %{tol:.1f} — {len(keys)} alanın tamamı ile eşleştiriliyor... {len(full)} maç bulundu.")
        if len(full) >= min_matches:
            return {"keys": keys, "tol": tol, "matches": full}
        if mult > 1.0:
            log("  Yetersiz eşleşme → tolerans büyütüldü.")
        best = None
        for r in range(len(keys) - 1, 0, -1):
            found_at_r = None
            for combo in combinations(keys, r):
                sub = {k: entered[k] for k in combo}
                mm = filter_by_fields(df, sub, tol)
                if len(mm) >= min_matches:
                    dom = dominant_pct(outcome_stats(mm))
                    if found_at_r is None or dom > found_at_r[0]:
                        found_at_r = (dom, combo, mm)
            if found_at_r:
                best = found_at_r
                labels = " + ".join(FIELD_BY_KEY[k]["label"] for k in found_at_r[1])
                log(f"  {r} alanlı en iyi kombinasyon: {labels} → {len(found_at_r[2])} maç.")
                break
        if best:
            _, combo, mm = best
            return {"keys": list(combo), "tol": tol, "matches": mm}
    log("Hiçbir kombinasyonda yeterli eşleşme bulunamadı.")
    return {"keys": keys, "tol": tol_base * 3, "matches": df.iloc[0:0]}


def prediction_history_stats(records):
    total = won = lost = pending = 0
    by_market = {}
    for record in records:
        status = str(record.get("durum", "Bekliyor"))
        if status == "Tuttu":
            won += 1
        elif status == "Yatti":
            lost += 1
        else:
            pending += 1
            continue
        total += 1
        market = str(record.get("tahmin", "-"))
        entry = by_market.setdefault(market, {"total": 0, "won": 0, "lost": 0})
        entry["total"] += 1
        entry["won"] += int(status == "Tuttu")
        entry["lost"] += int(status == "Yatti")
    success_pct = (won / total * 100) if total else 0.0
    return {
        "graded_total": total,
        "won": won,
        "lost": lost,
        "pending": pending,
        "success_pct": success_pct,
        "by_market": by_market,
    }


def confidence_calibration_report(records):
    """Arayüzden bağımsız güven denetimi; yalnızca sonuçlanmış kayıtları değerlendirir."""
    bands = {"<60": [], "60-74": [], "75+": []}
    for record in records:
        status = str(record.get("durum", record.get("status", ""))).casefold()
        if status not in {"tuttu", "yatti"}:
            continue
        confidence = to_float(record.get("guven", record.get("confidence")))
        if confidence is None:
            continue
        key = "75+" if confidence >= 75 else "60-74" if confidence >= 60 else "<60"
        bands[key].append((confidence, status == "tuttu"))
    return {key: {"n": len(values), "expected_pct": round(sum(v[0] for v in values) / len(values), 1) if values else 0.0,
                  "actual_pct": round(sum(v[1] for v in values) * 100 / len(values), 1) if values else 0.0}
            for key, values in bands.items()}


def strategy_performance_report(records, min_sample=5):
    """Rank completed predictions by the odds-field combination and tolerance.

    A group needs ``min_sample`` results to be eligible.  Ranking favours ROI,
    then a Wilson lower bound on the win rate so a small hot streak is not
    presented as the best strategy.
    """
    groups = {}
    for record in records:
        status = str(record.get("durum", record.get("status", ""))).casefold()
        if status not in {"tuttu", "yatti"}:
            continue
        combo = str(record.get("kombo", record.get("combo", ""))).strip() or "Kombinasyon kaydedilmedi"
        tolerance = to_float(record.get("tolerans", record.get("tolerance")))
        tolerance_label = "—" if tolerance is None else f"%{tolerance:.1f}"
        key = (combo, tolerance_label)
        entry = groups.setdefault(key, {"combo": combo, "tolerance": tolerance, "total": 0, "won": 0, "lost": 0,
                                        "stake": 0.0, "return": 0.0, "markets": set(), "leagues": set()})
        entry["total"] += 1
        entry["won"] += int(status == "tuttu")
        entry["lost"] += int(status == "yatti")
        entry["stake"] += 1.0
        if status == "tuttu":
            entry["return"] += to_float(record.get("oran", record.get("price"))) or 0.0
        market = str(record.get("tahmin", record.get("market", ""))).strip()
        league = str(record.get("lig", record.get("league", ""))).strip()
        if market:
            entry["markets"].add(market)
        if league:
            entry["leagues"].add(league)

    report = []
    z = 1.96
    for entry in groups.values():
        n = entry["total"]
        rate = entry["won"] / n
        # Conservative 95% lower bound for a binomial success rate.
        lower_bound = (rate + z * z / (2 * n) - z * ((rate * (1 - rate) + z * z / (4 * n)) / n) ** 0.5) / (1 + z * z / n)
        entry.update({
            "success_rate": round(rate * 100, 1),
            "roi": round((entry["return"] - entry["stake"]) * 100 / entry["stake"], 1) if entry["stake"] else 0.0,
            "net": round(entry["return"] - entry["stake"], 2),
            "lower_bound": round(lower_bound * 100, 1),
            "eligible": n >= min_sample,
            "markets": ", ".join(sorted(entry["markets"])) or "—",
            "leagues": ", ".join(sorted(entry["leagues"])) or "—",
        })
        report.append(entry)
    report.sort(key=lambda item: (not item["eligible"], -item["roi"], -item["lower_bound"], -item["total"]))
    return {"min_sample": min_sample, "groups": report, "best": next((item for item in report if item["eligible"]), None)}
