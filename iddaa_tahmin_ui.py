# -*- coding: utf-8 -*-
"""
İddaa Oran Eşleştirici — Tahmin Arayüzü  (v9.0)
================================================

Kullanım mantığı:
  1) Maç bilgisini (opsiyonel) ve elindeki oranları gir.
  2) "ORANLARI EŞLEŞTİR" butonuna bas.
  3) Uygulama, excel/bahis_oranlari_2026.xlsx ana dosyası ve excel klasöründeki
     diğer veri dosyalarındaki geçmiş maçlardan, girdiğin
     oranlara (tolerans dahilinde) en çok benzeyen maçları bulur ve bu
     maçların sonuçlarından (Ev/Berabere/Deplasman, 2.5 Alt/Üst, KG,
     en sık skor) bir olasılık tablosu çıkarır.
  4) Beğendiğin tahmini tek tıkla predictions.json geçmişine kaydedebilirsin.

Veri dosyaları:
  excel/bahis_oranlari_2026.xlsx  -> ana geçmiş oran & skor verisi
  excel/*.xlsx                    -> uygulamanın birlikte kullandığı ek veriler
  predictions.json     -> kaydedilen tahminler
  settings.json         -> sütun genişlikleri / tercihler
"""

import json
import os
import threading
from datetime import datetime
from itertools import combinations
from pathlib import Path

import tkinter as tk
from tkinter import ttk, messagebox, filedialog

import customtkinter as ctk
import pandas as pd

from excel.merger import discover_excel_files, merge_excel_files
from data_store import (live_signal_summary, load_matches, load_today_display, load_tracked_predictions, pending_live_signals,
                        pending_tracked_predictions, replace_upcoming_matches, save_live_signal,
                        replace_today_display, save_live_snapshot, save_tracked_prediction,
                        settle_live_signal, settle_tracked_prediction, tracked_prediction_summary)
import iddaa_scraper
from live_api import get_fixture, get_fixture_statistics, get_live_fixtures, get_live_odds, quota_status
from live_signals import build_goal_signals
from iddaa_core import (
    inclusive_day_count,
    latest_data_date,
    missing_update_range,
    prediction_history_stats,
    strategy_performance_report,
)
from iddaa_oynanmamis_scraper import fetch_match_result, fetch_today_unplayed_odds
from football_data_api import get_finished_score
from prematch_analysis import settle_market, split_match

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("dark-blue")

# ─────────────────────────────────────────────────────────────
#  SABİTLER
# ─────────────────────────────────────────────────────────────
BASE_DIR = Path(__file__).resolve().parent
EXCEL_DIR = BASE_DIR / "excel"
DATA_FILE = BASE_DIR / "data" / "iddaa.sqlite"
PREDICTIONS_FILE = BASE_DIR / "predictions.json"
SETTINGS_FILE = BASE_DIR / "settings.json"
DATA_SCOPE_ALL = "Tum Veri"
DATA_SCOPE_2026 = "Sadece 2026"
DATA_SCOPE_365 = "Son 365 Gun"
DATA_SCOPE_180 = "Son 180 Gun"
DATA_SCOPES = [DATA_SCOPE_ALL, DATA_SCOPE_2026, DATA_SCOPE_365, DATA_SCOPE_180]

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

# key, kısa etiket, excel sütun adı, renk
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

FIELD_GROUPS = [
    ("⚽  Maç Sonucu",        ["ms1", "msx", "ms2"]),
    ("🎯  2.5 Alt / Üst",     ["au25", "uu25"]),
    ("🤝  Karşılıklı Gol",    ["kgvar", "kgyok"]),
    ("🔀  Çifte Şans",        ["cs1x", "cs12", "csx2"]),
    ("⏱️  İlk Yarı Alt/Üst",  ["iy05a", "iy05u", "iy15a", "iy15u"]),
]

# Veri gözat sekmesinde gösterilecek sütunlar: (başlık, kaynak, genişlik)
BROWSE_COLUMNS = [
    ("Tarih", "Tarih", 92), ("Saat", "Saat", 60), ("Lig", "Lig", 150),
    ("Ev Sahibi", "_ev", 150), ("Deplasman", "_dep", 150),
    ("MS", "Skor", 60), ("IY", "İY Skor", 55),
    ("MS 1", "Maç Sonucu (1)", 62), ("MS X", "Maç Sonucu (X)", 62), ("MS 2", "Maç Sonucu (2)", 62),
    ("AU 2.5", "2,5 Alt/Üst (Alt)", 65), ("UU 2.5", "2,5 Alt/Üst (Üst)", 65),
    ("KG Var", "Karşılıklı Gol (Var)", 70), ("KG Yok", "Karşılıklı Gol (Yok)", 70),
    ("Çift 1X", "Çifte Şans (1-X)", 68), ("Çift 12", "Çifte Şans (1-2)", 68), ("Çift X2", "Çifte Şans (X-2)", 68),
]
BROWSE_KEY_BY_TITLE = {
    "MS 1": "ms1", "MS X": "msx", "MS 2": "ms2", "AU 2.5": "au25", "UU 2.5": "uu25",
    "KG Var": "kgvar", "KG Yok": "kgyok", "Çift 1X": "cs1x", "Çift 12": "cs12", "Çift X2": "csx2",
}

PREDICTION_ODDS_COLUMNS = {
    "MS 1": "Maç Sonucu (1)",
    "MS X": "Maç Sonucu (X)",
    "MS 2": "Maç Sonucu (2)",
    "2.5 Alt": "2,5 Alt/Üst (Alt)",
    "2.5 Üst": "2,5 Alt/Üst (Üst)",
    "KG Var": "Karşılıklı Gol (Var)",
    "KG Yok": "Karşılıklı Gol (Yok)",
}

TODAY_TREE_COLUMNS = ("Saat", "Lig", "Maç", "MS 1", "MS X", "MS 2", "Tahmin", "Oran", "Model", "Piyasa", "Grup", "Karar", "Eşleşme")
CORE_ODDS_COLUMNS = ["Maç Sonucu (1)", "Maç Sonucu (X)", "Maç Sonucu (2)"]
COUPON_RISK_PROFILES = {
    "Guvenli": {"min_confidence": 80, "max_selections": 3, "min_gap": 90},
    "Dengeli": {"min_confidence": 70, "max_selections": 5, "min_gap": 60},
    "Agresif": {"min_confidence": 60, "max_selections": 7, "min_gap": 30},
}


# ─────────────────────────────────────────────────────────────
#  YARDIMCI FONKSİYONLAR
# ─────────────────────────────────────────────────────────────
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


def load_json(path, default):
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
                return data if isinstance(data, type(default)) else default
        except Exception:
            return default
    return default


def save_json(path, data):
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as exc:
        print(f"[UYARI] {path} yazılamadı: {exc}")
        return False


def load_odds_data(path):
    """Excel'i okuyup oran sütunlarını sayısala çevirir, yardımcı sütunlar ekler."""
    path = Path(path)
    df = load_matches(path) if path.suffix.lower() == ".sqlite" else pd.read_excel(path)
    return prepare_odds_data(df)


def prepare_odds_data(df):
    """Ham veri çerçevesini arayüzde kullanılacak biçime dönüştürür."""
    df = df.copy()
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
    if "Tarih" in df.columns:
        df["_tarih_dt"] = pd.to_datetime(df["Tarih"], format="%d-%m-%Y", errors="coerce")
        # Bazı elle hazırlanmış Excel dosyalarında başlık satırı veri olarak tekrar edebiliyor.
        # Tarihi olmayan satırlar geçmiş maç olarak kullanılmamalı.
        df = df.loc[df["_tarih_dt"].notna()].copy()
    df["_search"] = (df.get("Lig", "").astype(str) + " " + df.get("Maç", "").astype(str)).str.lower()
    df.attrs["duplicates_removed"] = duplicates_removed
    return df


def load_excel_folder(primary_file=DATA_FILE):
    """Excel klasöründeki tüm uygun dosyaları, ana dosyayı önceliklendirerek yükler."""
    primary_file = Path(primary_file)
    files = discover_excel_files(EXCEL_DIR)
    if primary_file.exists() and primary_file not in files:
        files.append(primary_file)
    if not files:
        raise FileNotFoundError(f"Excel klasöründe veri dosyası bulunamadı: {EXCEL_DIR}")

    primary_resolved = primary_file.resolve()
    files = [path for path in files if path.resolve() != primary_resolved] + [
        path for path in files if path.resolve() == primary_resolved
    ]
    frames = []
    used_files = []
    for file_path in files:
        frame = pd.read_excel(file_path)
        required = {"Tarih", "Saat", "Lig", "Maç"}
        if required.issubset(frame.columns):
            frames.append(frame)
            used_files.append(file_path.name)

    if not frames:
        raise ValueError("Excel klasöründe maç verisi sütunlarını içeren bir dosya bulunamadı.")
    combined = prepare_odds_data(pd.concat(frames, ignore_index=True))
    combined.attrs["source_files"] = used_files
    return combined


def filter_data_by_scope(df, scope, today=None):
    """Tahminlerde kullanılacak geçmiş veri aralığını seçer."""
    if df.empty or scope == DATA_SCOPE_ALL or "_tarih_dt" not in df.columns:
        return df.copy()

    dates = pd.to_datetime(df["_tarih_dt"], dayfirst=True, errors="coerce")
    if scope == DATA_SCOPE_2026:
        return df.loc[dates.dt.year == 2026].copy()

    day_count_by_scope = {
        DATA_SCOPE_365: 365,
        DATA_SCOPE_180: 180,
    }
    day_count = day_count_by_scope.get(scope)
    if not day_count:
        return df.copy()

    today_ts = pd.Timestamp(today or datetime.now().date())
    start = today_ts - pd.Timedelta(days=day_count)
    return df.loc[dates.notna() & (dates >= start) & (dates <= today_ts)].copy()


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
    """Eşleşen maçların sonuç istatistiklerini hesaplar.

    Bugunun maclari icin ayni lig ve daha yeni maclara hafif bir agirlik
    verilebilir. Ham eslesme sayisi her zaman ``n`` alaninda korunur.
    """
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
    }


def filter_by_fields(df, entered, tol_pct):
    m = df
    for key, val in entered.items():
        col = FIELD_BY_KEY[key]["col"]
        if col not in df.columns:
            return df.iloc[0:0]
        lo, hi = val * (1 - tol_pct / 100.0), val * (1 + tol_pct / 100.0)
        m = m[m[col].notna() & (m[col] >= lo) & (m[col] <= hi)]
    return m


def best_prediction(stats):
    """Geçmiş eşleşmelerde en yüksek destek oranına sahip bahis sonucunu seçer."""
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
    label, confidence, color = max(candidates, key=lambda item: item[1])
    return {"label": label, "confidence": confidence, "color": color}


def dominant_pct(stats):
    return best_prediction(stats)["confidence"]


def prediction_odd(row, prediction):
    column = PREDICTION_ODDS_COLUMNS.get(str(prediction or ""))
    if not column:
        return None
    value = to_float(row.get(column))
    return value if value and value > 1.0 else None


def format_odd(value):
    odd = to_float(value)
    return "" if odd is None else f"{odd:.2f}"


def add_prediction_odds(data):
    if data.empty or "Tahmin" not in data.columns:
        return data.copy()
    result = data.copy()
    result["Oran"] = [format_odd(prediction_odd(row, row.get("Tahmin"))) for _, row in result.iterrows()]
    return result


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


def select_coupon_candidates(data, min_confidence, max_selections, min_kickoff_gap_minutes=60):
    """Yuksek guvenli secimlerden lig ve saat cakismasi az bir kupon kurar."""
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
    candidates = candidates[candidates["_confidence"] >= min_confidence]
    candidates = candidates.drop_duplicates(subset=["Tarih", "Saat", "Lig", "Maç"], keep="first")
    candidates["_kickoff"] = pd.to_datetime(
        candidates["Tarih"].astype(str) + " " + candidates["Saat"].astype(str),
        dayfirst=True,
        errors="coerce",
    )
    candidates = candidates.sort_values(["_confidence", "_match_count"], ascending=[False, False])

    selected = []
    used_leagues = set()
    kickoff_times = []
    for _, row in candidates.iterrows():
        league = str(row.get("Lig", "")).strip().casefold()
        kickoff = row.get("_kickoff")
        duplicate_league = bool(league) and league in used_leagues
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
    return pd.DataFrame(selected).drop(columns=["_kickoff"], errors="ignore").reset_index(drop=True)


def best_combo_search(df, entered, tol_base, min_matches, logger=None):
    """
    Girilen tüm oranlarla eşleşme arar; yetersizse toleransı kademeli
    büyütür (x1 → x1.5 → x2 → x3); yine de yetersizse en iyi alt
    kombinasyonu (daha az alan) dener. Destek oranı en yüksek bahis
    sonucunu veren kombinasyonu seçer.
    """
    def log(msg):
        if logger:
            logger(msg)

    keys = list(entered.keys())
    if not keys:
        return None

    for mult in (1.0, 1.5, 2.0, 3.0):
        tol = round(tol_base * mult, 2)
        full = filter_by_fields(df, entered, tol)
        log(f"Tolerans %{tol:.1f} — {len(keys)} alanın tamamı ile eşleştiriliyor... {len(full)} maç bulundu.")
        if len(full) >= min_matches:
            return {"keys": keys, "tol": tol, "matches": full}

        if mult > 1.0:
            log(f"  Yetersiz eşleşme → tolerans büyütüldü.")

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


from iddaa_core import (  # noqa: E402,F401
    COUPON_RISK_PROFILES,
    CORE_ODDS_COLUMNS,
    DATA_SCOPE_180,
    DATA_SCOPE_2026,
    DATA_SCOPE_365,
    DATA_SCOPE_ALL,
    DATA_SCOPES,
    FIELD_BY_KEY,
    FIXED_MINIMUM_MATCHES,
    FIXED_TOLERANCE,
    MATCH_SCOPE_SMART,
    MATCH_SCOPES,
    ODDS_FIELDS,
    PREDICTION_ODDS_COLUMNS,
    TODAY_SORT_CONFIDENCE,
    TODAY_SORT_OPTIONS,
    TODAY_SORT_VALUE,
    add_prediction_odds,
    best_combo_search,
    best_prediction,
    coupon_simulation,
    data_quality_summary,
    filter_history_by_match_scope,
    filter_data_by_scope,
    format_odd,
    inclusive_day_count,
    infer_country_from_league,
    latest_data_date,
    load_odds_data,
    missing_update_range,
    model_market_predictions,
    outcome_stats,
    parse_score,
    prediction_odd,
    prediction_quality_score,
    prediction_history_stats,
    select_coupon_candidates,
    sort_today_data,
    to_float,
    total_coupon_odds,
)


# ─────────────────────────────────────────────────────────────
#  KÜÇÜK BİLEŞENLER
# ─────────────────────────────────────────────────────────────
class Pill(ctk.CTkFrame):
    """Başlık + büyük değer gösteren küçük istatistik kartı."""
    def __init__(self, parent, title, **kw):
        super().__init__(parent, fg_color=C["bg3"], corner_radius=12,
                          border_width=1, border_color=C["border"], **kw)
        ctk.CTkLabel(self, text=title, font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(anchor="w", padx=14, pady=(10, 0))
        self.value_lbl = ctk.CTkLabel(self, text="—", font=ctk.CTkFont("Segoe UI", 20, "bold"),
                                       text_color=C["text1"])
        self.value_lbl.pack(anchor="w", padx=14, pady=(0, 10))

    def set(self, text, color=None):
        self.value_lbl.configure(text=text, text_color=color or C["text1"])


class StatusChip(ctk.CTkFrame):
    def __init__(self, parent, title, **kw):
        super().__init__(parent, fg_color=C["bg2"], corner_radius=8, border_width=1, border_color=C["border"], **kw)
        ctk.CTkLabel(self, text=title, font=ctk.CTkFont("Segoe UI", 9, "bold"), text_color=C["text3"]).pack(side="left", padx=(9, 5), pady=5)
        self.value = ctk.CTkLabel(self, text="—", font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=C["text1"])
        self.value.pack(side="left", padx=(0, 9), pady=5)

    def set(self, text, color=None):
        self.value.configure(text=text, text_color=color or C["text1"])


class SignalCard(ctk.CTkFrame):
    def __init__(self, parent, signal):
        actionable = signal.get("actionable")
        super().__init__(parent, fg_color=C["bg3"] if actionable else C["bg2"], corner_radius=12,
                         border_width=1, border_color=C["green"] if actionable else C["border"])
        ctk.CTkLabel(self, text=signal["match"], font=ctk.CTkFont("Segoe UI", 13, "bold"), text_color=C["text1"]).pack(anchor="w", padx=14, pady=(12, 2))
        state = "BAHİS YAPILABİLİR" if actionable else "İZLEMEDE"
        ctk.CTkLabel(self, text=f"{state}  ·  {signal['minute']}. dk", font=ctk.CTkFont("Segoe UI", 10, "bold"),
                     text_color=C["green"] if actionable else C["yellow"]).pack(anchor="w", padx=14)
        ctk.CTkLabel(self, text=f"{signal['market']}   |   Oran {signal.get('price') or '—'}   |   Güven %{signal['confidence']}",
                     font=ctk.CTkFont("Segoe UI", 12), text_color=C["cyan"] if actionable else C["text2"]).pack(anchor="w", padx=14, pady=(8, 2))
        ctk.CTkLabel(self, text=" • ".join(signal["reasons"]), font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"], wraplength=460, justify="left").pack(anchor="w", padx=14, pady=(0, 12))


class OddsField(ctk.CTkFrame):
    """Renkli etiket + oran girişi tek satırlık bileşen."""
    def __init__(self, parent, key, label, color, on_enter=None, **kw):
        super().__init__(parent, fg_color="transparent", **kw)
        self.key = key
        chip = ctk.CTkFrame(self, fg_color=C["bg4"], corner_radius=6, width=78, height=26)
        chip.pack(side="left"); chip.pack_propagate(False)
        ctk.CTkLabel(chip, text=label, font=ctk.CTkFont("Segoe UI", 10, "bold"),
                     text_color=color).place(relx=0.5, rely=0.5, anchor="center")
        self.var = tk.StringVar()
        self.entry = ctk.CTkEntry(self, width=72, height=26, fg_color=C["bg2"],
                                   border_color=C["border"], text_color=C["text1"],
                                   font=ctk.CTkFont("Consolas", 12), corner_radius=6,
                                   placeholder_text="orn 2.10", textvariable=self.var)
        self.entry.pack(side="left", padx=(6, 0))
        if on_enter:
            self.entry.bind("<Return>", lambda e: on_enter())

    def get_float(self):
        return to_float(self.var.get())

    def set_value(self, value):
        self.var.set("" if value is None else f"{value:.2f}".rstrip("0").rstrip("."))

    def clear(self):
        self.var.set("")


class ProbBar(tk.Canvas):
    """Ev / Berabere / Deplasman olasılıklarını gösteren yatay yığın çubuk."""
    def __init__(self, parent, height=34, **kw):
        super().__init__(parent, height=height, bg=C["bg2"], highlightthickness=0, **kw)
        self.bind("<Configure>", lambda e: self._redraw())
        self._pcts = (33.3, 33.4, 33.3)
        self._labels = ("Ev %33", "X %33", "Dep %33")

    def update_probs(self, home_pct, draw_pct, away_pct):
        self._pcts = (home_pct, draw_pct, away_pct)
        self._labels = (f"Ev %{home_pct:.0f}", f"X %{draw_pct:.0f}", f"Dep %{away_pct:.0f}")
        self._redraw()

    def _redraw(self):
        self.delete("all")
        w = self.winfo_width() or 400
        h = self.winfo_height() or 34
        colors = (C["green"], C["yellow"], C["red"])
        x = 0
        for pct, lbl, col in zip(self._pcts, self._labels, colors):
            seg_w = max(w * pct / 100.0, 1)
            self.create_rectangle(x, 0, x + seg_w, h, fill=col, width=0)
            if seg_w > 34:
                self.create_text(x + seg_w / 2, h / 2, text=lbl, fill=C["bg0"],
                                  font=("Segoe UI", 10, "bold"))
            x += seg_w


class AnalysisGraph(tk.Canvas):
    """Tahmin sonucunu hızlı okumak için çoklu olasılık paneli."""
    def __init__(self, parent, height=188, **kw):
        super().__init__(parent, height=height, bg=C["bg2"], highlightthickness=0, **kw)
        self._stats = None
        self._confidence = 0.0
        self._note = ""
        self.bind("<Configure>", lambda _event: self._redraw())

    def update_analysis(self, stats=None, confidence=0.0, note=""):
        self._stats = stats
        self._confidence = confidence or 0.0
        self._note = note or ""
        self._redraw()

    def _bar(self, x, y, w, label, pct, color):
        pct = max(0.0, min(100.0, float(pct or 0.0)))
        self.create_text(x, y, text=label, fill=C["text2"], anchor="nw", font=("Segoe UI", 9, "bold"))
        self.create_text(x + w, y, text=f"%{pct:.0f}", fill=C["text1"], anchor="ne", font=("Segoe UI", 9, "bold"))
        track_y = y + 21
        self.create_rectangle(x, track_y, x + w, track_y + 9, fill=C["bg3"], outline="")
        self.create_rectangle(x, track_y, x + max(2, w * pct / 100.0), track_y + 9, fill=color, outline="")

    def _score_chips(self, x, y, items):
        if not items:
            self.create_text(x, y, text="Skor dağılımı yok", fill=C["text3"], anchor="nw", font=("Segoe UI", 9))
            return
        chip_x = x
        for item in items:
            text = f"{item['score']}  %{item['pct']:.0f}"
            chip_w = max(70, 10 + len(text) * 7)
            self.create_rectangle(chip_x, y, chip_x + chip_w, y + 26, fill=C["bg3"], outline=C["border"], width=1)
            self.create_text(chip_x + chip_w / 2, y + 13, text=text, fill=C["text1"],
                             font=("Segoe UI", 9, "bold"))
            chip_x += chip_w + 8

    def _redraw(self):
        self.delete("all")
        w = self.winfo_width() or 760
        if not self._stats:
            self.create_text(8, 18, text="Analiz grafiği için önce eşleştirme yap.",
                             fill=C["text3"], anchor="nw", font=("Segoe UI", 10))
            return

        stats = self._stats
        margin = 8
        left_w = max(260, int(w * 0.46))
        right_x = left_w + 28
        right_w = max(230, w - right_x - margin)
        ms = [
            ("MS 1", stats["home_pct"], C["green"]),
            ("MS X", stats["draw_pct"], C["yellow"]),
            ("MS 2", stats["away_pct"], C["red"]),
        ]
        x = margin
        y = 24
        self.create_text(x, 2, text="Maç sonucu", fill=C["text2"], anchor="nw", font=("Segoe UI", 9, "bold"))
        for label, pct, color in ms:
            seg_w = max(1, left_w * pct / 100.0)
            self.create_rectangle(x, y, x + seg_w, y + 30, fill=color, outline="")
            if seg_w >= 44:
                self.create_text(x + seg_w / 2, y + 15, text=f"{label} %{pct:.0f}",
                                 fill=C["bg0"], font=("Segoe UI", 9, "bold"))
            x += seg_w

        self.create_text(margin, 67, text=f"Ortalama gol: {stats['avg_goals']:.2f}",
                         fill=C["text1"], anchor="nw", font=("Segoe UI", 10, "bold"))
        data_note = f"{stats['n']} skorlu maç"
        if self._note:
            data_note += f" | {self._note}"
        self.create_text(margin, 88, text=data_note, fill=C["text3"], anchor="nw", font=("Segoe UI", 9))

        self.create_text(margin, 118, text="En sık skorlar", fill=C["text2"], anchor="nw",
                         font=("Segoe UI", 9, "bold"))
        self._score_chips(margin, 140, stats.get("top_ft_scores", []))

        self._bar(right_x, 2, right_w, "İstatistik desteği (garanti değil)", self._confidence, self._confidence_color())
        self._bar(right_x, 45, right_w, "2.5 Üst", stats["over25_pct"], C["cyan"])
        self._bar(right_x, 88, right_w, "KG Var", stats["btts_pct"], C["purple"])
        half_w = (right_w - 12) / 2
        self._bar(right_x, 131, half_w, "Ev 1.5+", stats["home_over15_pct"], C["green"])
        self._bar(right_x + half_w + 12, 131, half_w, "Dep 1.5+", stats["away_over15_pct"], C["red"])

    def _confidence_color(self):
        if self._confidence >= 75:
            return C["green"]
        if self._confidence >= 55:
            return C["blue"]
        return C["yellow"]


# ─────────────────────────────────────────────────────────────
#  ANA UYGULAMA
# ─────────────────────────────────────────────────────────────
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("İddaa Oran Eşleştirici — Tahmin Arayüzü")
        self.root.geometry("1560x960")
        self.root.minsize(1280, 800)
        self.root.configure(fg_color=C["bg0"])

        self.settings = load_json(SETTINGS_FILE, {})
        self.predictions = load_json(PREDICTIONS_FILE, [])
        self.df = pd.DataFrame()
        self.data_path = DATA_FILE
        self.active_df = pd.DataFrame()
        self.browse_filtered = pd.DataFrame()
        self.today_df = pd.DataFrame()
        self.coupon_df = pd.DataFrame()
        self._last_result = None  # kaydedilecek son tahmin
        default_scope = self.settings.get("veri_kapsami", DATA_SCOPE_ALL)
        if default_scope not in DATA_SCOPES:
            default_scope = DATA_SCOPE_ALL
        self.data_scope_var = tk.StringVar(value=default_scope)
        default_risk = self.settings.get("kupon_risk", "Dengeli")
        if default_risk not in COUPON_RISK_PROFILES:
            default_risk = "Dengeli"
        self.coupon_risk_var = tk.StringVar(value=default_risk)
        coupon_settings = self.settings.get("kupon_ayarlari", {})
        if not isinstance(coupon_settings, dict):
            coupon_settings = {}
        self.coupon_scope_var = tk.StringVar(
            value=coupon_settings.get("kaynak", "Sadece İncele")
        )
        if self.coupon_scope_var.get() not in ("Sadece İncele", "Tüm oranlı tahminler"):
            self.coupon_scope_var.set("Sadece İncele")
        self.coupon_same_league_var = tk.BooleanVar(
            value=bool(coupon_settings.get("lig_cesitliligi", True))
        )
        self.coupon_tier_var = tk.StringVar(value=coupon_settings.get("grup", "Tümü"))
        if self.coupon_tier_var.get() not in ("Tümü", "Banko", "Riskli", "Ya Tutarsa"):
            self.coupon_tier_var.set("Tümü")
        self.coupon_saved_settings = coupon_settings
        default_match_scope = self.settings.get("eslesme_kapsami", MATCH_SCOPE_SMART)
        if default_match_scope not in MATCH_SCOPES:
            default_match_scope = MATCH_SCOPE_SMART
        self.match_scope_var = tk.StringVar(value=default_match_scope)
        default_today_sort = self.settings.get("bugun_siralama", TODAY_SORT_CONFIDENCE)
        if default_today_sort not in TODAY_SORT_OPTIONS:
            default_today_sort = TODAY_SORT_CONFIDENCE
        self.today_sort_var = tk.StringVar(value=default_today_sort)
        self.today_league_var = tk.StringVar(value=self.settings.get("bugun_lig_filtresi", "Tüm Ligler"))
        self.performance_market_var = tk.StringVar(value="Tüm pazarlar")
        self.performance_confidence_var = tk.StringVar(value="Tüm güvenler")
        self.performance_range_var = tk.StringVar(value="Tüm zamanlar")

        self._build_header()
        self._build_tabs()
        self._configure_tree_style()
        self._load_data(silent=True)
        self._restore_today_display()
        self._refresh_history_tree()

    # ---------- Genel yapı ----------
    def _build_header(self):
        header = ctk.CTkFrame(self.root, fg_color=C["bg1"], corner_radius=0, height=64)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)

        left = ctk.CTkFrame(header, fg_color="transparent")
        left.pack(side="left", padx=20, pady=10)
        ctk.CTkLabel(left, text="🎯  İddaa Oran Eşleştirici", font=ctk.CTkFont("Segoe UI", 18, "bold"),
                     text_color=C["text1"]).pack(anchor="w")
        ctk.CTkLabel(left, text="Oranını gir, geçmiş verilerle eşleştir, tahminini kaydet",
                     font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"]).pack(anchor="w")

        right = ctk.CTkFrame(header, fg_color="transparent")
        right.pack(side="right", padx=20, pady=10)
        ctk.CTkLabel(right, text="Veri:", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text3"]).pack(side="left", padx=(0, 6))
        self.data_scope_menu = ctk.CTkOptionMenu(
            right,
            values=DATA_SCOPES,
            variable=self.data_scope_var,
            width=132,
            height=30,
            fg_color=C["bg3"],
            button_color=C["bg4"],
            button_hover_color=C["bg5"],
            text_color=C["text1"],
            command=lambda _value: self._apply_data_filter(silent=False),
        )
        self.data_scope_menu.pack(side="left", padx=(0, 10))
        self.status_lbl = ctk.CTkLabel(right, text="Veri yükleniyor...", font=ctk.CTkFont("Segoe UI", 11),
                                        text_color=C["text2"])
        self.status_lbl.pack(side="left", padx=(0, 14))
        ctk.CTkButton(right, text="📁 Excel Seç", width=110, height=30, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=self._pick_file).pack(side="left", padx=4)
        self.update_data_btn = ctk.CTkButton(right, text="Veriyi Güncelle", width=120, height=30, corner_radius=8,
                                             fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                                             command=self._update_missing_data)
        self.update_data_btn.pack(side="left", padx=4)
        ctk.CTkButton(right, text="🔄 Yenile", width=100, height=30, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=lambda: self._load_data(silent=False)).pack(side="left", padx=4)

        strip = ctk.CTkFrame(self.root, fg_color=C["bg0"], corner_radius=0, height=38)
        strip.pack(fill="x", side="top")
        strip.pack_propagate(False)
        inner = ctk.CTkFrame(strip, fg_color="transparent")
        inner.pack(fill="x", padx=20, pady=4)
        self.header_chips = {}
        for key, title in (("source", "KAYNAK"), ("rows", "KAYIT"), ("scope", "KAPSAM"), ("quota", "CANLI API")):
            chip = StatusChip(inner, title); chip.pack(side="left", padx=(0, 8)); self.header_chips[key] = chip

    def _build_tabs(self):
        shell = ctk.CTkFrame(self.root, fg_color=C["bg0"], corner_radius=0)
        shell.pack(fill="both", expand=True, padx=14, pady=(10, 14))
        sidebar = ctk.CTkFrame(shell, width=215, fg_color=C["bg1"], corner_radius=16,
                               border_width=1, border_color=C["border"])
        sidebar.pack(side="left", fill="y", padx=(0, 12)); sidebar.pack_propagate(False)
        ctk.CTkLabel(sidebar, text="ÇALIŞMA ALANI", font=ctk.CTkFont("Segoe UI", 10, "bold"), text_color=C["text3"]).pack(anchor="w", padx=18, pady=(22, 8))
        content = ctk.CTkFrame(shell, fg_color="transparent")
        content.pack(side="left", fill="both", expand=True)
        self.pages, self.nav_buttons = {}, {}
        page_specs = [
            ("match", "🎯", "Oran Eşleştirme", "Geçmiş benzerliklerden olasılık üret"),
            ("today", "🟢", "Bugünün Maçları", "Fikstür ve oranları incele; tahmin için maçı eşleştirmeye gönder"),
            ("live", "📡", "Canlı Sinyaller", "Baskı ve oran verisiyle canlı karar desteği"),
            ("coupon", "🎟", "Kupon Önerisi", "Risk profiline göre seçimler"),
            ("history", "🕘", "Tahmin Geçmişi", "Kaydedilen tahmin sonuçları"),
            ("browse", "📊", "Veri Gözat", "Geçmiş oran arşivini ara"),
            ("health", "🩺", "Veri Sağlığı", "Veri kalitesini kontrol et"),
            ("settings", "⚙️", "Ayarlar", "Eşleştirme ve kupon tercihlerini düzenle"),
        ]
        builders = {"match": self._build_match_tab, "today": self._build_today_tab, "live": self._build_live_tab,
                    "coupon": self._build_coupon_tab, "history": self._build_history_tab, "browse": self._build_browse_tab,
                    "health": self._build_health_tab, "settings": self._build_settings_tab}
        for key, icon, title, subtitle in page_specs:
            button = ctk.CTkButton(sidebar, text=f"{icon}  {title}", anchor="w", height=42, corner_radius=9,
                                   fg_color="transparent", hover_color=C["bg3"], text_color=C["text2"],
                                   font=ctk.CTkFont("Segoe UI", 12, "bold"), command=lambda value=key: self._show_page(value))
            button.pack(fill="x", padx=10, pady=2); self.nav_buttons[key] = button
            page = ctk.CTkFrame(content, fg_color="transparent"); self.pages[key] = page
            heading = ctk.CTkFrame(page, fg_color="transparent"); heading.pack(fill="x", padx=4, pady=(2, 10))
            ctk.CTkLabel(heading, text=title, font=ctk.CTkFont("Segoe UI", 20, "bold"), text_color=C["text1"]).pack(anchor="w")
            ctk.CTkLabel(heading, text=subtitle, font=ctk.CTkFont("Segoe UI", 11), text_color=C["text3"]).pack(anchor="w", pady=(1, 0))
            body = ctk.CTkFrame(page, fg_color="transparent"); body.pack(fill="both", expand=True)
            builders[key](body)
        ctk.CTkLabel(sidebar, text="Canlı karar destek aracı\nKesin sonuç garantisi vermez.", justify="left", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"]).pack(side="bottom", anchor="w", padx=18, pady=18)
        self._show_page("match")

    def _show_page(self, key):
        for page_key, page in self.pages.items():
            page.pack_forget()
            self.nav_buttons[page_key].configure(fg_color="transparent", text_color=C["text2"])
        self.pages[key].pack(fill="both", expand=True)
        self.nav_buttons[key].configure(fg_color=C["accent"], text_color="white")

    def _configure_tree_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("Dark.Treeview", background=C["bg2"], fieldbackground=C["bg2"],
                        foreground=C["text1"], rowheight=30, borderwidth=0, font=("Segoe UI", 10))
        style.configure("Dark.Treeview.Heading", background=C["bg3"], foreground=C["text1"],
                        relief="flat", font=("Segoe UI", 9, "bold"))
        style.map("Dark.Treeview", background=[("selected", C["accent"])],
                  foreground=[("selected", C["bg0"])])
        style.configure("Dark.Vertical.TScrollbar", background=C["bg3"], troughcolor=C["bg1"])

    # ---------- Sekme 1: Oran Eşleştirme ----------
    def _build_match_tab(self, parent):
        parent.grid_columnconfigure(0, weight=0)
        parent.grid_columnconfigure(1, weight=1)
        parent.grid_rowconfigure(0, weight=1)

        left = ctk.CTkScrollableFrame(parent, fg_color="transparent", width=430)
        left.grid(row=0, column=0, sticky="ns", padx=(4, 10), pady=4)
        right = ctk.CTkFrame(parent, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew", pady=4)

        # --- Maç bilgisi kartı ---
        info_card = ctk.CTkFrame(left, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        info_card.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(info_card, text="⚽  Maç Bilgisi (opsiyonel)", font=ctk.CTkFont("Segoe UI", 12, "bold"),
                     text_color=C["text1"]).pack(anchor="w", padx=14, pady=(12, 6))
        row1 = ctk.CTkFrame(info_card, fg_color="transparent"); row1.pack(fill="x", padx=14, pady=(0, 6))
        self.ev_entry = ctk.CTkEntry(row1, placeholder_text="Ev Sahibi", height=30, fg_color=C["bg3"],
                                      border_color=C["border"], text_color=C["text1"])
        self.ev_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))
        self.dep_entry = ctk.CTkEntry(row1, placeholder_text="Deplasman", height=30, fg_color=C["bg3"],
                                       border_color=C["border"], text_color=C["text1"])
        self.dep_entry.pack(side="left", fill="x", expand=True)
        row2 = ctk.CTkFrame(info_card, fg_color="transparent"); row2.pack(fill="x", padx=14, pady=(0, 14))
        self.lig_entry = ctk.CTkEntry(row2, placeholder_text="Lig (opsiyonel)", height=30, fg_color=C["bg3"],
                                       border_color=C["border"], text_color=C["text1"])
        self.lig_entry.pack(fill="x")

        # --- Oranlar kartı ---
        odds_card = ctk.CTkFrame(left, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        odds_card.pack(fill="x", pady=(0, 10))
        hdr = ctk.CTkFrame(odds_card, fg_color="transparent"); hdr.pack(fill="x", padx=14, pady=(12, 4))
        ctk.CTkLabel(hdr, text="📈  Elindeki Oranlar", font=ctk.CTkFont("Segoe UI", 12, "bold"),
                     text_color=C["text1"]).pack(side="left")
        ctk.CTkButton(hdr, text="Temizle", width=70, height=24, corner_radius=6,
                      fg_color=C["bg4"], hover_color=C["bg5"], text_color=C["text2"],
                      command=self._clear_odds).pack(side="right")

        self.odds_widgets = {}
        for group_title, keys in FIELD_GROUPS:
            ctk.CTkLabel(odds_card, text=group_title, font=ctk.CTkFont("Segoe UI", 10, "bold"),
                         text_color=C["text3"]).pack(anchor="w", padx=14, pady=(8, 2))
            grid = ctk.CTkFrame(odds_card, fg_color="transparent"); grid.pack(fill="x", padx=14)
            for i, key in enumerate(keys):
                f = FIELD_BY_KEY[key]
                w = OddsField(grid, key, f["label"], f["color"], on_enter=self._run_match)
                w.grid(row=i // 2, column=i % 2, padx=(0, 10), pady=3, sticky="w")
                self.odds_widgets[key] = w
        ctk.CTkFrame(odds_card, fg_color="transparent", height=8).pack()

        ctk.CTkButton(left, text="⚙️  Eşleştirme ayarlarını aç", height=32, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=lambda: self._show_page("settings")).pack(fill="x", pady=(0, 10))

        ctk.CTkButton(left, text="🔎  ORANLARI EŞLEŞTİR", height=44, corner_radius=10,
                      fg_color=C["accent"], hover_color="#3D74E0", text_color="white",
                      font=ctk.CTkFont("Segoe UI", 13, "bold"),
                      command=self._run_match).pack(fill="x", pady=(2, 4))
        self.match_status = ctk.CTkLabel(left, text="", font=ctk.CTkFont("Segoe UI", 10),
                                         text_color=C["text3"])
        self.match_status.pack(anchor="w", pady=(2, 0))

        # --- Sağ taraf: sonuçlar ---
        self._build_match_results(right)

    def _build_match_results(self, parent):
        parent.grid_rowconfigure(3, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        pills_row = ctk.CTkFrame(parent, fg_color="transparent")
        pills_row.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        pills_row.grid_columnconfigure((0, 1, 2, 3), weight=1)
        self.pill_count = Pill(pills_row, "EŞLEŞEN MAÇ"); self.pill_count.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.pill_pred = Pill(pills_row, "TAHMİN"); self.pill_pred.grid(row=0, column=1, sticky="ew", padx=8)
        self.pill_conf = Pill(pills_row, "İSTATİSTİK DESTEĞİ"); self.pill_conf.grid(row=0, column=2, sticky="ew", padx=8)
        self.pill_score = Pill(pills_row, "EN SIK SKOR"); self.pill_score.grid(row=0, column=3, sticky="ew", padx=(8, 0))

        prob_card = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        prob_card.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        ctk.CTkLabel(prob_card, text="Analiz Sonuç Grafiği", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C["text2"]).pack(anchor="w", padx=14, pady=(10, 6))
        self.prob_bar = ProbBar(prob_card)
        self.prob_bar.pack(fill="x", padx=14, pady=(0, 6))
        self.analysis_graph = AnalysisGraph(prob_card)
        self.analysis_graph.pack(fill="x", padx=14, pady=(2, 8))

        chips_row = ctk.CTkFrame(prob_card, fg_color="transparent"); chips_row.pack(fill="x", padx=14, pady=(0, 12))
        self.chip_over = ctk.CTkLabel(chips_row, text="2.5 Üst: —", font=ctk.CTkFont("Segoe UI", 10),
                                       text_color=C["text2"]); self.chip_over.pack(side="left", padx=(0, 16))
        self.chip_btts = ctk.CTkLabel(chips_row, text="KG Var: —", font=ctk.CTkFont("Segoe UI", 10),
                                       text_color=C["text2"]); self.chip_btts.pack(side="left", padx=(0, 16))
        self.chip_goals = ctk.CTkLabel(chips_row, text="Ort. Gol: —", font=ctk.CTkFont("Segoe UI", 10),
                                        text_color=C["text2"]); self.chip_goals.pack(side="left", padx=(0, 16))
        self.chip_ht = ctk.CTkLabel(chips_row, text="En Sık İY: —", font=ctk.CTkFont("Segoe UI", 10),
                                     text_color=C["text2"]); self.chip_ht.pack(side="left")
        self.chip_team_goals = ctk.CTkLabel(chips_row, text="Ev/Dep 1.5: —", font=ctk.CTkFont("Segoe UI", 10),
                                             text_color=C["text2"]); self.chip_team_goals.pack(side="left", padx=(16, 0))

        self.combo_lbl = ctk.CTkLabel(parent, text="Henüz eşleştirme yapılmadı.",
                                       font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"],
                                       anchor="w", justify="left", wraplength=900)
        self.combo_lbl.grid(row=2, column=0, sticky="ew", pady=(0, 8))

        table_card = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        table_card.grid(row=3, column=0, sticky="nsew")
        table_card.grid_rowconfigure(1, weight=1)
        table_card.grid_columnconfigure(0, weight=1)
        top_bar = ctk.CTkFrame(table_card, fg_color="transparent"); top_bar.grid(row=0, column=0, sticky="ew", padx=14, pady=(10, 4))
        ctk.CTkLabel(top_bar, text="Eşleşen Maçlar", font=ctk.CTkFont("Segoe UI", 11, "bold"),
                     text_color=C["text2"]).pack(side="left")
        self.save_btn = ctk.CTkButton(top_bar, text="💾  Bu Tahmini Kaydet", height=30, corner_radius=8,
                                       fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                                       state="disabled", command=self._save_prediction)
        self.save_btn.pack(side="right")

        tree_frame = ctk.CTkFrame(table_card, fg_color="transparent")
        tree_frame.grid(row=1, column=0, sticky="nsew", padx=14, pady=(0, 6))
        self.match_tree = ttk.Treeview(tree_frame, show="headings", style="Dark.Treeview")
        vscroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.match_tree.yview, style="Dark.Vertical.TScrollbar")
        self.match_tree.configure(yscrollcommand=vscroll.set)
        self.match_tree.pack(side="left", fill="both", expand=True)
        vscroll.pack(side="right", fill="y")

        log_card = ctk.CTkFrame(parent, fg_color=C["bg1"], corner_radius=10)
        log_card.grid(row=4, column=0, sticky="ew", pady=(8, 0))
        self.log_text = tk.Text(log_card, height=5, bg=C["bg1"], fg=C["text3"], relief="flat",
                                 font=("Consolas", 9), state="disabled", wrap="word", padx=10, pady=6)
        self.log_text.pack(fill="x")

    # ---------- Sekme 2: Bugunun Maclari ----------
    def _build_today_tab(self, parent):
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Bugunun Oynanmamis Maclari", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text1"]).pack(side="left")
        self.today_status = ctk.CTkLabel(top, text="Maclar henuz cekilmedi.",
                                          font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.today_status.pack(side="left", padx=14)
        ctk.CTkLabel(top, text="Sırala", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(side="left", padx=(6, 6))
        self.today_sort_menu = ctk.CTkOptionMenu(
            top,
            values=TODAY_SORT_OPTIONS,
            variable=self.today_sort_var,
            width=92,
            height=28,
            fg_color=C["bg3"],
            button_color=C["bg4"],
            button_hover_color=C["bg5"],
            text_color=C["text1"],
            command=lambda value: self._sort_today_matches(value),
        )
        self.today_sort_menu.pack(side="left")
        ctk.CTkLabel(top, text="Lig", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(side="left", padx=(14, 6))
        self.today_league_menu = ctk.CTkOptionMenu(
            top,
            values=["Tüm Ligler"],
            variable=self.today_league_var,
            width=170,
            height=28,
            fg_color=C["bg3"],
            button_color=C["bg4"],
            button_hover_color=C["bg5"],
            text_color=C["text1"],
            command=self._set_today_league_filter,
        )
        self.today_league_menu.pack(side="left")
        self.today_refresh_btn = ctk.CTkButton(
            top, text="Bugunun Maclarini Cek", width=165, height=30, corner_radius=8,
            fg_color=C["accent"], hover_color="#3D74E0", text_color="white",
            command=self._load_today_matches,
        )
        self.today_refresh_btn.pack(side="right")
        ctk.CTkButton(top, text="Kupon Oluştur", width=120, height=30, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=self._generate_coupon).pack(side="right", padx=(0, 8))
        ctk.CTkButton(top, text="Tahmini Takip Et", width=125, height=30, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=self._track_selected_today_prediction).pack(side="right", padx=(0, 8))

        ctk.CTkLabel(parent, text="Bir maça çift tıklayarak oranlarını Eşleştirme ekranına aktarabilir ve tahmini orada çalıştırabilirsin.",
                     font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"]).grid(
                         row=2, column=0, sticky="w", padx=8, pady=(0, 4)
                     )

        frame = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)

        self.today_tree = ttk.Treeview(frame, columns=TODAY_TREE_COLUMNS, show="headings", style="Dark.Treeview")
        titles = {
            "Saat": "Saat", "Lig": "Lig", "Maç": "Maç", "MS 1": "MS 1", "MS X": "MS X", "MS 2": "MS 2",
            "Tahmin": "Tahmin", "Oran": "Oran", "Model": "Model", "Piyasa": "Piyasa", "Grup": "Grup", "Karar": "Karar", "Eşleşme": "Eşleşme",
        }
        widths = {"Saat": 60, "Lig": 180, "Maç": 260, "MS 1": 65, "MS X": 65, "MS 2": 65,
                  "Tahmin": 110, "Oran": 65, "Model": 70, "Piyasa": 75, "Grup": 100, "Karar": 80, "Eşleşme": 75}
        for column in TODAY_TREE_COLUMNS:
            self.today_tree.heading(column, text=titles[column])
            self.today_tree.column(column, width=widths[column], anchor="w")
        vscroll = ttk.Scrollbar(frame, orient="vertical", command=self.today_tree.yview, style="Dark.Vertical.TScrollbar")
        self.today_tree.configure(yscrollcommand=vscroll.set)
        self.today_tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vscroll.grid(row=0, column=1, sticky="ns", pady=12, padx=(0, 12))
        self.today_tree.bind("<Double-1>", self._send_today_to_match)

    # ---------- Sekme 3: Canli Sinyaller ----------
    def _build_live_tab(self, parent):
        parent.grid_rowconfigure(2, weight=1)
        parent.grid_columnconfigure(0, weight=1)
        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Canli Gol Sinyalleri", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text1"]).pack(side="left")
        self.live_status = ctk.CTkLabel(top, text="Sinyal taramasi baslatilmadi.",
                                         font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.live_status.pack(side="left", padx=14)
        self.live_refresh_btn = ctk.CTkButton(
            top, text="Canli Maclari Tara", width=150, height=30, corner_radius=8,
            fg_color=C["accent"], hover_color="#3D74E0", text_color="white", command=self._refresh_live_signals,
        )
        self.live_refresh_btn.pack(side="right")
        ctk.CTkButton(top, text="Sonuçları Güncelle", width=135, height=30, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=self._settle_live_signals).pack(side="right", padx=(0, 8))
        metrics = ctk.CTkFrame(parent, fg_color="transparent")
        metrics.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 8))
        metrics.grid_columnconfigure((0, 1, 2, 3, 4, 5), weight=1)
        self.live_pills = []
        for index, title in enumerate(("CANLI MAÇ", "ADAY MAÇ", "SİNYAL", "BAHİS UYGUN", "API KOTA", "BAŞARI")):
            pill = Pill(metrics, title); pill.grid(row=0, column=index, sticky="ew", padx=4); self.live_pills.append(pill)
        self.live_cards = ctk.CTkScrollableFrame(parent, fg_color=C["bg1"], corner_radius=12,
                                                   border_width=1, border_color=C["border"])
        self.live_cards.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 4))

    # ---------- Sekme 4: Kupon Onerisi ----------
    def _build_coupon_tab(self, parent):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Yuksek Guvenli Kupon", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text1"]).pack(side="left")
        ctk.CTkLabel(top, text="Risk", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(side="left", padx=(18, 6))
        self.coupon_risk_menu = ctk.CTkOptionMenu(
            top,
            values=list(COUPON_RISK_PROFILES),
            variable=self.coupon_risk_var,
            width=92,
            height=28,
            fg_color=C["bg3"],
            button_color=C["bg4"],
            button_hover_color=C["bg5"],
            text_color=C["text1"],
            command=lambda value: self._apply_coupon_risk(value),
        )
        self.coupon_risk_menu.pack(side="left")
        ctk.CTkLabel(top, text="Min. Guven %", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(side="left", padx=(24, 6))
        self.coupon_min_conf_entry = ctk.CTkEntry(top, width=54, height=28, fg_color=C["bg3"],
                                                   border_color=C["border"], text_color=C["text1"])
        profile = COUPON_RISK_PROFILES.get(self.coupon_risk_var.get(), COUPON_RISK_PROFILES["Dengeli"])
        self.coupon_min_conf_entry.insert(0, str(self.coupon_saved_settings.get("min_guven", profile.get("min_confidence", 60))))
        self.coupon_min_conf_entry.pack(side="left")
        ctk.CTkLabel(top, text="Maks. Secim", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(side="left", padx=(14, 6))
        self.coupon_max_selection_entry = ctk.CTkEntry(top, width=48, height=28, fg_color=C["bg3"],
                                                        border_color=C["border"], text_color=C["text1"])
        self.coupon_max_selection_entry.insert(0, str(self.coupon_saved_settings.get("maks_secim", profile.get("max_selections", 5))))
        self.coupon_max_selection_entry.pack(side="left")
        ctk.CTkLabel(top, text="Tutar", font=ctk.CTkFont("Segoe UI", 10),
                     text_color=C["text2"]).pack(side="left", padx=(14, 6))
        self.coupon_stake_entry = ctk.CTkEntry(top, width=64, height=28, fg_color=C["bg3"],
                                               border_color=C["border"], text_color=C["text1"])
        self.coupon_stake_entry.insert(0, str(self.settings.get("kupon_tutar", 100)))
        self.coupon_stake_entry.pack(side="left")
        self.coupon_stake_entry.bind("<Return>", lambda _event: self._update_coupon_simulation())
        self.coupon_stake_entry.bind("<FocusOut>", lambda _event: self._update_coupon_simulation())
        self.coupon_status = ctk.CTkLabel(top, text="Bugunun maclarini cekerek kupon olusturabilirsin.",
                                           font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.coupon_status.pack(side="left", padx=14)
        ctk.CTkButton(top, text="Kuponu Guncelle", width=135, height=30, corner_radius=8,
                      fg_color=C["accent"], hover_color="#3D74E0", text_color="white",
                      command=self._generate_coupon).pack(side="right")
        ctk.CTkButton(top, text="Seçimi Takip Et", width=120, height=30, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=self._track_selected_coupon_prediction).pack(side="right", padx=(0, 8))

        advanced = ctk.CTkFrame(parent, fg_color="transparent")
        advanced.grid(row=1, column=0, sticky="ew", padx=4, pady=(0, 8))
        ctk.CTkLabel(advanced, text="Kaynak", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(side="left", padx=(4, 6))
        self.coupon_scope_menu = ctk.CTkOptionMenu(
            advanced, values=["Sadece İncele", "Tüm oranlı tahminler"], variable=self.coupon_scope_var,
            width=155, height=27, fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"], text_color=C["text1"],
        )
        self.coupon_scope_menu.pack(side="left")
        ctk.CTkLabel(advanced, text="Grup", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(side="left", padx=(18, 6))
        self.coupon_tier_menu = ctk.CTkOptionMenu(
            advanced, values=["Tümü", "Banko", "Riskli", "Ya Tutarsa"], variable=self.coupon_tier_var,
            width=115, height=27, fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"], text_color=C["text1"],
        )
        self.coupon_tier_menu.pack(side="left")
        ctk.CTkLabel(advanced, text="Zaman aralığı dk", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).pack(side="left", padx=(18, 6))
        self.coupon_gap_entry = ctk.CTkEntry(advanced, width=54, height=27, fg_color=C["bg3"], border_color=C["border"], text_color=C["text1"])
        self.coupon_gap_entry.insert(0, str(self.coupon_saved_settings.get("zaman_araligi", profile.get("min_gap", 60))))
        self.coupon_gap_entry.pack(side="left")
        self.coupon_same_league_check = ctk.CTkCheckBox(
            advanced, text="Lig çeşitliliği", variable=self.coupon_same_league_var,
            font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"], fg_color=C["accent"], hover_color="#3D74E0",
        )
        self.coupon_same_league_check.pack(side="left", padx=(18, 0))

        frame = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        frame.grid(row=2, column=0, sticky="nsew", padx=4, pady=(0, 4))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        columns = ("Saat", "Lig", "Maç", "Seçim", "Grup", "Oran", "Güven", "Kalite", "Eşleşme")
        widths = {"Saat": 60, "Lig": 180, "Maç": 300, "Seçim": 140, "Grup": 100, "Oran": 70, "Güven": 85, "Kalite": 95, "Eşleşme": 90}
        self.coupon_tree = ttk.Treeview(frame, columns=columns, show="headings", style="Dark.Treeview")
        for column in columns:
            self.coupon_tree.heading(column, text=column)
            self.coupon_tree.column(column, width=widths[column], anchor="w")
        vscroll = ttk.Scrollbar(frame, orient="vertical", command=self.coupon_tree.yview, style="Dark.Vertical.TScrollbar")
        self.coupon_tree.configure(yscrollcommand=vscroll.set)
        self.coupon_tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vscroll.grid(row=0, column=1, sticky="ns", pady=12, padx=(0, 12))
        summary = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=10, border_width=1, border_color=C["border"])
        summary.grid(row=3, column=0, sticky="ew", padx=4, pady=(8, 0))
        self.coupon_sim_lbl = ctk.CTkLabel(summary, text="Simülasyon: kupon oluşunca hesaplanır.",
                                           font=ctk.CTkFont("Segoe UI", 11), text_color=C["text2"])
        self.coupon_sim_lbl.pack(anchor="w", padx=14, pady=10)
        self._apply_coupon_risk(self.coupon_risk_var.get(), regenerate=False)

    # ---------- Sekme 2: Tahmin Geçmişi ----------
    def _build_history_tab(self, parent):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(parent, fg_color="transparent"); top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Kaydedilen Tahminler", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text1"]).pack(side="left")
        ctk.CTkButton(top, text="🔄 Yenile", width=90, height=28, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=self._refresh_history_tree).pack(side="right", padx=4)
        ctk.CTkButton(top, text="🗑 Seçileni Sil", width=110, height=28, corner_radius=8,
                      fg_color=C["red_dim"], hover_color=C["red"], text_color=C["text1"],
                      command=self._delete_selected_history).pack(side="right", padx=4)
        ctk.CTkButton(top, text="📤 Excel'e Aktar", width=120, height=28, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=self._export_history).pack(side="right", padx=4)
        ctk.CTkButton(top, text="📊 Performans Merkezi", width=150, height=28, corner_radius=8,
                      fg_color=C["blue_dim"], hover_color=C["accent"], text_color=C["text1"],
                      command=self._show_performance_center).pack(side="right", padx=4)
        ctk.CTkButton(top, text="Sonuçları Güncelle", width=145, height=28, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=self._settle_tracked_predictions).pack(side="right", padx=4)
        for values, variable in ((["Tüm zamanlar", "Son 7 gün", "Son 30 gün"], self.performance_range_var),
                                 (["Tüm güvenler", "<60", "60-74", "75+"], self.performance_confidence_var),
                                 (["Tüm pazarlar", "MS 1", "MS X", "MS 2", "2.5 Üst", "2.5 Alt", "KG Var", "KG Yok"], self.performance_market_var)):
            ctk.CTkOptionMenu(top, values=values, variable=variable, width=105, height=28,
                              fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"],
                              text_color=C["text1"], command=lambda _value: self._refresh_history_stats()).pack(side="right", padx=3)

        frame = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)

        cols = ("ts", "mac", "tahmin", "durum", "guven", "esleme", "tolerans", "kombo")
        titles = {"ts": "Tarih", "mac": "Maç", "tahmin": "Tahmin", "guven": "Güven %",
                  "durum": "Durum", "esleme": "Eşleşme", "tolerans": "Tolerans %", "kombo": "Kombinasyon"}
        widths = {"ts": 120, "mac": 220, "tahmin": 110, "durum": 90, "guven": 80, "esleme": 80, "tolerans": 90, "kombo": 300}
        self.history_tree = ttk.Treeview(frame, columns=cols, show="headings", style="Dark.Treeview")
        for c in cols:
            self.history_tree.heading(c, text=titles[c])
            self.history_tree.column(c, width=widths[c], anchor="w")
        vscroll = ttk.Scrollbar(frame, orient="vertical", command=self.history_tree.yview, style="Dark.Vertical.TScrollbar")
        self.history_tree.configure(yscrollcommand=vscroll.set)
        self.history_tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=12)
        vscroll.grid(row=0, column=1, sticky="ns", pady=12, padx=(0, 12))
        actions = ctk.CTkFrame(parent, fg_color="transparent")
        actions.grid(row=2, column=0, sticky="ew", padx=4, pady=(6, 0))
        self.history_stats_lbl = ctk.CTkLabel(actions, text="", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.history_stats_lbl.pack(side="left")
        self.history_performance_lbl = ctk.CTkLabel(actions, text="", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"])
        self.history_performance_lbl.pack(side="left", padx=16)
        ctk.CTkLabel(actions, text="Yedek sonuç verisi: football-data.org", font=ctk.CTkFont("Segoe UI", 9), text_color=C["text3"]).pack(side="left", padx=10)
        ctk.CTkButton(actions, text="Bekliyor", width=90, height=28, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=lambda: self._set_selected_prediction_status("Bekliyor")).pack(side="right", padx=4)
        ctk.CTkButton(actions, text="Yatti", width=80, height=28, corner_radius=8,
                      fg_color=C["red_dim"], hover_color=C["red"], text_color=C["text1"],
                      command=lambda: self._set_selected_prediction_status("Yatti")).pack(side="right", padx=4)
        ctk.CTkButton(actions, text="Tuttu", width=80, height=28, corner_radius=8,
                      fg_color=C["green_dim"], hover_color=C["green"], text_color=C["text1"],
                      command=lambda: self._set_selected_prediction_status("Tuttu")).pack(side="right", padx=4)

    # ---------- Sekme 3: Veri Gözat ----------
    def _build_browse_tab(self, parent):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(parent, fg_color="transparent"); top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Ara:", font=ctk.CTkFont("Segoe UI", 11), text_color=C["text2"]).pack(side="left")
        self.browse_search_var = tk.StringVar()
        self.browse_search_var.trace_add("write", lambda *_: self._search_browse())
        entry = ctk.CTkEntry(top, textvariable=self.browse_search_var, width=320, height=30,
                              placeholder_text="Takım veya lig adı...", fg_color=C["bg3"],
                              border_color=C["border"], text_color=C["text1"])
        entry.pack(side="left", padx=8)
        ctk.CTkLabel(top, text="Çift tıkla → oranları Eşleştirme sekmesine gönder",
                     font=ctk.CTkFont("Segoe UI", 9), text_color=C["text3"]).pack(side="left", padx=12)
        self.browse_status = ctk.CTkLabel(top, text="", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.browse_status.pack(side="right")

        frame = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        frame.grid_rowconfigure(0, weight=1); frame.grid_columnconfigure(0, weight=1)

        col_widths = self.settings.get("col_widths", {})
        cols = [title for title, _, _ in BROWSE_COLUMNS]
        self.browse_tree = ttk.Treeview(frame, columns=cols, show="headings", style="Dark.Treeview")
        for title, _, default_w in BROWSE_COLUMNS:
            self.browse_tree.heading(title, text=title)
            self.browse_tree.column(title, width=col_widths.get(title, default_w), anchor="w")
        vscroll = ttk.Scrollbar(frame, orient="vertical", command=self.browse_tree.yview, style="Dark.Vertical.TScrollbar")
        hscroll = ttk.Scrollbar(frame, orient="horizontal", command=self.browse_tree.xview, style="Dark.Vertical.TScrollbar")
        self.browse_tree.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)
        self.browse_tree.grid(row=0, column=0, sticky="nsew", padx=(12, 0), pady=(12, 0))
        vscroll.grid(row=0, column=1, sticky="ns", pady=(12, 0), padx=(0, 12))
        hscroll.grid(row=1, column=0, sticky="ew", padx=(12, 0), pady=(0, 12))
        self.browse_tree.bind("<Double-1>", self._send_browse_to_match)

    def _build_health_tab(self, parent):
        parent.grid_rowconfigure(1, weight=1)
        parent.grid_columnconfigure(0, weight=1)

        top = ctk.CTkFrame(parent, fg_color="transparent")
        top.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 8))
        ctk.CTkLabel(top, text="Veri Sagligi", font=ctk.CTkFont("Segoe UI", 14, "bold"),
                     text_color=C["text1"]).pack(side="left")
        ctk.CTkButton(top, text="Yenile", width=90, height=28, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=self._refresh_data_health).pack(side="right")

        frame = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        frame.grid(row=1, column=0, sticky="nsew", padx=4, pady=(0, 4))
        frame.grid_rowconfigure(0, weight=1)
        frame.grid_columnconfigure(0, weight=1)
        self.health_text = tk.Text(frame, bg=C["bg2"], fg=C["text1"], relief="flat",
                                   font=("Consolas", 11), state="disabled", wrap="word",
                                   padx=14, pady=14)
        self.health_text.grid(row=0, column=0, sticky="nsew", padx=12, pady=12)

    def _build_settings_tab(self, parent):
        """Kalıcı analiz tercihlerini tek yerde düzenlenebilir tutar."""
        defaults = self.settings.get("tahmin_ayarlari", {})
        coupon = self.settings.get("kupon_ayarlari", {})
        if not isinstance(defaults, dict):
            defaults = {}
        if not isinstance(coupon, dict):
            coupon = {}

        card = ctk.CTkFrame(parent, fg_color=C["bg2"], corner_radius=12, border_width=1, border_color=C["border"])
        card.pack(fill="x", padx=4, pady=4)
        ctk.CTkLabel(card, text="Eşleştirme", font=ctk.CTkFont("Segoe UI", 14, "bold"), text_color=C["text1"]).grid(
            row=0, column=0, columnspan=4, sticky="w", padx=18, pady=(16, 8))

        def label(text, row, column):
            ctk.CTkLabel(card, text=text, font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"]).grid(
                row=row, column=column, sticky="w", padx=(18 if column == 0 else 10, 6), pady=6)

        def entry(value, row, column, width=90):
            widget = ctk.CTkEntry(card, width=width, height=30, fg_color=C["bg3"], border_color=C["border"], text_color=C["text1"])
            widget.insert(0, str(value)); widget.grid(row=row, column=column, sticky="w", pady=6)
            return widget

        label("Tolerans (%)", 1, 0)
        self.settings_tolerance_entry = entry(defaults.get("tolerans", FIXED_TOLERANCE), 1, 1)
        label("Min. eşleşme", 1, 2)
        self.settings_min_matches_entry = entry(defaults.get("min_eslesme", FIXED_MINIMUM_MATCHES), 1, 3)
        label("Maç kapsamı", 2, 0)
        ctk.CTkOptionMenu(card, values=MATCH_SCOPES, variable=self.match_scope_var, width=180, height=30,
                          fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"], text_color=C["text1"]).grid(row=2, column=1, sticky="w", pady=6)
        label("Geçmiş veri", 2, 2)
        ctk.CTkOptionMenu(card, values=DATA_SCOPES, variable=self.data_scope_var, width=150, height=30,
                          fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"], text_color=C["text1"]).grid(row=2, column=3, sticky="w", pady=6)

        ctk.CTkLabel(card, text="Kupon", font=ctk.CTkFont("Segoe UI", 14, "bold"), text_color=C["text1"]).grid(
            row=3, column=0, columnspan=4, sticky="w", padx=18, pady=(20, 8))
        label("Min. güven (%)", 4, 0)
        self.settings_coupon_conf_entry = entry(coupon.get("min_guven", 70), 4, 1)
        label("Maks. seçim", 4, 2)
        self.settings_coupon_max_entry = entry(coupon.get("maks_secim", 5), 4, 3)
        label("Zaman aralığı (dk)", 5, 0)
        self.settings_coupon_gap_entry = entry(coupon.get("zaman_araligi", 60), 5, 1)
        label("Kupon tutarı (TL)", 5, 2)
        self.settings_coupon_stake_entry = entry(self.settings.get("kupon_tutar", 100), 5, 3)
        label("Tahmin kaynağı", 6, 0)
        ctk.CTkOptionMenu(card, values=["Sadece İncele", "Tüm oranlı tahminler"], variable=self.coupon_scope_var,
                          width=180, height=30, fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"],
                          text_color=C["text1"]).grid(row=6, column=1, sticky="w", pady=6)
        label("Kupon grubu", 6, 2)
        ctk.CTkOptionMenu(card, values=["Tümü", "Banko", "Riskli", "Ya Tutarsa"], variable=self.coupon_tier_var,
                          width=150, height=30, fg_color=C["bg3"], button_color=C["bg4"], button_hover_color=C["bg5"],
                          text_color=C["text1"]).grid(row=6, column=3, sticky="w", pady=6)
        ctk.CTkCheckBox(card, text="Aynı ligden birden fazla seçim alma", variable=self.coupon_same_league_var,
                        font=ctk.CTkFont("Segoe UI", 10), text_color=C["text2"], fg_color=C["accent"],
                        hover_color="#3D74E0").grid(row=7, column=0, columnspan=3, sticky="w", padx=18, pady=(8, 4))

        self.settings_status = ctk.CTkLabel(card, text="Değişiklikler Kaydet'e basınca kalıcı olur.", font=ctk.CTkFont("Segoe UI", 10), text_color=C["text3"])
        self.settings_status.grid(row=8, column=0, columnspan=3, sticky="w", padx=18, pady=(16, 16))
        ctk.CTkButton(card, text="💾 Ayarları Kaydet", width=150, height=34, corner_radius=8,
                      fg_color=C["accent"], hover_color="#3D74E0", text_color="white",
                      command=self._save_settings_tab).grid(row=8, column=3, sticky="e", padx=18, pady=(16, 16))

    @staticmethod
    def _setting_number(widget, default, minimum, maximum, integer=False):
        value = to_float(widget.get())
        if value is None:
            value = default
        value = max(minimum, min(maximum, value))
        return int(value) if integer else round(value, 1)

    def _save_settings_tab(self):
        tolerance = self._setting_number(self.settings_tolerance_entry, FIXED_TOLERANCE, 0.5, 30.0)
        min_matches = self._setting_number(self.settings_min_matches_entry, FIXED_MINIMUM_MATCHES, 1, 500, integer=True)
        min_confidence = self._setting_number(self.settings_coupon_conf_entry, 70, 0, 75)
        max_selections = self._setting_number(self.settings_coupon_max_entry, 5, 1, 12, integer=True)
        min_gap = self._setting_number(self.settings_coupon_gap_entry, 60, 0, 360, integer=True)
        stake = self._setting_number(self.settings_coupon_stake_entry, 100, 1, 100000)
        self.settings = {**self.settings, "tahmin_ayarlari": {"tolerans": tolerance, "min_eslesme": min_matches},
                         "eslesme_kapsami": self.match_scope_var.get(), "veri_kapsami": self.data_scope_var.get(),
                         "kupon_tutar": str(stake), "kupon_ayarlari": {**self.coupon_saved_settings,
                         "min_guven": min_confidence, "maks_secim": max_selections, "zaman_araligi": min_gap,
                         "kaynak": self.coupon_scope_var.get(), "lig_cesitliligi": self.coupon_same_league_var.get(),
                         "grup": self.coupon_tier_var.get()}}
        self.coupon_saved_settings = self.settings["kupon_ayarlari"]
        save_json(SETTINGS_FILE, self.settings)
        for widget, value in ((getattr(self, "coupon_min_conf_entry", None), min_confidence),
                              (getattr(self, "coupon_max_selection_entry", None), max_selections),
                              (getattr(self, "coupon_gap_entry", None), min_gap),
                              (getattr(self, "coupon_stake_entry", None), stake)):
            if widget is not None:
                widget.delete(0, "end"); widget.insert(0, str(value))
        self._apply_data_filter(silent=True)
        self.settings_status.configure(text="✓ Ayarlar kaydedildi. Yeni eşleştirme ve kuponlarda kullanılacak.", text_color=C["green"])

    # ─────────────────────────────────────────────────────────
    #  VERİ YÜKLEME
    # ─────────────────────────────────────────────────────────
    def _load_data(self, path=None, silent=False):
        path = path or DATA_FILE
        is_default_source = Path(path).resolve() == DATA_FILE.resolve()
        if not os.path.exists(path):
            self.status_lbl.configure(text=f"⚠ Veri dosyası bulunamadı: {path}")
            if not silent:
                messagebox.showwarning("Uyarı", f"{path} bulunamadı. 'Excel Seç' ile konumunu belirtebilirsin.")
            return
        try:
            self.data_path = Path(path)
            self.df = load_odds_data(path)
            source_files = self.df.attrs.get("source_files", [])
            source_name = f"{len(source_files)} Excel dosyası" if source_files else os.path.basename(path)
            self._apply_data_filter(silent=True, filename=source_name)
        except Exception as exc:
            self.status_lbl.configure(text="⚠ Veri yüklenirken hata oluştu")
            messagebox.showerror("Hata", f"Excel okunamadı:\n{exc}")

    def _update_missing_data(self):
        if self.df.empty:
            messagebox.showwarning("Uyarı", "Önce ana veri dosyasını yükle.")
            return
        update_range = missing_update_range(self.df)
        if not update_range:
            messagebox.showinfo("Veri Güncelle", "Veri dosyası düne kadar güncel görünüyor.")
            return
        start, end = update_range
        updates_dir = BASE_DIR / "updates"
        update_file = updates_dir / f"eksik_{start}_{end}.xlsx"
        self._show_update_preview(start, end, update_file)

    def _show_update_preview(self, start, end, update_file):
        last_date = latest_data_date(self.df)
        day_count = inclusive_day_count(start, end)

        window = ctk.CTkToplevel(self.root)
        window.title("Veri Güncelleme Önizleme")
        window.geometry("640x390")
        window.minsize(560, 360)
        window.configure(fg_color=C["bg0"])
        window.transient(self.root)
        window.grab_set()

        body = ctk.CTkFrame(window, fg_color=C["bg2"], corner_radius=12,
                            border_width=1, border_color=C["border"])
        body.pack(fill="both", expand=True, padx=18, pady=18)

        ctk.CTkLabel(body, text="Güncelleme önizlemesi",
                     font=ctk.CTkFont("Segoe UI", 18, "bold"),
                     text_color=C["text1"]).pack(anchor="w", padx=18, pady=(16, 4))
        ctk.CTkLabel(body, text="Başlatmadan önce çekilecek aralığı kontrol et.",
                     font=ctk.CTkFont("Segoe UI", 11),
                     text_color=C["text2"]).pack(anchor="w", padx=18, pady=(0, 12))

        rows_frame = ctk.CTkFrame(body, fg_color="transparent")
        rows_frame.pack(fill="both", expand=True, padx=18, pady=(0, 10))
        rows_frame.grid_columnconfigure(1, weight=1)

        rows = [
            ("Ana dosya", str(self.data_path)),
            ("Son veri tarihi", last_date.strftime("%d-%m-%Y") if last_date else "-"),
            ("Çekilecek aralık", f"{start} - {end}"),
            ("Gün sayısı", str(day_count)),
            ("Ara çıktı", str(update_file)),
            ("Mod", "Kaldığı yerden devam + ara kayıt"),
        ]
        for index, (label, value) in enumerate(rows):
            ctk.CTkLabel(rows_frame, text=label, font=ctk.CTkFont("Segoe UI", 11, "bold"),
                         text_color=C["text2"]).grid(row=index, column=0, sticky="nw", padx=(0, 14), pady=5)
            ctk.CTkLabel(rows_frame, text=value, font=ctk.CTkFont("Segoe UI", 11),
                         text_color=C["text1"], wraplength=410,
                         justify="left").grid(row=index, column=1, sticky="ew", pady=5)

        actions = ctk.CTkFrame(body, fg_color="transparent")
        actions.pack(fill="x", padx=18, pady=(4, 16))
        ctk.CTkButton(actions, text="Vazgeç", width=110, height=34, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=window.destroy).pack(side="right", padx=(8, 0))
        ctk.CTkButton(actions, text="Başlat", width=120, height=34, corner_radius=8,
                      fg_color=C["accent"], hover_color="#3F7EE8", text_color="white",
                      command=lambda: self._confirm_missing_data_update(window, start, end, update_file)).pack(side="right")

        window.after(50, window.focus_force)

    def _confirm_missing_data_update(self, window, start, end, update_file):
        window.destroy()
        self._start_missing_data_update(start, end, update_file)

    def _start_missing_data_update(self, start, end, update_file):
        self.update_data_btn.configure(state="disabled")
        self.status_lbl.configure(text=f"Eksik veri çekiliyor: {start} - {end}")
        threading.Thread(target=self._update_missing_data_worker, args=(start, end, update_file), daemon=True).start()

    def _update_missing_data_worker(self, start, end, update_file):
        try:
            if self.data_path.suffix.lower() == ".sqlite":
                before_rows = len(load_matches(self.data_path))
                iddaa_scraper.main(start, end, self.data_path, workers=8, checkpoint_every=100, resume=True)
                after_rows = len(load_matches(self.data_path))
                result = {
                    "added_rows": max(0, after_rows - before_rows),
                    "updated_rows": 0,
                    "output_rows": after_rows,
                }
                self.root.after(0, lambda result=result: self._finish_missing_data_update(result))
                return
            update_file = Path(update_file)
            updates_dir = update_file.parent
            updates_dir.mkdir(exist_ok=True)
            iddaa_scraper.main(start, end, update_file, workers=8, checkpoint_every=100, resume=True)
            if not update_file.exists():
                raise RuntimeError("Eksik veri dosyası oluşturulamadı.")
            result = merge_excel_files(self.data_path, update_file)
            self.root.after(0, lambda result=result: self._finish_missing_data_update(result))
        except Exception as exc:
            self.root.after(0, lambda message=str(exc): self._fail_missing_data_update(message))

    def _finish_missing_data_update(self, result):
        self.update_data_btn.configure(state="normal")
        self._load_data(path=self.data_path, silent=True)
        messagebox.showinfo(
            "Veri Güncelle",
            "Güncelleme tamamlandı.\n"
            f"Yeni maç: {result.get('added_rows', 0)}\n"
            f"Tekrar/güncelleme: {result.get('updated_rows', 0)}\n"
            f"Toplam satır: {result.get('output_rows', 0)}",
        )

    def _fail_missing_data_update(self, error):
        self.update_data_btn.configure(state="normal")
        self.status_lbl.configure(text="Veri güncelleme tamamlanamadı")
        messagebox.showerror("Veri Güncelle", error)

    def _apply_data_filter(self, silent=False, filename=None):
        if self.df.empty:
            self.active_df = pd.DataFrame()
            self.browse_filtered = pd.DataFrame()
            if hasattr(self, "browse_tree"):
                self._render_browse_tree(self.browse_filtered)
            return

        scope = self.data_scope_var.get()
        self.active_df = filter_data_by_scope(self.df, scope)
        self.browse_filtered = self.active_df.copy()
        if hasattr(self, "browse_tree"):
            self._render_browse_tree(self.browse_filtered)
        if hasattr(self, "health_text"):
            self._refresh_data_health()

        removed = self.df.attrs.get("duplicates_removed", 0)
        suffix = f"; {removed} tekrar kaydi temizlendi" if removed else ""
        source_name = filename or "aktif dosya"
        if scope == DATA_SCOPE_ALL:
            text = f"✓ {len(self.df)} maç yüklendi ({source_name}){suffix}"
        else:
            text = f"✓ {len(self.active_df)} / {len(self.df)} maç aktif ({scope}){suffix}"
        self.status_lbl.configure(text=text)
        if hasattr(self, "header_chips"):
            quota = quota_status()
            self.header_chips["source"].set("SQLite", C["cyan"])
            self.header_chips["rows"].set(f"{len(self.df):,} maç")
            self.header_chips["scope"].set(scope)
            self.header_chips["quota"].set(f"{quota['remaining']} kaldı", C["green"] if quota["remaining"] > 15 else C["yellow"])

        self.settings = {**self.settings, "veri_kapsami": scope}
        save_json(SETTINGS_FILE, self.settings)

    def _refresh_data_health(self):
        summary = data_quality_summary(self.df, self.active_df)
        lines = [
            f"Toplam mac          : {summary['total_rows']}",
            f"Aktif veri          : {summary['active_rows']}",
            f"Tarih araligi       : {summary['date_min'] or '-'} / {summary['date_max'] or '-'}",
            f"Lig sayisi          : {summary['league_count']}",
            f"Tekrar eden mac     : {summary['duplicate_rows']}",
            f"Skoru eksik mac     : {summary['missing_score_rows']}",
            f"MS orani eksik mac  : {summary['missing_core_odds_rows']}",
            "",
            "En cok mac olan ligler:",
        ]
        top_leagues = summary["top_leagues"] or []
        if top_leagues:
            lines.extend(f"- {league}: {count}" for league, count in top_leagues)
        else:
            lines.append("- Veri yok")

        self.health_text.configure(state="normal")
        self.health_text.delete("1.0", "end")
        self.health_text.insert("end", "\n".join(lines))
        self.health_text.configure(state="disabled")

    def _pick_file(self):
        path = filedialog.askopenfilename(title="Oran Excel Dosyası Seç",
                                           filetypes=[("Excel Dosyaları", "*.xlsx *.xls")])
        if path:
            self._load_data(path=path, silent=False)

    def _set_match_scope(self, value):
        self.settings = {**self.settings, "eslesme_kapsami": value}
        save_json(SETTINGS_FILE, self.settings)

    # ─────────────────────────────────────────────────────────
    #  BUGUNUN MACLARI
    # ─────────────────────────────────────────────────────────
    def _refresh_live_signals(self):
        self.live_refresh_btn.configure(state="disabled")
        self.live_status.configure(text="Canli veri taraniyor...", text_color=C["text2"])
        threading.Thread(target=self._live_signal_worker, daemon=True).start()

    def _live_signal_worker(self):
        try:
            fixtures = get_live_fixtures()
            signals = []
            # The signal rules only cover 55-80th-minute matches with fewer
            # than three goals.  Do not spend API requests on the rest.
            candidates = [fixture for fixture in fixtures if self._is_live_signal_candidate(fixture)]
            failed_count = 0
            for index, fixture in enumerate(candidates, start=1):
                fixture_id = fixture.get("fixture", {}).get("id")
                if not fixture_id:
                    continue
                self.root.after(0, lambda current=index, total=len(candidates): self.live_status.configure(
                    text=f"Canli tarama suruyor... uygun mac {current}/{total}", text_color=C["text2"]))
                try:
                    stats = get_fixture_statistics(fixture_id)
                    # Live odds are queried only after the statistics make a
                    # match a possible signal, saving one request per weak match.
                    if not build_goal_signals(fixture, stats, []):
                        continue
                    odds = get_live_odds(fixture_id)
                    save_live_snapshot(DATA_FILE, fixture_id, {"fixture": fixture, "statistics": stats, "odds": odds})
                    for signal in build_goal_signals(fixture, stats, odds):
                        save_live_signal(DATA_FILE, signal)
                        signals.append(signal)
                except Exception:
                    # A single fixture can be unavailable or rate-limited;
                    # retain the usable results from the rest of the scan.
                    failed_count += 1
            self.root.after(0, lambda: self._display_live_signals(
                signals, fixtures, candidates, quota_status(), failed_count, datetime.now()))
        except Exception as exc:
            self.root.after(0, lambda error=str(exc): self._display_live_signal_error(error))

    @staticmethod
    def _is_live_signal_candidate(fixture):
        elapsed = fixture.get("fixture", {}).get("status", {}).get("elapsed")
        goals = fixture.get("goals", {})
        total_goals = (goals.get("home") or 0) + (goals.get("away") or 0)
        return isinstance(elapsed, (int, float)) and 55 <= elapsed <= 80 and total_goals < 3

    def _display_live_signals(self, signals, fixtures, candidates, quota, failed_count, scanned_at):
        self.live_refresh_btn.configure(state="normal")
        scanned_count = len(fixtures)
        summary = live_signal_summary(DATA_FILE)
        actionable = sum(1 for signal in signals if signal.get("actionable"))
        values = (str(scanned_count), str(len(candidates)), str(len(signals)), str(actionable),
                  f"{quota['used']}/{quota['limit']}", f"%{summary['success_rate']:.0f}")
        colors = (C["text1"], C["yellow"], C["cyan"], C["green"], C["yellow"],
                  C["green"] if summary["success_rate"] >= 50 else C["text2"])
        for pill, value, color in zip(self.live_pills, values, colors): pill.set(value, color)
        for widget in self.live_cards.winfo_children(): widget.destroy()
        scan_note = f"Son tarama: {scanned_at.strftime('%H:%M:%S')}"
        if failed_count:
            scan_note += f" · {failed_count} aday maçın verisi alınamadı"
        if not signals:
            ctk.CTkLabel(self.live_cards, text="Şu an eşikleri geçen canlı sinyal yok.", font=ctk.CTkFont("Segoe UI", 15, "bold"), text_color=C["text1"]).pack(pady=(40, 6))
            ctk.CTkLabel(self.live_cards, text="Sistem 55–80. dakika aralığında baskı ve canlı oran verisini birlikte değerlendirir.", text_color=C["text3"]).pack()
            self.live_status.configure(text=f"{scanned_count} canlı maç, {len(candidates)} aday tarandı · {scan_note}",
                                       text_color=C["yellow"] if failed_count else C["text3"])
        else:
            for signal in signals:
                SignalCard(self.live_cards, signal).pack(fill="x", padx=8, pady=6)
            self.live_status.configure(text=f"{scanned_count} canlı maç, {len(candidates)} aday tarandı; {len(signals)} sinyal bulundu · {scan_note}",
                                       text_color=C["yellow"] if failed_count else C["green"])

    def _settle_live_signals(self):
        try:
            for signal in pending_live_signals(DATA_FILE):
                fixture = get_fixture(signal["fixture_id"])
                if not fixture or fixture.get("fixture", {}).get("status", {}).get("short") not in {"FT", "AET", "PEN"}:
                    continue
                goals = fixture.get("goals", {}); home, away = int(goals.get("home") or 0), int(goals.get("away") or 0)
                total = home + away
                if signal["market"] == "2.5 Üst": outcome = "tuttu" if total > 2 else "yatti"
                else: outcome = "tuttu" if total > signal["score_home"] + signal["score_away"] else "yatti"
                settle_live_signal(DATA_FILE, signal["id"], outcome)
            self.live_status.configure(text="Bekleyen sinyal sonuçları güncellendi.", text_color=C["green"])
        except Exception as exc:
            self._display_live_signal_error(str(exc))

    def _display_live_signal_error(self, error):
        self.live_refresh_btn.configure(state="normal")
        self.live_status.configure(text="Canlı sinyal verisi alınamadı.", text_color=C["red"])
        for widget in self.live_cards.winfo_children(): widget.destroy()
        ctk.CTkLabel(self.live_cards, text="Canlı veri hatası", font=ctk.CTkFont("Segoe UI", 15, "bold"), text_color=C["red"]).pack(pady=(40, 6))
        ctk.CTkLabel(self.live_cards, text=str(error), text_color=C["text2"], wraplength=600).pack()

    def _load_today_matches(self):
        if self.active_df.empty:
            messagebox.showwarning("Uyarı", "Önce geçmiş oran verisinin yüklenmesini bekle.")
            return
        self.today_refresh_btn.configure(state="disabled")
        self.today_status.configure(text="Maclar ve oranlar cekiliyor...", text_color=C["text2"])
        threading.Thread(target=self._today_worker, daemon=True).start()

    def _today_worker(self):
        try:
            today_data = fetch_today_unplayed_odds()
            if today_data.empty:
                replace_upcoming_matches([], DATA_FILE)
                replace_today_display([], DATA_FILE)
                self.root.after(0, lambda: self._display_today_matches(pd.DataFrame(), 0))
                return

            replace_upcoming_matches(today_data.to_dict("records"), DATA_FILE)

            replace_today_display(today_data.to_dict("records"), DATA_FILE)
            self.root.after(0, lambda: self._display_today_matches(today_data, len(today_data)))
        except Exception as exc:
            error_message = str(exc)
            self.root.after(0, lambda message=error_message: self._display_today_error(message))

    def _prediction_summaries(self, row, history_data, tolerance, minimum_matches, match_scope=None):
        def unavailable(reason):
            return [
                {"Tahmin": reason, "Grup": tier, "Oran": "", "Güven": "—", "Kalite": "Az veri 0",
                 "Eşleşme": 0, "Karar": "Pas geç", "_quality_score": 0}
                for tier in ("Banko", "Riskli", "Ya Tutarsa")
            ]

        entered = {}
        for field in ODDS_FIELDS:
            value = to_float(row.get(field["col"]))
            if value is not None and value > 1.0:
                entered[field["key"]] = value
        if not entered:
            return unavailable("Oran yok")

        scoped_history, scope_label = filter_history_by_match_scope(
            history_data, row.get("Lig"), match_scope or self.match_scope_var.get(), minimum_matches
        )
        result = best_combo_search(scoped_history, entered, tolerance, minimum_matches)
        matches = result["matches"] if result else pd.DataFrame()
        if matches.empty:
            return unavailable("Veri yok")

        stats = outcome_stats(matches, target_league=row.get("Lig"), weight_recent=False)
        summaries = []
        for signal in model_market_predictions(stats, row):
            confidence = signal["model_pct"]
            quality = prediction_quality_score(
                confidence,
                len(matches),
                key_count=len(result.get("keys", [])),
                tolerance=result.get("tol", tolerance),
                scoped=scope_label != "Genel",
            )
            summaries.append({
                "Tahmin": signal["prediction"],
                "Grup": signal["tier"],
                "Oran": format_odd(signal["price"]),
                "Güven": f"%{confidence:.0f}",
                "Kalite": f"{quality['label']} {quality['score']:.0f}",
                "Eşleşme": len(matches),
                "Kapsam": scope_label,
                "Model": f"%{signal['model_pct']:.0f}",
                "Piyasa": "—" if signal["market_pct"] is None else f"%{signal['market_pct']:.0f}",
                "Karar": signal["decision"],
                "Kombinasyon": " + ".join(FIELD_BY_KEY[key]["label"] for key in result.get("keys", [])),
                "Tolerans": result.get("tol", tolerance),
                "_quality_score": quality["score"],
            })
        return summaries

    def _display_today_matches(self, data, count):
        self.today_refresh_btn.configure(state="normal")
        self.today_df = data
        self._update_today_league_options()
        self._sort_today_matches(self.today_sort_var.get(), persist=False)
        if data.empty:
            self.today_status.configure(text="Bugun icin oranli, oynanmamis mac bulunamadi.", text_color=C["text3"])
        else:
            self.today_status.configure(text=f"{count} maç çekildi. Tahmin için bir maça çift tıkla.", text_color=C["green"])

    def _display_today_error(self, error):
        self.today_refresh_btn.configure(state="normal")
        self.today_status.configure(text="Mac verisi cekilemedi.", text_color=C["red"])
        messagebox.showerror("Mac Verisi Hatası", error)

    def _restore_today_display(self):
        """Ağ isteği yapmadan son çekimdeki tahmin listesini açılışta gösterir."""
        cached = load_today_display(DATA_FILE)
        if cached.empty:
            return
        self.today_df = cached
        self._update_today_league_options()
        self._sort_today_matches(self.today_sort_var.get(), persist=False)
        match_count = len(cached[[column for column in ("Tarih", "Saat", "Lig", "Maç") if column in cached.columns]].drop_duplicates())
        self.today_status.configure(text=f"Son çekimden {match_count} maç yüklendi. Tahmin için maça çift tıkla.", text_color=C["green"])

    def _sort_today_matches(self, sort_mode=None, persist=True):
        sort_mode = sort_mode or self.today_sort_var.get()
        if persist:
            self.settings = {**self.settings, "bugun_siralama": sort_mode}
            save_json(SETTINGS_FILE, self.settings)
        self._render_today_tree(sort_today_data(self._filtered_today_data(), sort_mode))

    def _filtered_today_data(self):
        if self.today_df.empty:
            return self.today_df.copy()
        league = self.today_league_var.get()
        if league == "Tüm Ligler" or "Lig" not in self.today_df.columns:
            return self.today_df.copy()
        return self.today_df.loc[self.today_df["Lig"].astype(str).str.strip() == league].copy()

    def _update_today_league_options(self):
        if not hasattr(self, "today_league_menu"):
            return
        leagues = sorted({str(value).strip() for value in self.today_df.get("Lig", pd.Series(dtype="object")) if str(value).strip()})
        values = ["Tüm Ligler", *leagues]
        self.today_league_menu.configure(values=values)
        if self.today_league_var.get() not in values:
            self.today_league_var.set("Tüm Ligler")

    def _set_today_league_filter(self, value):
        self.settings = {**self.settings, "bugun_lig_filtresi": value}
        save_json(SETTINGS_FILE, self.settings)
        self._sort_today_matches(self.today_sort_var.get(), persist=False)
        visible_count = len(self._filtered_today_data())
        if not self.today_df.empty:
            self.today_status.configure(
                text=f"{visible_count}/{len(self.today_df)} maç gösteriliyor ({value}).",
                text_color=C["green"],
            )
            self._generate_coupon(silent=True)

    def _apply_coupon_risk(self, risk, regenerate=True):
        profile = COUPON_RISK_PROFILES.get(risk, COUPON_RISK_PROFILES["Dengeli"])
        # Özel mod, kullanıcının girdiği değerleri korur.
        for entry, value in (
            (self.coupon_min_conf_entry, profile.get("min_confidence")),
            (self.coupon_max_selection_entry, profile.get("max_selections")),
            (getattr(self, "coupon_gap_entry", None), profile.get("min_gap")),
        ):
            if entry is None or value is None:
                continue
            entry.delete(0, "end")
            entry.insert(0, str(value))
        self.settings = {**self.settings, "kupon_risk": risk}
        save_json(SETTINGS_FILE, self.settings)
        if regenerate and not self.today_df.empty:
            self._generate_coupon(silent=True)

    @staticmethod
    def _coupon_number(entry, default, minimum=0.0, maximum=10000.0, integer=False):
        try:
            value = float(entry.get().strip().replace(",", "."))
        except (AttributeError, ValueError):
            value = default
        value = max(minimum, min(maximum, value))
        return int(value) if integer else value

    def _persist_coupon_settings(self, min_confidence, max_selections, min_gap):
        risk = self.coupon_risk_var.get()
        profile = COUPON_RISK_PROFILES.get(risk, {})
        profile_values = (profile.get("min_confidence"), profile.get("max_selections"), profile.get("min_gap"))
        actual_values = (min_confidence, max_selections, min_gap)
        if risk != "Ozel" and any(value is None or float(value) != float(actual) for value, actual in zip(profile_values, actual_values)):
            risk = "Ozel"
            self.coupon_risk_var.set(risk)
        self.settings = {
            **self.settings,
            "kupon_risk": risk,
            "kupon_ayarlari": {
                "min_guven": min_confidence,
                "maks_secim": max_selections,
                "zaman_araligi": min_gap,
                "kaynak": self.coupon_scope_var.get(),
                "lig_cesitliligi": bool(self.coupon_same_league_var.get()),
                "grup": self.coupon_tier_var.get(),
            },
        }
        save_json(SETTINGS_FILE, self.settings)

    def _generate_coupon(self, silent=False):
        source_today = self._filtered_today_data()
        if source_today.empty:
            self.coupon_df = pd.DataFrame()
            self._render_coupon_tree(self.coupon_df)
            self._update_coupon_simulation()
            self.coupon_status.configure(text="Seçilen lig filtresinde uygun maç yok.", text_color=C["text3"])
            if not silent:
                messagebox.showinfo("Kupon Önerisi", "Seçilen lig filtresinde uygun maç yok.")
            return
        min_confidence = self._coupon_number(self.coupon_min_conf_entry, 55.0, 0.0, 75.0)
        max_selections = self._coupon_number(self.coupon_max_selection_entry, 5, 1, 12, integer=True)
        min_gap = self._coupon_number(self.coupon_gap_entry, 60, 0, 360, integer=True)
        self._persist_coupon_settings(min_confidence, max_selections, min_gap)

        risk = self.coupon_risk_var.get()
        source_mode = self.coupon_scope_var.get()
        if source_mode == "Sadece İncele":
            eligible = source_today.loc[
                source_today.get("Karar", pd.Series("Pas geç", index=source_today.index)) == "İncele"
            ].copy()
        else:
            eligible = source_today.copy()
        tier = self.coupon_tier_var.get()
        if tier != "Tümü":
            eligible = eligible.loc[eligible.get("Grup", pd.Series("—", index=eligible.index)) == tier].copy()
        self.coupon_df = select_coupon_candidates(
            add_prediction_odds(eligible),
            min_confidence,
            max_selections,
            min_kickoff_gap_minutes=min_gap,
            diversify_leagues=bool(self.coupon_same_league_var.get()),
        )
        self._render_coupon_tree(self.coupon_df)

        if self.coupon_df.empty:
            self._update_coupon_simulation()
            self.coupon_status.configure(
                text=f"{source_mode}: %{min_confidence:.0f} destek esiginde uygun secim bulunamadi.", text_color=C["yellow"]
            )
        else:
            total_odd = total_coupon_odds(self.coupon_df)
            self._update_coupon_simulation(total_odd)
            total_text = f" Toplam oran: {total_odd:.2f}." if total_odd else ""
            self.coupon_status.configure(
                text=f"{risk} · {tier}: {len(self.coupon_df)} secimlik kupon hazir.{total_text}", text_color=C["green"]
            )
        if not silent:
            self._show_page("coupon")

    def _update_coupon_simulation(self, total_odd=None):
        if not hasattr(self, "coupon_sim_lbl"):
            return
        if total_odd is None:
            total_odd = total_coupon_odds(self.coupon_df)
        stake_text = self.coupon_stake_entry.get().strip() if hasattr(self, "coupon_stake_entry") else "100"
        simulation = coupon_simulation(total_odd, stake_text)
        if not simulation["payout"]:
            self.coupon_sim_lbl.configure(text="Simülasyon: kupon oluşunca hesaplanır.", text_color=C["text3"])
            return
        self.settings = {**self.settings, "kupon_tutar": stake_text}
        save_json(SETTINGS_FILE, self.settings)
        self.coupon_sim_lbl.configure(
            text=(
                f"Simülasyon: {simulation['stake']:.2f} TL x {simulation['total_odds']:.2f} = "
                f"{simulation['payout']:.2f} TL olası dönüş | Net: {simulation['profit']:.2f} TL"
            ),
            text_color=C["green"],
        )

    def _render_coupon_tree(self, data):
        for item in self.coupon_tree.get_children():
            self.coupon_tree.delete(item)
        for index, row in data.iterrows():
            self.coupon_tree.insert("", "end", iid=str(index), values=(
                row.get("Saat", ""), row.get("Lig", ""), row.get("Maç", ""),
                row.get("Tahmin", ""), row.get("Grup", ""), row.get("Oran", ""), row.get("Güven", ""),
                row.get("Kalite", ""), row.get("Eşleşme", ""),
            ))

    def _render_today_tree(self, data):
        for item in self.today_tree.get_children():
            self.today_tree.delete(item)
        for index, row in data.iterrows():
            odds = [to_float(row.get(column)) for column in (
                "Maç Sonucu (1)", "Maç Sonucu (X)", "Maç Sonucu (2)"
            )]
            values = [row.get("Saat", ""), row.get("Lig", ""), row.get("Maç", "")]
            values.extend("" if value is None else f"{value:.2f}" for value in odds)
            values.extend([
                row.get("Tahmin", ""), row.get("Oran", ""), row.get("Model", ""),
                row.get("Piyasa", ""), row.get("Grup", ""), row.get("Karar", ""), row.get("Eşleşme", ""),
            ])
            self.today_tree.insert("", "end", iid=str(index), values=values)

    def _send_today_to_match(self, event):
        selected = self.today_tree.selection()
        if not selected:
            return
        index = int(selected[0])
        if index not in self.today_df.index:
            return
        row = self.today_df.loc[index]
        match_name = str(row.get("Maç", ""))
        teams = match_name.split(" vs ", 1)
        home = teams[0].strip()
        away = teams[1].strip() if len(teams) > 1 else ""
        self.ev_entry.delete(0, "end"); self.ev_entry.insert(0, home)
        self.dep_entry.delete(0, "end"); self.dep_entry.insert(0, away)
        self.lig_entry.delete(0, "end"); self.lig_entry.insert(0, str(row.get("Lig", "")))
        for key, widget in self.odds_widgets.items():
            value = to_float(row.get(FIELD_BY_KEY[key]["col"]))
            widget.set_value(value)
        self._show_page("match")

    def _selected_today_row(self):
        selected = self.today_tree.selection()
        if not selected or int(selected[0]) not in self.today_df.index:
            messagebox.showinfo("Bilgi", "Önce Bugünün Maçları listesinden bir maç seç.")
            return None
        return self.today_df.loc[int(selected[0])]

    def _track_selected_today_prediction(self):
        row = self._selected_today_row()
        if row is None:
            return
        self._track_prediction_row(row, "Bugünün Maçları")

    def _track_selected_coupon_prediction(self):
        selected = self.coupon_tree.selection()
        if not selected or int(selected[0]) not in self.coupon_df.index:
            messagebox.showinfo("Takip", "Önce Kupon Önerisi listesinden bir seçim seç.")
            return
        row = self.coupon_df.loc[int(selected[0])]
        self._track_prediction_row(row, "Kupon Önerisi")

    def _track_prediction_row(self, row, source):
        match_id = str(row.get("_mackolik_id", "")).strip()
        market = str(row.get("Tahmin", "")).strip()
        price = to_float(row.get("Oran"))
        invalid_markets = {"", "Oran yok", "Veri yok", "Piyasa verisi yok", "Pas geç"}
        if not match_id or match_id.lower() == "nan" or market in invalid_markets or price is None or price <= 1:
            messagebox.showinfo("Takip", "Bu seçimde takip edilebilir kaynak kimliği, tahmin veya oran yok.")
            return
        payload = {
            "mackolik_id": match_id, "match_name": str(row.get("Maç", "")),
            "match_date": str(row.get("Tarih", "")), "kickoff": str(row.get("Saat", "")),
            "league": str(row.get("Lig", "")), "market": market, "price": price,
            "confidence": to_float(row.get("Güven")), "combo": row.get("Kombinasyon"),
            "tolerance": to_float(row.get("Tolerans")),
        }
        created = save_tracked_prediction(DATA_FILE, payload)
        self.predictions = load_json(PREDICTIONS_FILE, self.predictions if isinstance(self.predictions, list) else [])
        visible = any(
            str(record.get("mackolik_id", "")) == match_id and str(record.get("market", "")) == market
            for record in self.predictions
        )
        if not visible:
            self.predictions.append({
                "ts": datetime.now().strftime("%d-%m-%Y %H:%M"), "mac": payload["match_name"],
                "tahmin": market, "durum": "Bekliyor", "guven": payload["confidence"] or "-",
                "esleme": row.get("Eşleşme", "-"), "tolerans": payload["tolerance"] or "-",
                "kombo": payload.get("combo") or source, "mackolik_id": match_id, "market": market,
                "takipli": True,
            })
            save_json(PREDICTIONS_FILE, self.predictions)
        self._refresh_history_tree()
        if created:
            messagebox.showinfo("Takibe alındı", f"Tahmin Geçmişine eklendi:\n{payload['match_name']} · {market}")
        else:
            messagebox.showinfo("Zaten takipte", f"Bu maç ve pazar zaten takipte. Tahmin Geçmişi yenilendi.")

    def _settle_tracked_predictions(self):
        threading.Thread(target=self._settle_tracked_worker, daemon=True).start()

    def _settle_tracked_worker(self):
        updated = 0
        settled = []
        for record in pending_tracked_predictions(DATA_FILE):
            result = fetch_match_result(record["mackolik_id"], record.get("match_date"))
            source = "Mackolik"
            if result is None:
                # Yalnızca Mackolik cevap vermediğinde ücretsiz, kapsama sınırlı yedeği dene.
                home, away = split_match(record.get("match_name", ""))
                result = get_finished_score(record.get("match_date", ""), home, away) if home and away else None
                source = "football-data.org"
            if result is None:
                continue
            status = settle_market(record["market"], *result)
            if status:
                settle_tracked_prediction(DATA_FILE, record["id"], status, f"{result[0]}-{result[1]}", source)
                updated += 1
                settled.append((str(record["mackolik_id"]), str(record["market"]), status, f"{result[0]}-{result[1]}"))
        self.root.after(0, lambda: self._finish_tracked_settlement(updated, settled))

    def _finish_tracked_settlement(self, updated, settled):
        if settled:
            self.predictions = load_json(PREDICTIONS_FILE, self.predictions if isinstance(self.predictions, list) else [])
            status_map = {(match_id, market): (status, score) for match_id, market, status, score in settled}
            changed = False
            for record in self.predictions:
                key = (str(record.get("mackolik_id", "")), str(record.get("market", "")))
                if key not in status_map:
                    continue
                status, score = status_map[key]
                record["durum"] = "Tuttu" if status == "tuttu" else "Yattı"
                record["skor"] = score
                changed = True
            if changed:
                save_json(PREDICTIONS_FILE, self.predictions)
        self._refresh_history_tree()
        messagebox.showinfo("Sonuç güncelleme", f"{updated} takip edilen tahmin sonuçlandırıldı.")

    # ─────────────────────────────────────────────────────────
    #  EŞLEŞTİRME
    # ─────────────────────────────────────────────────────────
    def _clear_odds(self):
        for w in self.odds_widgets.values():
            w.clear()

    def _log(self, msg):
        def _write():
            self.log_text.configure(state="normal")
            self.log_text.insert("end", msg + "\n")
            self.log_text.see("end")
            self.log_text.configure(state="disabled")
        self.root.after(0, _write)

    def _run_match(self):
        if self.active_df.empty:
            messagebox.showwarning("Uyarı", "Önce oran verisini yükle (bahis_oranlari.xlsx).")
            return
        entered = {}
        for key, w in self.odds_widgets.items():
            val = w.get_float()
            if val is not None and val > 1.0:
                entered[key] = val
        if len(entered) < 1:
            messagebox.showwarning("Uyarı", "Lütfen en az bir oran gir.")
            return
        match_settings = self.settings.get("tahmin_ayarlari", {})
        tol_base = to_float(match_settings.get("tolerans")) or FIXED_TOLERANCE
        min_m = int(to_float(match_settings.get("min_eslesme")) or FIXED_MINIMUM_MATCHES)

        match_scope = self.match_scope_var.get()
        self.settings = {
            **self.settings,
            "tahmin_ayarlari": {"tolerans": tol_base, "min_eslesme": min_m},
            "eslesme_kapsami": match_scope,
        }
        save_json(SETTINGS_FILE, self.settings)

        self.log_text.configure(state="normal"); self.log_text.delete("1.0", "end"); self.log_text.configure(state="disabled")
        self.match_status.configure(text="Hesaplanıyor...")
        self.save_btn.configure(state="disabled")
        league = self.lig_entry.get().strip()
        history_data, scope_label = filter_history_by_match_scope(self.active_df, league, match_scope, min_m)
        threading.Thread(target=self._match_worker, args=(entered, tol_base, min_m, history_data, scope_label), daemon=True).start()

    def _match_worker(self, entered, tol_base, min_m, history_data, scope_label):
        try:
            self._log(
                f"Girilen oran sayısı: {len(entered)}  |  Tolerans: %{tol_base}  |  "
                f"Min eşleşme: {min_m}  |  Kapsam: {scope_label}"
            )
            result = best_combo_search(history_data, entered, tol_base, min_m, logger=self._log)
            if result:
                result["scope_label"] = scope_label
            self.root.after(0, lambda: self._display_results(entered, result))
        except Exception as exc:
            self._log(f"Eşleştirme hatası: {exc}")
            self.root.after(0, lambda: self._display_match_error(str(exc)))

    def _display_match_error(self, detail):
        self.match_status.configure(text="Eşleştirme sırasında hata oluştu.")
        self.save_btn.configure(state="disabled")
        messagebox.showerror("Eşleştirme hatası", f"Sonuçlar oluşturulamadı:\n{detail}")

    def _display_results(self, entered, result):
        self.match_status.configure(text="")
        matches = result["matches"] if result else pd.DataFrame()
        keys_used = result["keys"] if result else []
        tol_used = result["tol"] if result else 0

        if matches is None or len(matches) == 0:
            self.pill_count.set("0", C["red"])
            self.pill_pred.set("—")
            self.pill_conf.set("—")
            self.pill_score.set("—")
            self.prob_bar.update_probs(33.3, 33.4, 33.3)
            self.analysis_graph.update_analysis()
            self.chip_over.configure(text="2.5 Üst: —")
            self.chip_btts.configure(text="KG Var: —")
            self.chip_goals.configure(text="Ort. Gol: —")
            self.chip_ht.configure(text="En Sık İY: —")
            self.chip_team_goals.configure(text="Ev/Dep 1.5: —")
            self.combo_lbl.configure(text="Yeterli eşleşme bulunamadı. Tolerans veya oran sayısını değiştirmeyi dene.")
            self._render_match_tree(pd.DataFrame(), [])
            self._last_result = None
            return

        stats = outcome_stats(matches)
        hp, dp, ap = stats["home_pct"], stats["draw_pct"], stats["away_pct"]
        # Sadece girilmiş ve fiyatı bilinen pazarlardan öneri üret.
        entered_odds = {FIELD_BY_KEY[key]["col"]: value for key, value in entered.items()}
        recommendation = best_prediction(stats, entered_odds)
        pred, color = recommendation["label"], recommendation["color"]
        raw_conf = recommendation["confidence"]
        if stats["n"] < 5:
            conf, note = min(raw_conf, 45.0), "az veri (maks %45)"
        elif stats["n"] < 10:
            conf, note = min(raw_conf, 60.0), "az veri (maks %60)"
        else:
            conf, note = raw_conf, ""
        scope_label = result.get("scope_label", "Genel")
        quality = prediction_quality_score(
            conf,
            len(matches),
            key_count=len(keys_used),
            tolerance=tol_used,
            scoped=scope_label != "Genel",
        )

        self.pill_count.set(str(len(matches)), C["text1"])
        self.pill_pred.set(pred, color)
        self.pill_conf.set(f"%{conf:.0f} | {quality['label']}", color)
        self.pill_score.set(f"{stats['best_ft']} ({stats['best_ft_n']}x)")
        self.prob_bar.update_probs(hp, dp, ap)
        graph_note = " | ".join(part for part in (scope_label, note, f"Veri yeterliliği {quality['score']:.0f}") if part)
        self.analysis_graph.update_analysis(stats, conf, graph_note)
        self.chip_over.configure(text=f"2.5 Üst: %{stats['over25_pct']:.0f}")
        self.chip_btts.configure(text=f"KG Var: %{stats['btts_pct']:.0f}")
        self.chip_goals.configure(text=f"Ort. Gol: {stats['avg_goals']:.2f}")
        self.chip_ht.configure(text=f"En Sık İY: {stats['best_ht']}")
        self.chip_team_goals.configure(
            text=f"Ev/Dep 1.5: %{stats['home_over15_pct']:.0f}/%{stats['away_over15_pct']:.0f}"
        )

        combo_txt = " + ".join(FIELD_BY_KEY[k]["label"] for k in keys_used)
        note_txt = f"  ⚠ {note}" if note else ""
        self.combo_lbl.configure(
            text=f"Kapsam: {scope_label}   |   Veri yeterliliği: {quality['label']} {quality['score']:.0f}   |   "
                 f"Kullanılan kombinasyon: {combo_txt}   |   Uygulanan tolerans: %{tol_used:.1f}   |   "
                 f"{len(matches)} geçmiş maç bulundu.{note_txt}")

        self._render_match_tree(matches, keys_used)

        self._last_result = {
            "pred": pred, "conf": conf, "count": len(matches),
            "tol": tol_used, "combo": combo_txt, "score": stats["best_ft"],
            "quality": quality["score"], "scope": scope_label,
        }
        self.save_btn.configure(state="normal")
        self._log(f"\nSonuç: {pred}  (güven %{conf:.1f})  —  {len(matches)} maç, kombinasyon: {combo_txt}")

    def _render_match_tree(self, df, keys_used):
        for item in self.match_tree.get_children():
            self.match_tree.delete(item)
        base_cols = ["Tarih", "Saat", "Lig", "Ev Sahibi", "Deplasman", "İY Skor", "Skor"]
        odds_cols = [FIELD_BY_KEY[k]["label"] for k in keys_used]
        columns = base_cols + odds_cols
        self.match_tree["columns"] = columns
        for c in columns:
            self.match_tree.heading(c, text=c)
            self.match_tree.column(c, width=90 if c not in ("Lig", "Ev Sahibi", "Deplasman") else 140, anchor="w")

        if df.empty:
            return
        for _, row in df.head(300).iterrows():
            values = [
                row.get("Tarih", ""), row.get("Saat", ""), row.get("Lig", ""),
                row.get("_ev", ""), row.get("_dep", ""), row.get("İY Skor", ""), row.get("Skor", ""),
            ]
            for k in keys_used:
                v = row.get(FIELD_BY_KEY[k]["col"])
                values.append("" if v is None or pd.isna(v) else f"{v:.2f}")
            self.match_tree.insert("", "end", values=values)

    # ---------- Tahmin kaydetme ----------
    def _save_prediction(self):
        if not self._last_result:
            return
        ev = self.ev_entry.get().strip()
        dep = self.dep_entry.get().strip()
        mac = f"{ev} vs {dep}" if ev and dep else (ev or dep or "Manuel Tahmin")
        record = {
            "ts": datetime.now().strftime("%d-%m-%Y %H:%M"),
            "mac": mac,
            "tahmin": self._last_result["pred"],
            "durum": "Bekliyor",
            "guven": round(self._last_result["conf"], 1),
            "esleme": self._last_result["count"],
            "tolerans": round(self._last_result["tol"], 1),
            "kombo": self._last_result["combo"],
            "kalite": round(self._last_result.get("quality", 0), 1),
            "kapsam": self._last_result.get("scope", "Genel"),
        }
        self.predictions.append(record)
        save_json(PREDICTIONS_FILE, self.predictions)
        self._refresh_history_tree()
        messagebox.showinfo("Kaydedildi", f"Tahmin geçmişe eklendi:\n{mac} → {record['tahmin']} (%{record['guven']})")

    # ---------- Geçmiş sekmesi ----------
    def _refresh_history_tree(self):
        self.predictions = load_json(PREDICTIONS_FILE, self.predictions if isinstance(self.predictions, list) else [])
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        for i, rec in enumerate(reversed(self.predictions)):
            self.history_tree.insert("", "end", iid=str(len(self.predictions) - 1 - i), values=(
                rec.get("ts", "-"), rec.get("mac", "-"), rec.get("tahmin", "-"),
                rec.get("durum", "Bekliyor"), rec.get("guven", "-"), rec.get("esleme", "-"),
                rec.get("tolerans", "-"), rec.get("kombo", "-"),
            ))
        self._refresh_history_stats()

    def _refresh_history_stats(self):
        if not hasattr(self, "history_stats_lbl"):
            return
        stats = prediction_history_stats(self.predictions)
        text = (
            f"Sonuçlanan: {stats['graded_total']}  |  "
            f"Tuttu: {stats['won']}  |  Yattı: {stats['lost']}  |  "
            f"Bekliyor: {stats['pending']}  |  Başarı: %{stats['success_pct']:.0f}"
        )
        self.history_stats_lbl.configure(text=text)
        days = {"Son 7 gün": 7, "Son 30 gün": 30}.get(self.performance_range_var.get())
        tracked = tracked_prediction_summary(
            DATA_FILE,
            market=None if self.performance_market_var.get() == "Tüm pazarlar" else self.performance_market_var.get(),
            confidence_band=None if self.performance_confidence_var.get() == "Tüm güvenler" else self.performance_confidence_var.get(),
            days=days,
        )
        market_text = " · ".join(f"{market}: %{won * 100 / total:.0f}" for market, (won, total) in tracked["markets"].items() if total)
        self.history_performance_lbl.configure(
            text=(f"Otomatik takip: {tracked['tuttu']} tuttu / {tracked['yatti']} yattı / {tracked['bekliyor']} bekliyor"
                  f" · Başarı %{tracked['success_rate']:.0f} · Net {tracked['net']:+.2f} birim"
                  + (f" · {market_text}" if market_text else ""))
        )

    def _filtered_tracked_records(self):
        days = {"Son 7 gün": 7, "Son 30 gün": 30}.get(self.performance_range_var.get())
        records = load_tracked_predictions(DATA_FILE, days=days)
        selected_market = self.performance_market_var.get()
        selected_confidence = self.performance_confidence_var.get()
        if selected_market != "Tüm pazarlar":
            records = [record for record in records if record.get("market") == selected_market]
        if selected_confidence != "Tüm güvenler":
            def matches_confidence(record):
                confidence = to_float(record.get("confidence")) or 0
                return ((selected_confidence == "<60" and confidence < 60) or
                        (selected_confidence == "60-74" and 60 <= confidence < 75) or
                        (selected_confidence == "75+" and confidence >= 75))
            records = [record for record in records if matches_confidence(record)]
        return records

    def _show_performance_center(self):
        records = self._filtered_tracked_records()
        report = strategy_performance_report(records, min_sample=5)
        self._latest_strategy_report = report
        window = ctk.CTkToplevel(self.root)
        window.title("Performans Merkezi")
        window.geometry("900x620")
        window.configure(fg_color=C["bg0"])
        top = ctk.CTkFrame(window, fg_color="transparent")
        top.pack(fill="x", padx=18, pady=(16, 8))
        ctk.CTkLabel(top, text="Performans Merkezi", font=ctk.CTkFont("Segoe UI", 18, "bold"),
                     text_color=C["text1"]).pack(side="left")
        ctk.CTkButton(top, text="CSV Aktar", width=105, height=28, corner_radius=8,
                      fg_color=C["bg3"], hover_color=C["bg4"], text_color=C["text1"],
                      command=self._export_strategy_report).pack(side="right")
        output = ctk.CTkTextbox(window, fg_color=C["bg2"], border_width=1, border_color=C["border"],
                                text_color=C["text1"], font=ctk.CTkFont("Consolas", 12), wrap="word")
        output.pack(fill="both", expand=True, padx=18, pady=(0, 18))
        settled = [record for record in records if str(record.get("status", "")).casefold() in {"tuttu", "yatti"}]
        lines = [
            f"Filtrede {len(records)} takip kaydı, {len(settled)} sonuçlanmış tahmin var.",
            "Her sonuç 1 birim bahis varsayımıyla hesaplanır; net = getiriler - toplam bahis.",
            "Strateji sıralaması en az 5 sonuçta ROI, ardından muhafazakâr başarı alt sınırına göre yapılır.",
            "",
        ]
        if report["best"]:
            best = report["best"]
            lines.extend([
                "EN İYİ DOĞRULANMIŞ STRATEJİ",
                f"Kombinasyon: {best['combo']}",
                f"Tolerans: %{best['tolerance']:.1f} | Sonuç: {best['won']}/{best['total']} | Başarı: %{best['success_rate']:.1f}",
                f"ROI: %{best['roi']:+.1f} | Net: {best['net']:+.2f} birim | Muhafazakâr başarı alt sınırı: %{best['lower_bound']:.1f}",
                "",
            ])
        else:
            lines.extend(["Henüz en az 5 sonuçlu bir kombinasyon/tolerans grubu yok.",
                          "Takibe alınan yeni tahminlerde kombinasyon ve tolerans otomatik kaydedilecektir.", ""])
        lines.append("KOMBİNASYON VE TOLERANS KARŞILAŞTIRMASI")
        if report["groups"]:
            for item in report["groups"]:
                status = "doğrulandı" if item["eligible"] else f"örneklem az (<{report['min_sample']})"
                tolerance = "—" if item["tolerance"] is None else f"%{item['tolerance']:.1f}"
                lines.append(f"• {item['combo']} | {tolerance} | {item['won']}/{item['total']} (%{item['success_rate']:.1f}) | ROI %{item['roi']:+.1f} | {status}")
        else:
            lines.append("Sonuçlanmış otomatik takip kaydı bulunmuyor.")
        output.insert("1.0", "\n".join(lines))
        output.configure(state="disabled")

    def _export_strategy_report(self):
        report = getattr(self, "_latest_strategy_report", None)
        if not report or not report["groups"]:
            messagebox.showinfo("Bilgi", "Dışa aktarılacak sonuçlanmış strateji verisi yok.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile="strateji_performansi.csv",
                                            filetypes=[("CSV Dosyası", "*.csv")])
        if not path:
            return
        pd.DataFrame(report["groups"]).drop(columns=["stake", "return"], errors="ignore").to_csv(path, index=False, encoding="utf-8-sig")
        messagebox.showinfo("Başarılı", f"Kaydedildi:\n{path}")

    def _set_selected_prediction_status(self, status):
        sel = self.history_tree.selection()
        if not sel:
            messagebox.showinfo("Bilgi", "Durum güncellemek için bir tahmin seç.")
            return
        for item in sel:
            idx = int(item)
            if 0 <= idx < len(self.predictions):
                self.predictions[idx]["durum"] = status
        save_json(PREDICTIONS_FILE, self.predictions)
        self._refresh_history_tree()

    def _delete_selected_history(self):
        sel = self.history_tree.selection()
        if not sel:
            messagebox.showinfo("Bilgi", "Silinecek satırı seç.")
            return
        if not messagebox.askyesno("Onay", f"{len(sel)} tahmini silmek istediğine emin misin?"):
            return
        indices = sorted((int(i) for i in sel), reverse=True)
        for idx in indices:
            if 0 <= idx < len(self.predictions):
                del self.predictions[idx]
        save_json(PREDICTIONS_FILE, self.predictions)
        self._refresh_history_tree()

    def _export_history(self):
        if not self.predictions:
            messagebox.showinfo("Bilgi", "Aktarılacak tahmin yok.")
            return
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile="tahmin_gecmisi.xlsx",
                                            filetypes=[("Excel Dosyası", "*.xlsx")])
        if not path:
            return
        pd.DataFrame(self.predictions).to_excel(path, index=False)
        messagebox.showinfo("Başarılı", f"Kaydedildi:\n{path}")

    # ---------- Veri gözat sekmesi ----------
    def _search_browse(self):
        if self.active_df.empty:
            return
        q = self.browse_search_var.get().strip().lower()
        if not q:
            self.browse_filtered = self.active_df.copy()
        else:
            terms = [t for t in q.split() if t]
            mask = self.active_df["_search"].apply(lambda t: all(term in t for term in terms))
            self.browse_filtered = self.active_df.loc[mask]
        self._render_browse_tree(self.browse_filtered)

    def _render_browse_tree(self, df):
        for item in self.browse_tree.get_children():
            self.browse_tree.delete(item)
        self.browse_status.configure(text=f"{len(df)} maç")
        for idx, row in df.head(500).iterrows():
            values = []
            for _, src, _ in BROWSE_COLUMNS:
                v = row.get(src, "")
                if isinstance(v, float) and pd.isna(v):
                    v = ""
                elif isinstance(v, float):
                    v = f"{v:.2f}"
                values.append(v)
            self.browse_tree.insert("", "end", iid=str(idx), values=values)

    def _send_browse_to_match(self, event):
        sel = self.browse_tree.selection()
        if not sel:
            return
        idx = int(sel[0])
        if idx not in self.active_df.index:
            return
        row = self.active_df.loc[idx]
        self.ev_entry.delete(0, "end"); self.ev_entry.insert(0, str(row.get("_ev", "")))
        self.dep_entry.delete(0, "end"); self.dep_entry.insert(0, str(row.get("_dep", "")))
        self.lig_entry.delete(0, "end"); self.lig_entry.insert(0, str(row.get("Lig", "")))
        for key, w in self.odds_widgets.items():
            col = FIELD_BY_KEY[key]["col"]
            val = row.get(col)
            w.set_value(None if val is None or pd.isna(val) else float(val))
        self._show_page("match")
        messagebox.showinfo("Aktarıldı", "Seçilen maçın oranları Eşleştirme sekmesine aktarıldı.\n"
                                          "İstersen düzenleyip 'ORANLARI EŞLEŞTİR' butonuna basabilirsin.")


if __name__ == "__main__":
    root = ctk.CTk()
    App(root)
    root.mainloop()
