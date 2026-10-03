"""football-data.org sonuç/fikstür yedeği.

Anahtar yalnızca `.env` içindeki FOOTBALL_DATA_KEY değişkeninden okunur.
Bu istemci oran veya canlı veri için kullanılmaz.
"""

from __future__ import annotations

import os
import re
import unicodedata
from functools import lru_cache
from datetime import datetime
from pathlib import Path

import requests


BASE_URL = "https://api.football-data.org/v4"
PROJECT_DIR = Path(__file__).resolve().parent


def _key() -> str | None:
    value = os.getenv("FOOTBALL_DATA_KEY")
    if value:
        return value
    path = PROJECT_DIR / ".env"
    if not path.exists():
        return None
    for line in path.read_text(encoding="utf-8").splitlines():
        if line.startswith("FOOTBALL_DATA_KEY="):
            return line.partition("=")[2].strip() or None
    return None


def _name(value: str) -> str:
    plain = unicodedata.normalize("NFKD", str(value).casefold()).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", plain)


@lru_cache(maxsize=32)
def _finished_matches(match_date: str) -> list[dict]:
    key = _key()
    if not key:
        return []
    try:
        date_iso = datetime.strptime(match_date, "%d-%m-%Y").strftime("%Y-%m-%d")
        response = requests.get(
            f"{BASE_URL}/matches", params={"dateFrom": date_iso, "dateTo": date_iso},
            headers={"X-Auth-Token": key}, timeout=15,
        )
        payload = response.json()
        if response.status_code >= 400:
            return []
    except (ValueError, requests.RequestException):
        return []
    return payload.get("matches", [])


def get_finished_score(match_date: str, home_team: str, away_team: str) -> tuple[int, int] | None:
    """Belirtilen günde kesinleşmiş, adı eşleşen maçın tam zaman skorunu döndürür."""
    home_key, away_key = _name(home_team), _name(away_team)
    for match in _finished_matches(match_date):
        if match.get("status") != "FINISHED":
            continue
        if _name(match.get("homeTeam", {}).get("name", "")) != home_key or _name(match.get("awayTeam", {}).get("name", "")) != away_key:
            continue
        score = match.get("score", {}).get("fullTime", {})
        if score.get("home") is not None and score.get("away") is not None:
            return int(score["home"]), int(score["away"])
    return None
