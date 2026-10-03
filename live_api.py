"""API-Football canlı veri istemcisi.

Anahtar yalnızca ortam değişkeninden veya proje kökündeki `.env` dosyasından
okunur; arayüz ve loglara yazılmaz.
"""

from __future__ import annotations

import os
import time
from pathlib import Path
from threading import Lock

import requests
from data_store import api_usage, record_api_request


BASE_URL = "https://v3.football.api-sports.io"
PROJECT_DIR = Path(__file__).resolve().parent
DATABASE_PATH = PROJECT_DIR / "data" / "iddaa.sqlite"
CACHE_SECONDS = 120
SAFE_DAILY_LIMIT = 90
# API-Football plans can additionally impose a per-minute quota.  Leave
# headroom so one scan does not immediately hit it.
REQUESTS_PER_MINUTE = 8
MIN_REQUEST_INTERVAL = 60 / REQUESTS_PER_MINUTE
_CACHE: dict[str, tuple[float, list | dict]] = {}
_REQUEST_LOCK = Lock()
_LAST_REQUEST_AT = 0.0


class LiveAPIError(RuntimeError):
    pass


class LiveRateLimitError(LiveAPIError):
    """The provider's short-term request quota has been reached."""

    pass


def _read_local_key() -> str | None:
    env_path = PROJECT_DIR / ".env"
    if not env_path.exists():
        return None
    for line in env_path.read_text(encoding="utf-8").splitlines():
        if line.startswith("API_FOOTBALL_KEY="):
            return line.partition("=")[2].strip()
    return None


def get_api_key() -> str:
    key = os.getenv("API_FOOTBALL_KEY") or _read_local_key()
    if not key:
        raise LiveAPIError("API_FOOTBALL_KEY tanımlı değil. `.env` dosyasını kontrol edin.")
    return key


def _get(endpoint: str, **params) -> list | dict:
    cache_key = f"{endpoint}:{sorted(params.items())}"
    cached = _CACHE.get(cache_key)
    if cached and time.monotonic() - cached[0] < CACHE_SECONDS:
        return cached[1]
    global _LAST_REQUEST_AT
    # Serialize calls from the scan and settlement workers so their combined
    # traffic also stays under the provider's minute limit.
    with _REQUEST_LOCK:
        wait_seconds = MIN_REQUEST_INTERVAL - (time.monotonic() - _LAST_REQUEST_AT)
        if wait_seconds > 0:
            time.sleep(wait_seconds)
        response = requests.get(
            f"{BASE_URL}/{endpoint.lstrip('/')}",
            headers={"x-apisports-key": get_api_key()},
            params={key: value for key, value in params.items() if value is not None},
            timeout=20,
        )
        _LAST_REQUEST_AT = time.monotonic()
    record_api_request(DATABASE_PATH, SAFE_DAILY_LIMIT)
    try:
        payload = response.json()
    except ValueError as exc:
        raise LiveAPIError(f"Canlı veri servisi geçerli yanıt döndürmedi (HTTP {response.status_code}).") from exc
    errors = payload.get("errors") if isinstance(payload, dict) else None
    error_text = str(errors or "")
    if response.status_code == 429 or "ratelimit" in error_text.lower() or "too many requests" in error_text.lower():
        raise LiveRateLimitError(
            "Canlı veri sağlayıcısının dakika kotası doldu. Lütfen yaklaşık bir dakika sonra yeniden deneyin."
        )
    if response.status_code >= 400 or errors:
        raise LiveAPIError(f"Canlı veri servisi hatası: {errors or f'HTTP {response.status_code}'}")
    result = payload.get("response", payload)
    _CACHE[cache_key] = (time.monotonic(), result)
    return result


def quota_status() -> dict:
    used = api_usage(DATABASE_PATH)
    return {"used": used, "limit": SAFE_DAILY_LIMIT, "remaining": max(0, SAFE_DAILY_LIMIT - used)}


def get_fixture(fixture_id: int | str) -> dict | None:
    response = _get("fixtures", id=fixture_id)
    return response[0] if isinstance(response, list) and response else None


def verify_connection() -> dict:
    """Anahtarı doğrular; hassas bilgileri döndürmez."""
    response = _get("status")
    if not isinstance(response, dict):
        return {"connected": True}
    subscription = response.get("subscription") or {}
    requests_info = response.get("requests") or {}
    return {
        "connected": True,
        "plan": subscription.get("plan") or subscription.get("name"),
        "remaining_requests": requests_info.get("remaining"),
    }


def get_live_fixtures() -> list[dict]:
    """Devam eden tüm maçların skor ve durum bilgisini döndürür."""
    response = _get("fixtures", live="all")
    return response if isinstance(response, list) else []


def get_fixture_statistics(fixture_id: int | str) -> list[dict]:
    response = _get("fixtures/statistics", fixture=fixture_id)
    return response if isinstance(response, list) else []


def get_live_odds(fixture_id: int | str) -> list[dict]:
    response = _get("odds/live", fixture=fixture_id)
    return response if isinstance(response, list) else []
