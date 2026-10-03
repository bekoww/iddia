"""Açıklanabilir canlı maç sinyalleri.

Bu modül sonuç garantisi vermez; yalnızca maç baskısı ve canlı fiyat uygunluğu
aynı anda oluştuğunda incelemeye değer sinyal üretir.
"""

from __future__ import annotations


def _number(value) -> float:
    try:
        return float(str(value).replace("%", "").replace(",", "."))
    except (TypeError, ValueError):
        return 0.0


def _stats_totals(statistics: list[dict]) -> dict[str, float]:
    totals = {"Total Shots": 0.0, "Shots on Goal": 0.0, "Corner Kicks": 0.0}
    for team in statistics:
        for item in team.get("statistics", []):
            stat_type = item.get("type")
            if stat_type in totals:
                totals[stat_type] += _number(item.get("value"))
    return totals


def _live_price(live_odds: list[dict], market: str) -> float | None:
    keywords = ("next goal", "sonraki gol") if market == "Sonraki Gol" else ("over/under", "over under", "toplam gol")
    for group in live_odds:
        for bookmaker in group.get("bookmakers", []):
            for bet in bookmaker.get("bets", []):
                if not any(word in str(bet.get("name", "")).lower() for word in keywords):
                    continue
                for value in bet.get("values", []):
                    label = str(value.get("value", "")).lower()
                    if market == "Sonraki Gol" and label in {"home", "away", "yes"}:
                        return _number(value.get("odd")) or None
                    if market == "2.5 Üst" and ("over 2.5" in label or "üst 2.5" in label):
                        return _number(value.get("odd")) or None
    return None


def build_goal_signals(fixture: dict, statistics: list[dict], live_odds: list[dict]) -> list[dict]:
    """55–80. dakika arası gol baskısı için muhafazakâr bir sinyal üretir."""
    fixture_info = fixture.get("fixture", {})
    elapsed = fixture_info.get("status", {}).get("elapsed") or 0
    elapsed = int(elapsed)
    goals = fixture.get("goals", {})
    score = int(goals.get("home") or 0) + int(goals.get("away") or 0)
    if not 55 <= elapsed <= 80 or score > 2:
        return []

    totals = _stats_totals(statistics)
    score_points = 0
    reasons = []
    if totals["Total Shots"] >= 14:
        score_points += 25; reasons.append(f"{int(totals['Total Shots'])} toplam şut")
    if totals["Shots on Goal"] >= 5:
        score_points += 35; reasons.append(f"{int(totals['Shots on Goal'])} isabetli şut")
    if totals["Corner Kicks"] >= 5:
        score_points += 20; reasons.append(f"{int(totals['Corner Kicks'])} korner")
    if score <= 1:
        score_points += 20; reasons.append("düşük skorlu maç")
    if score_points < 70:
        return []

    confidence = min(88, score_points)
    teams = fixture.get("teams", {})
    match = f"{teams.get('home', {}).get('name', '?')} - {teams.get('away', {}).get('name', '?')}"
    result = []
    for market in ("Sonraki Gol", "2.5 Üst"):
        price = _live_price(live_odds, market)
        result.append({"fixture_id": fixture_info.get("id"), "match": match, "minute": elapsed,
                       "market": market, "confidence": confidence, "actionable": price is not None,
                       "price": price, "reasons": reasons, "score_home": int(goals.get("home") or 0),
                       "score_away": int(goals.get("away") or 0)})
    return result
