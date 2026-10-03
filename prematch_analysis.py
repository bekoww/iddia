"""Geçmiş Mackolik arşivinden açıklanabilir form ve H2H özeti üretir."""

from __future__ import annotations

import re

import pandas as pd


def _team_key(value) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().casefold())


def split_match(match_name: str) -> tuple[str, str]:
    parts = re.split(r"\s+(?:vs|v\.|-|–)\s+", str(match_name or ""), maxsplit=1, flags=re.I)
    return (parts[0].strip(), parts[1].strip()) if len(parts) == 2 else ("", "")


def _score(value) -> tuple[int, int] | None:
    found = re.search(r"(\d+)\s*[-:]\s*(\d+)", str(value or ""))
    return (int(found.group(1)), int(found.group(2))) if found else None


def _team_summary(rows: pd.DataFrame, team: str, limit=5) -> dict:
    key = _team_key(team)
    games = []
    for _, row in rows.iterrows():
        home, away = str(row.get("_ev", "")), str(row.get("_dep", ""))
        score = _score(row.get("Skor"))
        if not score or key not in {_team_key(home), _team_key(away)}:
            continue
        is_home = key == _team_key(home)
        own, opponent = (score[0], score[1]) if is_home else (score[1], score[0])
        games.append({"opponent": away if is_home else home, "score": f"{own}-{opponent}", "result": "G" if own > opponent else "B" if own == opponent else "M", "goals": own + opponent, "btts": own > 0 and opponent > 0})
        if len(games) >= limit:
            break
    total = len(games)
    return {"team": team, "games": games, "played": total, "wins": sum(g["result"] == "G" for g in games), "draws": sum(g["result"] == "B" for g in games), "losses": sum(g["result"] == "M" for g in games), "avg_goals": round(sum(g["goals"] for g in games) / total, 2) if total else 0.0, "over25": sum(g["goals"] > 2 for g in games), "btts": sum(g["btts"] for g in games)}


def build_prematch_analysis(history: pd.DataFrame, match_name: str, limit=5) -> dict:
    home, away = split_match(match_name)
    rows = history.copy()
    if "_tarih_dt" in rows:
        rows = rows.sort_values("_tarih_dt", ascending=False)
    h_key, a_key = _team_key(home), _team_key(away)
    h2h = []
    for _, row in rows.iterrows():
        teams = {_team_key(row.get("_ev")), _team_key(row.get("_dep"))}
        score = _score(row.get("Skor"))
        if score and {h_key, a_key}.issubset(teams):
            h2h.append({"date": str(row.get("Tarih", "")), "match": str(row.get("Maç", "")), "score": f"{score[0]}-{score[1]}"})
        if len(h2h) >= limit:
            break
    return {"match": match_name, "home": _team_summary(rows, home, limit), "away": _team_summary(rows, away, limit), "h2h": h2h}


def settle_market(market: str, home: int, away: int) -> str | None:
    total = home + away
    checks = {"MS 1": home > away, "MS X": home == away, "MS 2": away > home, "2.5 Üst": total > 2, "2.5 Alt": total < 3, "KG Var": home > 0 and away > 0, "KG Yok": home == 0 or away == 0}
    if market not in checks:
        return None
    return "tuttu" if checks[market] else "yatti"
