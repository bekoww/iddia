"""SQLite tabanlı maç veri deposu.

Excel dosyaları içe/dışa aktarma için korunur; uygulamanın ana çalışma verisi
bu depoda tutulur. Market sütunları değişken olabildiğinden her maçın tüm
verisi JSON olarak saklanır, ana alanlar ise hızlı sorgu ve tekilleştirme için
ayrı kolonlardadır.
"""

from __future__ import annotations

import json
import math
import sqlite3
from datetime import date, datetime
from pathlib import Path

import pandas as pd


MATCH_KEY_COLUMNS = ["Tarih", "Saat", "Lig", "Maç"]


def _key(row: dict) -> str:
    values = (str(row.get(column, "")).strip() for column in MATCH_KEY_COLUMNS)
    return "\x1f".join(values)


def _json_value(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(value, (datetime, date, pd.Timestamp)):
        return value.isoformat()
    return value


def _connect(database_path: str | Path) -> sqlite3.Connection:
    database_path = Path(database_path)
    database_path.parent.mkdir(parents=True, exist_ok=True)
    connection = sqlite3.connect(database_path)
    connection.execute("PRAGMA journal_mode=WAL")
    connection.execute("PRAGMA synchronous=NORMAL")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS matches (
            match_key TEXT PRIMARY KEY,
            tarih TEXT NOT NULL,
            saat TEXT NOT NULL,
            lig TEXT NOT NULL,
            mac TEXT NOT NULL,
            payload TEXT NOT NULL,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    connection.execute("CREATE INDEX IF NOT EXISTS idx_matches_tarih ON matches(tarih)")
    connection.execute("CREATE INDEX IF NOT EXISTS idx_matches_lig ON matches(lig)")
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS upcoming_matches (
            match_key TEXT PRIMARY KEY,
            tarih TEXT NOT NULL,
            saat TEXT NOT NULL,
            lig TEXT NOT NULL,
            mac TEXT NOT NULL,
            payload TEXT NOT NULL,
            fetched_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    connection.execute(
        """
        CREATE TABLE IF NOT EXISTS today_display (
            display_order INTEGER PRIMARY KEY,
            payload TEXT NOT NULL,
            fetched_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )
        """
    )
    connection.execute("CREATE TABLE IF NOT EXISTS api_usage (usage_day TEXT PRIMARY KEY, request_count INTEGER NOT NULL DEFAULT 0)")
    connection.execute("CREATE TABLE IF NOT EXISTS live_snapshots (fixture_id INTEGER NOT NULL, captured_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, payload TEXT NOT NULL)")
    connection.execute("""CREATE TABLE IF NOT EXISTS prematch_snapshots (
        mackolik_id TEXT PRIMARY KEY, captured_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, payload TEXT NOT NULL
    )""")
    connection.execute("""CREATE TABLE IF NOT EXISTS tracked_predictions (
        id INTEGER PRIMARY KEY AUTOINCREMENT, mackolik_id TEXT NOT NULL, match_name TEXT NOT NULL,
        match_date TEXT, kickoff TEXT, league TEXT, market TEXT NOT NULL, price REAL, confidence REAL,
        status TEXT NOT NULL DEFAULT 'bekliyor', score TEXT, created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
        settled_at TEXT, result_source TEXT, UNIQUE(mackolik_id, market)
    )""")
    tracked_columns = {row[1] for row in connection.execute("PRAGMA table_info(tracked_predictions)")}
    if "result_source" not in tracked_columns:
        connection.execute("ALTER TABLE tracked_predictions ADD COLUMN result_source TEXT")
    if "combo" not in tracked_columns:
        connection.execute("ALTER TABLE tracked_predictions ADD COLUMN combo TEXT")
    if "tolerance" not in tracked_columns:
        connection.execute("ALTER TABLE tracked_predictions ADD COLUMN tolerance REAL")
    connection.execute("""CREATE TABLE IF NOT EXISTS live_signals (
        id INTEGER PRIMARY KEY AUTOINCREMENT, fixture_id INTEGER NOT NULL, market TEXT NOT NULL,
        match_name TEXT NOT NULL, minute INTEGER NOT NULL, score_home INTEGER NOT NULL, score_away INTEGER NOT NULL,
        price REAL, confidence INTEGER NOT NULL, reasons TEXT NOT NULL, status TEXT NOT NULL DEFAULT 'bekliyor',
        created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP, settled_at TEXT, UNIQUE(fixture_id, market, minute)
    )""")
    return connection


def api_usage(database_path: str | Path) -> int:
    today = date.today().isoformat(); connection = _connect(database_path)
    try:
        row = connection.execute("SELECT request_count FROM api_usage WHERE usage_day=?", (today,)).fetchone()
        return int(row[0]) if row else 0
    finally: connection.close()


def record_api_request(database_path: str | Path, limit=90) -> int:
    today = date.today().isoformat(); connection = _connect(database_path)
    try:
        used = connection.execute("SELECT request_count FROM api_usage WHERE usage_day=?", (today,)).fetchone()
        count = int(used[0]) if used else 0
        if count >= limit: raise RuntimeError(f"Günlük canlı API güvenli limiti doldu ({limit}).")
        connection.execute("INSERT INTO api_usage(usage_day, request_count) VALUES (?, 1) ON CONFLICT(usage_day) DO UPDATE SET request_count=request_count+1", (today,))
        connection.commit(); return count + 1
    finally: connection.close()


def save_live_snapshot(database_path: str | Path, fixture_id: int, payload: dict) -> None:
    connection = _connect(database_path)
    try:
        connection.execute("INSERT INTO live_snapshots(fixture_id, payload) VALUES (?, ?)", (fixture_id, json.dumps(payload, ensure_ascii=False, default=str)))
        connection.commit()
    finally: connection.close()


def save_live_signal(database_path: str | Path, signal: dict) -> None:
    connection = _connect(database_path)
    try:
        connection.execute("""INSERT OR IGNORE INTO live_signals
        (fixture_id, market, match_name, minute, score_home, score_away, price, confidence, reasons)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""", (signal["fixture_id"], signal["market"], signal["match"], signal["minute"], signal["score_home"], signal["score_away"], signal.get("price"), signal["confidence"], json.dumps(signal["reasons"], ensure_ascii=False)))
        connection.commit()
    finally: connection.close()


def pending_live_signals(database_path: str | Path) -> list[dict]:
    connection = _connect(database_path)
    try:
        rows = connection.execute("SELECT id, fixture_id, market, score_home, score_away FROM live_signals WHERE status='bekliyor'").fetchall()
        return [{"id": row[0], "fixture_id": row[1], "market": row[2], "score_home": row[3], "score_away": row[4]} for row in rows]
    finally: connection.close()


def settle_live_signal(database_path: str | Path, signal_id: int, status: str) -> None:
    connection = _connect(database_path)
    try:
        connection.execute("UPDATE live_signals SET status=?, settled_at=CURRENT_TIMESTAMP WHERE id=?", (status, signal_id)); connection.commit()
    finally: connection.close()


def live_signal_summary(database_path: str | Path) -> dict[str, int | float]:
    connection = _connect(database_path)
    try:
        counts = {row[0]: row[1] for row in connection.execute("SELECT status, COUNT(*) FROM live_signals GROUP BY status")}
        won, lost = counts.get("tuttu", 0), counts.get("yatti", 0)
        settled = won + lost
        return {"bekliyor": counts.get("bekliyor", 0), "tuttu": won, "yatti": lost,
                "success_rate": round(won * 100 / settled, 1) if settled else 0.0}
    finally: connection.close()


def load_prematch_snapshot(database_path: str | Path, mackolik_id: str) -> dict | None:
    connection = _connect(database_path)
    try:
        row = connection.execute("SELECT payload FROM prematch_snapshots WHERE mackolik_id=?", (str(mackolik_id),)).fetchone()
        return json.loads(row[0]) if row else None
    finally: connection.close()


def save_prematch_snapshot(database_path: str | Path, mackolik_id: str, payload: dict) -> None:
    connection = _connect(database_path)
    try:
        connection.execute("INSERT INTO prematch_snapshots(mackolik_id, payload) VALUES (?, ?) ON CONFLICT(mackolik_id) DO UPDATE SET payload=excluded.payload, captured_at=CURRENT_TIMESTAMP", (str(mackolik_id), json.dumps(payload, ensure_ascii=False)))
        connection.commit()
    finally: connection.close()


def save_tracked_prediction(database_path: str | Path, record: dict) -> bool:
    """Takip kaydını ekler; aynı maç/pazar zaten varsa ``False`` döndürür."""
    connection = _connect(database_path)
    try:
        cursor = connection.execute("""INSERT OR IGNORE INTO tracked_predictions
        (mackolik_id, match_name, match_date, kickoff, league, market, price, confidence, combo, tolerance)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", (str(record["mackolik_id"]), record["match_name"], record.get("match_date"), record.get("kickoff"), record.get("league"), record["market"], record.get("price"), record.get("confidence"), record.get("combo"), record.get("tolerance")))
        connection.commit()
        return cursor.rowcount == 1
    finally: connection.close()


def pending_tracked_predictions(database_path: str | Path) -> list[dict]:
    connection = _connect(database_path)
    try:
        rows = connection.execute("SELECT id, mackolik_id, market, match_date, match_name FROM tracked_predictions WHERE status='bekliyor'").fetchall()
        return [{"id": r[0], "mackolik_id": r[1], "market": r[2], "match_date": r[3], "match_name": r[4]} for r in rows]
    finally: connection.close()


def settle_tracked_prediction(database_path: str | Path, prediction_id: int, status: str, score: str, source: str = "Mackolik") -> None:
    connection = _connect(database_path)
    try:
        connection.execute("UPDATE tracked_predictions SET status=?, score=?, result_source=?, settled_at=CURRENT_TIMESTAMP WHERE id=?", (status, score, source, prediction_id)); connection.commit()
    finally: connection.close()


def tracked_prediction_summary(database_path: str | Path, market: str | None = None, confidence_band: str | None = None, days: int | None = None) -> dict:
    connection = _connect(database_path)
    try:
        rows = connection.execute("SELECT market, confidence, price, status, created_at FROM tracked_predictions").fetchall()
    finally: connection.close()
    if days:
        cutoff = datetime.now().timestamp() - days * 86400
        rows = [row for row in rows if datetime.fromisoformat(row[4]).timestamp() >= cutoff]
    if market:
        rows = [row for row in rows if row[0] == market]
    if confidence_band:
        def in_band(value):
            return (confidence_band == "<60" and (value or 0) < 60) or (confidence_band == "60-74" and 60 <= (value or 0) < 75) or (confidence_band == "75+" and (value or 0) >= 75)
        rows = [row for row in rows if in_band(row[1])]
    result = {"total": len(rows), "bekliyor": 0, "tuttu": 0, "yatti": 0, "success_rate": 0.0, "stake": 0.0, "return": 0.0, "net": 0.0, "markets": {}, "confidence": {"<60": [0, 0], "60-74": [0, 0], "75+": [0, 0]}}
    for market, confidence, price, status, _created_at in rows:
        result[status] = result.get(status, 0) + 1
        entry = result["markets"].setdefault(market, [0, 0])
        if status in {"tuttu", "yatti"}:
            entry[1] += 1
            band = "75+" if (confidence or 0) >= 75 else "60-74" if (confidence or 0) >= 60 else "<60"
            result["confidence"][band][1] += 1
            if status == "tuttu":
                entry[0] += 1; result["confidence"][band][0] += 1
            result["stake"] += 1
            if status == "tuttu" and price:
                result["return"] += float(price)
    settled = result["tuttu"] + result["yatti"]
    result["success_rate"] = round(result["tuttu"] * 100 / settled, 1) if settled else 0.0
    result["net"] = round(result["return"] - result["stake"], 2)
    return result


def load_tracked_predictions(database_path: str | Path, days: int | None = None) -> list[dict]:
    """Return tracked predictions with strategy metadata for performance analysis."""
    connection = _connect(database_path)
    try:
        rows = connection.execute("""SELECT league, market, confidence, price, status, created_at, combo, tolerance
                                  FROM tracked_predictions""").fetchall()
    finally:
        connection.close()
    cutoff = datetime.now().timestamp() - days * 86400 if days else None
    records = []
    for league, market, confidence, price, status, created_at, combo, tolerance in rows:
        if cutoff and datetime.fromisoformat(created_at).timestamp() < cutoff:
            continue
        records.append({"league": league, "market": market, "confidence": confidence, "price": price,
                        "status": status, "created_at": created_at, "combo": combo, "tolerance": tolerance})
    return records


def load_matches(database_path: str | Path) -> pd.DataFrame:
    """Depodaki tüm maçları DataFrame olarak döndürür."""
    path = Path(database_path)
    if not path.exists():
        return pd.DataFrame()
    connection = _connect(path)
    try:
        rows = connection.execute("SELECT payload FROM matches").fetchall()
    finally:
        connection.close()
    return pd.DataFrame([json.loads(row[0]) for row in rows])


def clear_matches(database_path: str | Path) -> None:
    """Depodaki tüm maçları siler; yalnızca açıkça sıfırdan çekim istendiğinde kullanılır."""
    connection = _connect(database_path)
    try:
        connection.execute("DELETE FROM matches")
        connection.commit()
    finally:
        connection.close()


def save_matches(rows: list[dict], database_path: str | Path) -> int:
    """Maçları anahtarına göre ekler/günceller ve depodaki toplamı döndürür."""
    if not rows:
        return 0
    latest_rows = {}
    for row in rows:
        if all(column in row for column in MATCH_KEY_COLUMNS):
            if all(str(row.get(column, "")).strip() == column for column in MATCH_KEY_COLUMNS):
                continue
            latest_rows[_key(row)] = row
    if not latest_rows:
        return 0

    payload_rows = []
    for match_key, row in latest_rows.items():
        clean_row = {key: _json_value(value) for key, value in row.items()}
        payload_rows.append((
            match_key,
            str(row.get("Tarih", "")).strip(), str(row.get("Saat", "")).strip(),
            str(row.get("Lig", "")).strip(), str(row.get("Maç", "")).strip(),
            json.dumps(clean_row, ensure_ascii=False, separators=(",", ":"), default=str),
        ))

    connection = _connect(database_path)
    try:
        connection.executemany(
            """
            INSERT INTO matches (match_key, tarih, saat, lig, mac, payload)
            VALUES (?, ?, ?, ?, ?, ?)
            ON CONFLICT(match_key) DO UPDATE SET
                tarih=excluded.tarih, saat=excluded.saat, lig=excluded.lig,
                mac=excluded.mac, payload=excluded.payload,
                updated_at=CURRENT_TIMESTAMP
            """,
            payload_rows,
        )
        total = connection.execute("SELECT COUNT(*) FROM matches").fetchone()[0]
        connection.commit()
        return total
    finally:
        connection.close()


def import_excel_files(paths: list[str | Path], database_path: str | Path) -> int:
    """Uygun Excel dosyalarını depoya aktarır; son dosyadaki kayıt önceliklidir."""
    frames = []
    for path in paths:
        frame = pd.read_excel(path)
        if set(MATCH_KEY_COLUMNS).issubset(frame.columns):
            frames.append(frame)
    if not frames:
        return 0
    return save_matches(pd.concat(frames, ignore_index=True).to_dict("records"), database_path)


def export_excel(database_path: str | Path, output_path: str | Path) -> int:
    """Depodaki maçları paylaşım/inceleme amacıyla Excel'e aktarır."""
    frame = load_matches(database_path)
    if frame.empty:
        return 0
    sort_columns = [column for column in MATCH_KEY_COLUMNS if column in frame.columns]
    if sort_columns:
        frame = frame.sort_values(sort_columns).reset_index(drop=True)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    frame.to_excel(output_path, index=False)
    return len(frame)


def load_upcoming_matches(database_path: str | Path) -> pd.DataFrame:
    """Son çekilen, henüz başlamamış maçları döndürür."""
    path = Path(database_path)
    if not path.exists():
        return pd.DataFrame()
    connection = _connect(path)
    try:
        rows = connection.execute("SELECT payload FROM upcoming_matches").fetchall()
    finally:
        connection.close()
    return pd.DataFrame([json.loads(row[0]) for row in rows])


def replace_upcoming_matches(rows: list[dict], database_path: str | Path) -> int:
    """Önceki fikstür anlık görüntüsünü yeni çekimle değiştirir."""
    clean_rows = []
    for row in rows:
        if all(column in row for column in MATCH_KEY_COLUMNS):
            clean_rows.append((
                _key(row),
                str(row.get("Tarih", "")).strip(), str(row.get("Saat", "")).strip(),
                str(row.get("Lig", "")).strip(), str(row.get("Maç", "")).strip(),
                json.dumps({key: _json_value(value) for key, value in row.items()}, ensure_ascii=False, separators=(",", ":"), default=str),
            ))
    connection = _connect(database_path)
    try:
        connection.execute("DELETE FROM upcoming_matches")
        connection.executemany(
            "INSERT INTO upcoming_matches (match_key, tarih, saat, lig, mac, payload) VALUES (?, ?, ?, ?, ?, ?)",
            clean_rows,
        )
        connection.commit()
        return len(clean_rows)
    finally:
        connection.close()


def load_today_display(database_path: str | Path) -> pd.DataFrame:
    """Son çekimde hesaplanan Bugünün Maçları görünümünü döndürür."""
    path = Path(database_path)
    if not path.exists():
        return pd.DataFrame()
    connection = _connect(path)
    try:
        rows = connection.execute("SELECT payload FROM today_display ORDER BY display_order").fetchall()
    finally:
        connection.close()
    return pd.DataFrame([json.loads(row[0]) for row in rows])


def replace_today_display(rows: list[dict], database_path: str | Path) -> int:
    """Bugünün Maçları için yeni analiz görünümünü atomik olarak saklar."""
    payloads = [
        (index, json.dumps({key: _json_value(value) for key, value in row.items()}, ensure_ascii=False,
                           separators=(",", ":"), default=str))
        for index, row in enumerate(rows)
    ]
    connection = _connect(database_path)
    try:
        connection.execute("DELETE FROM today_display")
        connection.executemany("INSERT INTO today_display (display_order, payload) VALUES (?, ?)", payloads)
        connection.commit()
        return len(payloads)
    finally:
        connection.close()
