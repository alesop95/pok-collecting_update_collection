"""
Storico prezzi su SQLite.

prices.db e' la fonte autoritativa: ogni run dello script aggiunge uno
snapshot con timestamp. prices_cache.xlsx (rigenerato a ogni run) e' solo
una vista derivata, usa-e-getta, pensata per essere consultata da
MAIN.xlsx via XLOOKUP.
"""

from __future__ import annotations
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

from paths import PRICES_DB

DB_PATH = PRICES_DB

SCHEMA = """
CREATE TABLE IF NOT EXISTS prices (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT NOT NULL,
    sheet_name TEXT NOT NULL,
    expansion_id INTEGER,
    blueprint_id INTEGER,
    card_name TEXT NOT NULL,
    collector_number TEXT,
    condition_input TEXT,
    condition_ct TEXT,
    language TEXT,
    reverse_holo INTEGER,
    first_edition INTEGER,
    price_eur REAL,
    source TEXT DEFAULT 'cardtrader'
);
CREATE INDEX IF NOT EXISTS idx_prices_blueprint ON prices(blueprint_id);
CREATE INDEX IF NOT EXISTS idx_prices_ts ON prices(ts);
CREATE INDEX IF NOT EXISTS idx_prices_lookup
    ON prices(sheet_name, collector_number, condition_input, language, reverse_holo, first_edition);
"""


def get_connection(path: str | Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    conn.executescript(SCHEMA)
    return conn


def insert_snapshot(conn: sqlite3.Connection, **row) -> None:
    """Inserisce uno snapshot. Il timestamp e' ISO-8601 UTC, generato qui se non passato."""
    if "ts" not in row or not row["ts"]:
        row["ts"] = datetime.now(timezone.utc).isoformat(timespec="seconds")
    cols = list(row.keys())
    placeholders = ", ".join("?" * len(cols))
    conn.execute(
        f"INSERT INTO prices ({', '.join(cols)}) VALUES ({placeholders})",
        list(row.values()),
    )
    conn.commit()


def latest_per_card(conn: sqlite3.Connection) -> list[sqlite3.Row]:
    """
    Ultimo snapshot per ogni combinazione univoca di carta+stato.
    La chiave logica e' (sheet, collector, condition_input, language, reverse_holo, first_edition).
    """
    q = """
    WITH ranked AS (
        SELECT
            *,
            ROW_NUMBER() OVER (
                PARTITION BY sheet_name, COALESCE(collector_number,''),
                             COALESCE(condition_input,''), COALESCE(language,''),
                             COALESCE(reverse_holo,0), COALESCE(first_edition,0)
                ORDER BY ts DESC, id DESC
            ) AS rn
        FROM prices
    )
    SELECT * FROM ranked WHERE rn = 1
    ORDER BY sheet_name, collector_number
    """
    return list(conn.execute(q))


def price_history(conn: sqlite3.Connection, blueprint_id: int) -> list[sqlite3.Row]:
    """Tutta la storia prezzi per un blueprint. Utile per grafici di andamento."""
    q = "SELECT ts, condition_ct, language, price_eur FROM prices WHERE blueprint_id = ? ORDER BY ts"
    return list(conn.execute(q, (blueprint_id,)))
