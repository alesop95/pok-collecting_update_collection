"""
Client CardTrader: chiamate API + caching su disco delle export blueprint.

Le export per espansione sono enormi ma quasi statiche, quindi le mettiamo in
cache su disco (.cache/blueprints_{expansion_id}.json) per evitare di
scaricarle a ogni esecuzione. Per i prezzi (marketplace/products) NIENTE cache:
sono dati di mercato live e devono essere sempre freschi.
"""

from __future__ import annotations
import json
import os
import time
import requests

from paths import CACHE_DIR

API_BASE = "https://api.cardtrader.com/api/v2"
CACHE_DIR.mkdir(exist_ok=True)


def _auth_header() -> dict:
    """Token letto da variabile d'ambiente (popolata via .env tramite paths.py)."""
    token = os.environ.get("CT_AUTH_TOKEN")
    if not token:
        raise RuntimeError(
            "CT_AUTH_TOKEN non impostato. Crea il file .env nella root del progetto:\n"
            "  echo CT_AUTH_TOKEN=eyJ... > .env\n"
            "Verifica con: python -c \"from dotenv import dotenv_values; print('OK' if dotenv_values().get('CT_AUTH_TOKEN') else 'MANCANTE')\""
        )
    return {"Authorization": f"Bearer {token}"}


def _get(path: str, params: dict | None = None, retries: int = 3) -> dict | list:
    """GET con backoff esponenziale su errori transitori."""
    url = f"{API_BASE}{path}"
    last_exc = None
    for attempt in range(retries):
        try:
            r = requests.get(url, params=params, headers=_auth_header(), timeout=30)
            if r.status_code == 429:  # rate limit
                wait = 2 ** attempt
                print(f"  [rate limit] attendo {wait}s...")
                time.sleep(wait)
                continue
            r.raise_for_status()
            return r.json()
        except requests.RequestException as e:
            last_exc = e
            time.sleep(2 ** attempt)
    raise RuntimeError(f"GET {url} fallito dopo {retries} tentativi") from last_exc


def get_blueprints(expansion_id: int, force_refresh: bool = False) -> list[dict]:
    """Scarica e cacha la lista blueprint per un'espansione."""
    cache_file = CACHE_DIR / f"blueprints_{expansion_id}.json"
    if cache_file.exists() and not force_refresh:
        return json.loads(cache_file.read_text(encoding="utf-8"))

    print(f"  [API] /blueprints/export?expansion_id={expansion_id}")
    data = _get("/blueprints/export", params={"expansion_id": expansion_id})
    cache_file.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def find_blueprint(
    blueprints: list[dict],
    name: str,
    version: str | None = None,
) -> dict | None:
    """
    Cerca un blueprint per nome (case-insensitive). Se più carte hanno lo
    stesso nome nello stesso set (es. reprint con collector number diverso),
    usa `version` per disambiguare (es. "26/132").
    """
    name_norm = name.strip().lower()
    candidates = [b for b in blueprints if b.get("name", "").strip().lower() == name_norm]
    if not candidates:
        return None
    if version and len(candidates) > 1:
        for c in candidates:
            if c.get("version") == version:
                return c
    return candidates[0]


def get_marketplace_products(blueprint_id: int) -> list[dict]:
    """
    Ritorna la lista dei 25 prodotti più economici per il blueprint.
    L'API risponde con {str(blueprint_id): [products]}, qui appiattiamo.
    """
    print(f"  [API] /marketplace/products?blueprint_id={blueprint_id}")
    data = _get("/marketplace/products", params={"blueprint_id": blueprint_id})
    return data.get(str(blueprint_id), [])
