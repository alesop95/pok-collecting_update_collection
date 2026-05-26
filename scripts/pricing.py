"""
Trasformazione: mapping condizioni Cardmarket -> CardTrader
e aggregazione del prezzo a partire dai prodotti marketplace.

Mapping condizioni (dal documento di analisi):
    M  -> NM     (Mint non esiste su CardTrader)
    NM -> NM
    EX -> SP
    LP -> MP
    GD -> PL
    PL -> PL
    PO -> PO
"""

from __future__ import annotations
from statistics import median

# Sigle Cardmarket (input utente sul foglio Excel) -> valore "condition" di CardTrader
CONDITION_MAP_SHORT_TO_CT = {
    "M":  "Near Mint",
    "NM": "Near Mint",
    "EX": "Slightly Played",
    "LP": "Moderately Played",
    "GD": "Played",
    "PL": "Played",
    "PO": "Poor",
}

# Forme estese che l'utente potrebbe scrivere a mano nel foglio
CONDITION_MAP_LONG_TO_CT = {
    "mint":              "Near Mint",
    "near mint":         "Near Mint",
    "excellent":         "Slightly Played",
    "light played":      "Moderately Played",
    "lightly played":    "Moderately Played",
    "good":              "Played",
    "played":            "Played",
    "poor":              "Poor",
}


def normalize_condition(raw) -> str | None:
    """Mappa un valore di condizione utente (sigla o esteso) al valore CardTrader."""
    if raw is None:
        return None
    key = str(raw).strip()
    if not key:
        return None
    if key.upper() in CONDITION_MAP_SHORT_TO_CT:
        return CONDITION_MAP_SHORT_TO_CT[key.upper()]
    return CONDITION_MAP_LONG_TO_CT.get(key.lower())


def aggregate_price(
    products: list[dict],
    condition_ct: str,
    language: str = "it",
    reverse_holo: bool | None = None,
    first_edition: bool | None = None,
    strategy: str = "min",
) -> float | None:
    """
    Filtra i prodotti per condizione + lingua + reverse_holo + first_edition,
    ritorna un prezzo aggregato in EURO.

    Per Pokemon, la property CardTrader del "reverse holo" e' 'pokemon_reverse'
    (NON 'foil'); quella della prima edizione e' 'first_edition'.

    strategy:
      - "min":    prezzo piu' basso disponibile (utile per "quanto costa ricomprarla")
      - "median": mediana (piu' robusta per stima valore collezione)
      - "avg":    media aritmetica
    """
    matched_cents = []
    for p in products:
        props = p.get("properties_hash", {}) or {}
        if props.get("condition") != condition_ct:
            continue
        if language and props.get("pokemon_language") not in (None, language):
            continue
        if reverse_holo is not None and bool(props.get("pokemon_reverse", False)) != reverse_holo:
            continue
        if first_edition is not None and bool(props.get("first_edition", False)) != first_edition:
            continue
        cents = (p.get("price") or {}).get("cents")
        if cents is None:
            continue
        matched_cents.append(cents)

    if not matched_cents:
        return None

    if strategy == "min":
        value = min(matched_cents)
    elif strategy == "median":
        value = median(matched_cents)
    elif strategy == "avg":
        value = sum(matched_cents) / len(matched_cents)
    else:
        raise ValueError(f"Strategia sconosciuta: {strategy}")

    return round(value / 100, 2)
