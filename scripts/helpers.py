"""
Helpers puri usati sia dalla pipeline sia dai test.
Niente dipendenze da xlwings/openpyxl: solo logica pura.
"""

from __future__ import annotations


# Mappa nome esteso lingua -> codice CardTrader (pokemon_language)
LANGUAGE_MAP = {
    "english": "en",  "en": "en",  "eng": "en",  "inglese": "en",
    "italian": "it",  "it": "it",  "ita": "it",  "italiano": "it",
    "french":  "fr",  "fr": "fr",  "francese": "fr",
    "german":  "de",  "de": "de",  "tedesco": "de",
    "spanish": "es",  "es": "es",  "spagnolo": "es",
    "portuguese": "pt", "pt": "pt", "portoghese": "pt",
    "dutch":   "nl",  "nl": "nl",  "olandese": "nl",
    "russian": "ru",  "ru": "ru",  "russo": "ru",
    "polish":  "pl",  "pl": "pl",  "polacco": "pl",
    "swedish": "sv",  "sv": "sv",  "svedese": "sv",
    "korean":  "kr",  "kr": "kr",  "ko": "kr",  "coreano": "kr",
    "japanese": "jp", "jp": "jp",  "ja": "jp",  "giapponese": "jp",
    "chinese (simplified)":  "zh-CN", "zh-cn": "zh-CN", "chinese simplified":  "zh-CN",
    "chinese (traditional)": "zh-TW", "zh-tw": "zh-TW", "chinese traditional": "zh-TW",
    "indonesian": "id", "thai": "th",
}

# Default quando la colonna "Language" e' vuota in MAIN.xlsx
DEFAULT_LANGUAGE = "en"


def parse_bool(v) -> bool:
    if isinstance(v, bool):
        return v
    if v is None:
        return False
    s = str(v).strip().lower()
    return s in ("true", "vero", "si", "si'", "1", "yes", "y", "x", "v")


def normalize_collector(c) -> str | None:
    """
    '026/132' -> '26/132'. Tollera int, float, stringhe con spazi.
    Em-dash ('—'), '-', '?' e vuoto -> None (carta senza collector number,
    tipico delle Vending Machine vintage e di alcune promo).
    """
    if c is None:
        return None
    s = str(c).strip()
    if not s or s in ("—", "-", "–", "?", "N/A", "n/a"):
        return None
    if "/" in s:
        a, b = s.split("/", 1)
        try:
            return f"{int(a.strip())}/{int(b.strip())}"
        except ValueError:
            return s
    return s


def normalize_language(lang) -> str:
    """Normalizza la lingua al codice CardTrader. Default DEFAULT_LANGUAGE se vuoto/None."""
    if lang is None:
        return DEFAULT_LANGUAGE
    s = str(lang).strip().lower()
    if not s:
        return DEFAULT_LANGUAGE
    return LANGUAGE_MAP.get(s, s)  # se non in mappa, ritorna il valore originale (potrebbe gia' essere un codice)


def find_blueprint_robust(blueprints: list[dict], name: str, collector: str | None) -> dict | None:
    """Match per nome (case-insensitive) + collector_number normalizzato."""
    target_collector = normalize_collector(collector)
    name_norm = name.strip().lower()
    candidates = [b for b in blueprints if b.get("name", "").strip().lower() == name_norm]
    if not candidates:
        return None
    if target_collector:
        for c in candidates:
            if normalize_collector(c.get("version")) == target_collector:
                return c
            fp = c.get("fixed_properties") or {}
            if normalize_collector(fp.get("collector_number")) == target_collector:
                return c
    return candidates[0]


def build_lookup_key(sheet_name, collector, condition_input, language, reverse_holo, first_edition) -> str:
    """
    Chiave concatenata che DEVE combaciare con quella generata dalla formula
    XLOOKUP su MAIN.xlsx. Se modifichi qui, modifica anche generate_formula_snippets.
    """
    return (
        f"{sheet_name}|{collector or ''}|{(condition_input or '').upper()}|"
        f"{(language or '').lower()}|{'true' if reverse_holo else 'false'}|"
        f"{'true' if first_edition else 'false'}"
    )


def is_owned(cell_value) -> bool:
    """
    True se la colonna 'yes/no' contiene 'Y' (o varianti).
    Convenzione utente: 'Y' = posseduta, vuoto/null = non posseduta.
    """
    if cell_value is None:
        return False
    return str(cell_value).strip().upper() in ("Y", "YES", "SI", "SI'", "TRUE", "1", "X")
