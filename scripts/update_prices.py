"""
PokeCollecting: pipeline read-only MAIN.xlsx -> SQLite + prices_cache.xlsx

NON tocca MAIN.xlsx. Apre il workbook in lettura, calcola i prezzi via
CardTrader, e produce:
  - prices.db            : SQLite storico (autoritativo, append-only) -- in cwd
  - prices_cache.xlsx    : vista corrente -- ACCANTO a MAIN.xlsx (necessario per XLOOKUP)
  - lookup_formulas.txt  : formule XLOOKUP pronte da incollare in MAIN.xlsx -- in cwd

Uso:
  set CT_AUTH_TOKEN=eyJ...
  python update_prices.py "C:\\percorso\\MAIN.xlsx"

Opzioni:
  --sheet NAME       Processa solo il foglio NAME (utile per test)
  --all              Processa anche carte NON possedute (default: solo Y in col I)

Configurazione (in expansions.json):
  Mappa nome_foglio -> expansion_id di CardTrader. Esempio:
    {"GYM_HEROES": 1480, "BASE_SET": 1469}
"""

from __future__ import annotations
import argparse
import json
import time
from datetime import datetime, timezone
from pathlib import Path

import xlwings as xw
from openpyxl import Workbook

from cardtrader_client import get_blueprints, get_marketplace_products
from pricing import normalize_condition, aggregate_price
from database import get_connection, insert_snapshot, latest_per_card
from helpers import (
    parse_bool, normalize_collector, find_blueprint_robust,
    build_lookup_key, normalize_language, is_owned,
)
from paths import EXPANSIONS_JSON, LOOKUP_FORMULAS, DEFAULT_WORKBOOK


# === Mappa colonne MAIN.xlsx (header riga 5, dati da riga 6) ===
# C: Rarity       I: Owned (Y/vuoto)   O: Estimated Value  <- formula XLOOKUP
# D: No.          J: Language          P: Notes
# E: Card name    K: 1st edition
# F: Type         L: Stamped
# G: Promotion    M: Reverse Holo
# H: Quantity     N: Condition
COLUMNS = {
    "card_name":     "E",
    "collector":     "D",
    "owned":         "I",
    "language":      "J",
    "first_edition": "K",
    "reverse_holo":  "M",
    "condition":     "N",
}

START_ROW = 6
SLEEP_BETWEEN_CARDS = 0.4


# === Helpers ===

def last_data_row(sheet, col: str, start_row: int) -> int:
    try:
        last = sheet.range(f"{col}{sheet.cells.last_cell.row}").end("up").row
    except Exception:
        last = start_row
    return max(last, start_row - 1)


def load_expansions_config(path: Path = EXPANSIONS_JSON) -> dict[str, int | None]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"{p} non trovato. Generalo con: python scripts/discover_expansions.py"
        )
    return {k: int(v) if v is not None else None for k, v in json.loads(p.read_text(encoding="utf-8")).items()}


# === Pipeline ===

def process_sheet(sheet, expansion_id: int, conn, ts: str, only_owned: bool) -> dict:
    """Ritorna {'processed': N, 'priced': M, 'skipped_unowned': K, 'not_matched': J}."""
    sheet_name = sheet.name
    print(f"\n=== {sheet_name}  (expansion_id={expansion_id}) ===")

    print("  carico blueprints...")
    blueprints = get_blueprints(expansion_id)
    print(f"  {len(blueprints)} blueprint disponibili")

    end_row = last_data_row(sheet, COLUMNS["card_name"], START_ROW)
    if end_row < START_ROW:
        print("  nessuna riga dati, salto")
        return {"processed": 0, "priced": 0, "skipped_unowned": 0, "not_matched": 0}

    stats = {"processed": 0, "priced": 0, "skipped_unowned": 0, "not_matched": 0}

    for row in range(START_ROW, end_row + 1):
        name = sheet.range(f"{COLUMNS['card_name']}{row}").value
        if not name or not str(name).strip():
            continue
        name = str(name).strip()

        owned = is_owned(sheet.range(f"{COLUMNS['owned']}{row}").value)
        if only_owned and not owned:
            stats["skipped_unowned"] += 1
            continue

        collector       = normalize_collector(sheet.range(f"{COLUMNS['collector']}{row}").value)
        condition_input = sheet.range(f"{COLUMNS['condition']}{row}").value
        condition_input = str(condition_input).strip() if condition_input else None
        language        = normalize_language(sheet.range(f"{COLUMNS['language']}{row}").value)
        reverse_holo    = parse_bool(sheet.range(f"{COLUMNS['reverse_holo']}{row}").value)
        first_edition   = parse_bool(sheet.range(f"{COLUMNS['first_edition']}{row}").value)

        condition_ct = normalize_condition(condition_input) if condition_input else "Near Mint"
        if condition_input and condition_ct is None:
            print(f"  riga {row}: condizione '{condition_input}' non riconosciuta, salto")
            continue

        bp = find_blueprint_robust(blueprints, name, collector)
        if bp is None:
            print(f"  riga {row}: '{name}' [{collector}] -> blueprint non trovato")
            stats["not_matched"] += 1
            stats["processed"] += 1
            continue

        products = get_marketplace_products(bp["id"])
        price = aggregate_price(
            products, condition_ct, language=language,
            reverse_holo=reverse_holo, first_edition=first_edition, strategy="min",
        )
        flags = f"{' rev' if reverse_holo else ''}{' 1st' if first_edition else ''}"
        print(f"  riga {row}: '{name}' [{collector}] {condition_ct}/{language}{flags} "
              f"-> {f'EUR {price:.2f}' if price is not None else 'n/d'}")

        insert_snapshot(
            conn, ts=ts, sheet_name=sheet_name, expansion_id=expansion_id,
            blueprint_id=bp["id"], card_name=name, collector_number=collector,
            condition_input=condition_input, condition_ct=condition_ct,
            language=language, reverse_holo=int(reverse_holo),
            first_edition=int(first_edition), price_eur=price,
        )
        stats["processed"] += 1
        if price is not None:
            stats["priced"] += 1
        time.sleep(SLEEP_BETWEEN_CARDS)

    return stats


def regenerate_cache_xlsx(conn, output_path: Path) -> None:
    rows = latest_per_card(conn)
    wb = Workbook()
    ws = wb.active
    ws.title = "Latest"
    ws.append([
        "key", "price_eur", "last_updated", "card_name", "sheet_name",
        "collector_number", "condition_input", "language", "reverse_holo", "first_edition",
    ])
    for r in rows:
        key = build_lookup_key(
            r["sheet_name"], r["collector_number"], r["condition_input"],
            r["language"], r["reverse_holo"], r["first_edition"],
        )
        ws.append([
            key, r["price_eur"], r["ts"], r["card_name"], r["sheet_name"],
            r["collector_number"], r["condition_input"], r["language"],
            "true" if r["reverse_holo"] else "false",
            "true" if r["first_edition"] else "false",
        ])
    for cell in ws[1]:
        cell.font = cell.font.copy(bold=True)
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 30
    wb.save(output_path)
    print(f"\nGenerato {output_path}  ({len(rows)} righe)")


def generate_formula_snippets(expansions: dict, output_path: Path = LOOKUP_FORMULAS) -> None:
    """
    Genera le formule CERCA.X (Excel italiano) per ogni foglio, una per foglio.

    Versione italiana: usa SE.ERRORE/CERCA.X/MAIUSC/MINUSC/SE e separatore ';'.
    Se in futuro serve la versione inglese (IFERROR/XLOOKUP/UPPER/LOWER/IF con ','),
    duplicare la funzione o parametrizzarla con un flag --locale.
    """
    c = COLUMNS
    lines = [
        "# Formule CERCA.X da incollare nella colonna 'Estimated Value' (O) di MAIN.xlsx, riga 6.",
        "# Versione per Excel ITALIANO (nomi funzione localizzati, separatore ';').",
        "# IMPORTANTE: prices_cache.xlsx deve essere nella stessa cartella di MAIN.xlsx.",
        "#",
        "# Per ogni foglio: copia la formula in O6 e trascina giu' fino all'ULTIMA riga",
        "# del catalogo (anche carte non possedute: mostreranno 'n/d' finche' non scrivi 'Y'",
        "# in colonna I e rilanci lo script).",
        "",
    ]
    for sheet_name, exp_id in expansions.items():
        if exp_id is None:
            continue
        formula = (
            f'=SE.ERRORE(CERCA.X("{sheet_name}"&"|"&{c["collector"]}6&"|"'
            f'&MAIUSC({c["condition"]}6)&"|"&MINUSC({c["language"]}6)&"|"'
            f'&SE({c["reverse_holo"]}6;"true";"false")&"|"'
            f'&SE({c["first_edition"]}6;"true";"false");'
            f"'[prices_cache.xlsx]Latest'!$A:$A;'[prices_cache.xlsx]Latest'!$B:$B);\"n/d\")"
        )
        lines.append(f"## {sheet_name}")
        lines.append(formula)
        lines.append("")
    Path(output_path).write_text("\n".join(lines), encoding="utf-8")
    print(f"Generato {output_path}")


def main():
    parser = argparse.ArgumentParser(description="PokeCollecting prices updater")
    parser.add_argument("workbook", nargs="?", default=str(DEFAULT_WORKBOOK),
                        help=f"Percorso a MAIN.xlsx (default: {DEFAULT_WORKBOOK})")
    parser.add_argument("--sheet", help="Processa solo questo foglio (test mode)")
    parser.add_argument("--all", action="store_true",
                        help="Processa anche carte non possedute (default: solo Y in col I)")
    args = parser.parse_args()

    workbook_path = Path(args.workbook).resolve()
    if not workbook_path.exists():
        print(f"ERRORE: {workbook_path} non trovato")
        return

    only_owned = not args.all
    cache_path = workbook_path.parent / "prices_cache.xlsx"

    expansions = load_expansions_config()
    if args.sheet:
        if args.sheet not in expansions:
            print(f"ERRORE: '{args.sheet}' non in expansions.json")
            return
        expansions = {args.sheet: expansions[args.sheet]}

    ts = datetime.now(timezone.utc).isoformat(timespec="seconds")
    conn = get_connection()

    print(f"Modalita': {'solo possedute (col I = Y)' if only_owned else 'TUTTE le carte'}")
    print(f"Fogli da processare: {len(expansions)}")
    print(f"Cache verra' scritto in: {cache_path}")

    app = xw.App(visible=False, add_book=False)
    totals = {"processed": 0, "priced": 0, "skipped_unowned": 0, "not_matched": 0}
    try:
        wb = app.books.open(str(workbook_path), read_only=True, update_links=False)
        try:
            sheets_in_wb = {s.name for s in wb.sheets}
            for sheet_name, exp_id in expansions.items():
                if exp_id is None:
                    print(f"\n[SKIP] {sheet_name}: expansion_id mancante in expansions.json")
                    continue
                if sheet_name not in sheets_in_wb:
                    print(f"\n[SKIP] {sheet_name}: assente in MAIN.xlsx")
                    continue
                s = process_sheet(wb.sheets[sheet_name], exp_id, conn, ts, only_owned)
                for k in totals:
                    totals[k] += s[k]
        finally:
            wb.close()
    finally:
        app.quit()

    print(f"\n=== Totali ===")
    print(f"  carte processate (chiamate API): {totals['processed']}")
    print(f"  con prezzo trovato:              {totals['priced']}")
    print(f"  saltate (non possedute):         {totals['skipped_unowned']}")
    print(f"  blueprint non trovato:           {totals['not_matched']}")

    regenerate_cache_xlsx(conn, cache_path)
    generate_formula_snippets(expansions)
    conn.close()


if __name__ == "__main__":
    main()
