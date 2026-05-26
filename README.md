# PokèCollecting

Tracker prezzi per collezione Pokémon TCG. Excel come UI operativa, SQLite come storico, CardTrader API come fonte prezzi.

## Stack

- **Python 3.10+** con `requests`, `xlwings`, `openpyxl`, `python-dotenv`
- **Excel** (locale, via xlwings COM) — il workbook personale resta inalterato, lo script lo legge soltanto
- **SQLite** — storico prezzi append-only (`prices.db`)
- **CardTrader v2 REST API** — unica fonte dati (Cardmarket API pubblica è chiusa dal 2024)

## Struttura

```
update_collection/                     ← root del progetto (BASE_DIR)
├── .env                               ← secret (CT_AUTH_TOKEN) — NON committato
├── .env.example                       ← template del .env
├── .gitignore
├── CLAUDE.md                          ← documentazione tecnica completa
├── README.md                          ← questo file
├── requirements.txt
├── expansions.json                    ← mappa fogli → expansion_id (autogenerata)
├── MAIN.xlsx                          ← workbook utente
├── prices.db                          ← storico SQLite (generato)
├── prices_cache.xlsx                  ← vista corrente per XLOOKUP (generato)
├── lookup_formulas.txt                ← formule pronte da incollare (generato)
├── .venv/
├── .cache/                            ← blueprint export cachate (generato)
├── .claude/agents/                    ← subagent specializzati per Claude
└── scripts/
    ├── paths.py                       ← path centralizzati + load_dotenv
    ├── cardtrader_client.py
    ├── pricing.py
    ├── database.py
    ├── helpers.py
    ├── update_prices.py               ← pipeline principale
    └── discover_expansions.py         ← discovery automatica expansion_id
```

## Setup

```cmd
cd C:\percorso\pokecollecting\update_collection
.venv\Scripts\activate
pip install -r requirements.txt

copy .env.example .env
notepad .env
:: sostituisci il placeholder con il tuo JWT CardTrader e salva
```

Verifica che il token sia stato caricato:

```cmd
python -c "from dotenv import dotenv_values; print('OK' if dotenv_values().get('CT_AUTH_TOKEN','').startswith('eyJ') else 'MANCANTE')"
```

## Uso

**Prima volta — popola `expansions.json` (~10 secondi):**

```cmd
python scripts\discover_expansions.py
```

Lo script matcha fuzzy ogni foglio di MAIN.xlsx con le espansioni Pokémon dell'API. I match marcati `[?]` o `null` vanno risolti a mano nel file generato.

**Test su un singolo foglio:**

```cmd
python scripts\update_prices.py --sheet GYM_HEROES
```

**Run completo (solo carte possedute, default):**

```cmd
python scripts\update_prices.py
```

**Includendo anche carte non possedute:**

```cmd
python scripts\update_prices.py --all
```

Al primo run completo, copia le formule da `lookup_formulas.txt` nella colonna `O` di MAIN.xlsx — una formula per foglio, trascinata dalla riga 6 fino all'**ultima riga del catalogo** (anche le carte non possedute: mostreranno "n/d" finché non scrivi "Y" in colonna I).

> **Excel italiano**: le formule generate sono in italiano (`CERCA.X`/`SE.ERRORE`/`MAIUSC`/`MINUSC`/`SE`) con separatore `;`. Se hai Excel inglese o regionale diverso, va modificata la funzione `generate_formula_snippets` in `scripts/update_prices.py`.

**Quando ottieni una carta nuova**: il catalogo è già pre-popolato in ogni foglio, quindi non devi aggiungere righe. Trovi la riga della carta nel set giusto, metti "Y" in colonna I, salvi. Al prossimo run lo script la prezza automaticamente — la formula CERCA.X la trovi già lì.

## Per dettagli tecnici

Tutto è in `CLAUDE.md`: architettura, decisioni di design, troubleshooting noto, query SQLite utili. È pensato per onboarding rapido di una nuova istanza di Claude o di un nuovo sviluppatore (anche il futuro te).

## Sicurezza

- Il `.env` con il token CardTrader è **ignorato da Git** (vedi `.gitignore`).
- `MAIN.xlsx` con la collezione personale è **ignorato da Git** per default. Se vuoi versionarlo (es. backup su repo privato), togli la riga corrispondente dal `.gitignore`.
- Nessun secret deve mai finire negli script `.py`.
