# PokèCollecting — tracker collezione Pokémon TCG

> Progetto personale (Alesop95, Milano) per tracciare il valore di una collezione Pokémon TCG via API CardTrader. Excel come interfaccia operativa, SQLite come storico autoritativo, Python come ETL.

**Per Claude (chat o Code):** se stai aprendo questo file in una nuova sessione, leggi tutto questo documento prima di toccare codice. Le scelte di architettura sono il risultato di un percorso di prove ed errori — i bug risolti sono sotto **Troubleshooting noto**, e ricaderci sarebbe regresso.

---

## Indice

1. [TL;DR e stato attuale](#tldr-e-stato-attuale)
2. [Architettura](#architettura)
3. [File del progetto](#file-del-progetto)
4. [Decisioni di design (perché così)](#decisioni-di-design-perché-così)
5. [Configurazione](#configurazione)
6. [Mappa colonne MAIN.xlsx](#mappa-colonne-mainxlsx)
7. [Mapping condizioni Cardmarket → CardTrader](#mapping-condizioni-cardmarket--cardtrader)
8. [Comandi comuni](#comandi-comuni)
9. [Troubleshooting noto](#troubleshooting-noto)
10. [Roadmap](#roadmap)
11. [Cosa NON è supportato](#cosa-non-è-supportato)
12. [Modalità operative (subagenti)](#modalità-operative-subagenti)

---

## TL;DR e stato attuale

**Cosa fa:** uno script Python legge MAIN.xlsx (read-only), interroga l'API CardTrader per ogni carta, scrive lo storico prezzi su SQLite e una vista corrente su `prices_cache.xlsx`. MAIN.xlsx ha una formula XLOOKUP nella colonna "Estimated Value" che pesca dal cache.

**Stato:**
- ✅ Client CardTrader API funzionante con caching dei blueprint e backoff su rate limit
- ✅ Mapping condizioni Cardmarket → CardTrader implementato e testato
- ✅ Pipeline read-only su MAIN.xlsx (niente più corruzioni)
- ✅ Storico su SQLite (`prices.db`, append-only)
- ✅ Vista corrente su `prices_cache.xlsx` rigenerata ogni run
- ✅ Generatore automatico di formule XLOOKUP (`lookup_formulas.txt`)
- ✅ Script `discover_expansions.py` per popolamento automatico expansions.json via API
- ✅ **Mappa colonne MAIN.xlsx verificata** sul file reale (55 fogli totali)
- ✅ **Decisioni ownership/quantity chiuse**: colonna I = "posseduta sì/no", filtro `only_owned=True` di default. Colonna H (Quantity) ignorata (utente non colleziona doppioni).
- ✅ **`prices_cache.xlsx` viene salvato accanto a MAIN.xlsx** (non in cwd) per garantire che XLOOKUP funzioni indipendentemente da dove si lancia lo script.
- ⏳ Compilazione di `expansions.json` con i set effettivi della collezione (via `discover_expansions.py`)
- ⏳ Primo run end-to-end su MAIN.xlsx reale
- ❓ Colonna L "Stamped": esiste una property dedicata su CardTrader (Prerelease/Staff/Worlds)? Per ora ignorata.
- ⏳ Carte graded (PSA/Beckett): non supportate, pianificata fase futura
- ⏳ Webapp (FastAPI + SQLite + HTML): in roadmap per "templatizzazione" futura, non prima della stabilizzazione del flusso Excel

---

## Architettura

```
                                                                  
    MAIN.xlsx  (read-only)                                        
        │                                                         
        │  xlwings via COM (Windows)                              
        ▼                                                         
    update_prices.py                                              
        │                                                         
        ├──►  cardtrader_client.py  ──►  api.cardtrader.com/v2    
        │         │                                               
        │         └──►  .cache/blueprints_*.json  (cache locale)  
        │                                                         
        ├──►  pricing.py    (mapping condizioni + aggregazione)   
        │                                                         
        ├──►  prices.db                       (SQLite, storico)   
        │                                                         
        ├──►  prices_cache.xlsx               (vista corrente)    
        │                                          ▲              
        │                                          │ XLOOKUP      
        └──►  lookup_formulas.txt   ──copy──►  MAIN.xlsx col O    
```

**Principio fondante:** lo script non scrive mai su MAIN.xlsx. La colonna "Estimated Value" è una formula, non un valore. Questo elimina alla radice la classe di bug di corruzione (formule perse, immagini eliminate, ecc.) che avevano bloccato lo sviluppo precedente.

---

## File del progetto

Root del progetto: `pokecollecting/update_collection/`. Layout standard "Python project" con `scripts/` separata. MAIN.xlsx vive nella root accanto a `expansions.json`, così tutti i percorsi sono coerenti.

```
update_collection/                     ← root del progetto (BASE_DIR)
├── CLAUDE.md                          ← questo file
├── CLAUDE.local.md                    ← settings personali (non versionato)
├── README.md
├── requirements.txt
├── expansions.json                    ← config (popolato da discover_expansions)
├── MAIN.xlsx                          ← workbook utente (mai modificato)
├── prices.db                          ← generato — SQLite storico
├── prices_cache.xlsx                  ← generato — vista corrente (accanto a MAIN.xlsx)
├── lookup_formulas.txt                ← generato — formule XLOOKUP
├── .venv/                             ← virtualenv
├── .cache/                            ← generata — blueprint export cachate
├── .claude/agents/                    ← subagent Claude Code
├── data/                              ← dati ausiliari utente (es. info_VINTED/)
└── scripts/                           ← TUTTI gli script Python
    ├── paths.py                       ← BASE_DIR + path centralizzati
    ├── cardtrader_client.py
    ├── database.py
    ├── pricing.py
    ├── helpers.py
    ├── update_prices.py
    └── discover_expansions.py
```

### Path resolution

Tutti gli script importano da `paths.py` che calcola `BASE_DIR = Path(__file__).resolve().parent.parent`. Questo significa che **i path importanti sono sempre assoluti e indipendenti dalla cwd**: puoi lanciare `python scripts/update_prices.py` dalla root, oppure `python update_prices.py` da dentro `scripts/`, oppure dare il path completo da `C:\altro\`. Il risultato è sempre lo stesso.

`MAIN.xlsx` ha un default convenzionale (`BASE_DIR / "MAIN.xlsx"`), quindi se sta nella root del progetto puoi omettere l'argomento.

### Ruolo dei singoli file

| File                    | Ruolo                                                                        |
|-------------------------|------------------------------------------------------------------------------|
| `CLAUDE.md`             | Master index per onboarding/handoff (questo file).                           |
| `CLAUDE.local.md`       | Settings personali (path personale, note). NON committare.                   |
| `scripts/paths.py`      | Calcola BASE_DIR e centralizza tutti i path costanti.                        |
| `scripts/cardtrader_client.py` | Client API + cache disco blueprint + backoff su rate limit            |
| `scripts/pricing.py`    | Mapping condizioni Cardmarket → CardTrader + aggregazione prezzo             |
| `scripts/database.py`   | SQLite: schema, insert_snapshot, latest_per_card, price_history              |
| `scripts/helpers.py`    | Utility pure: parse_bool, normalize_collector, normalize_language, is_owned, find_blueprint_robust, build_lookup_key |
| `scripts/update_prices.py` | Pipeline principale: legge MAIN.xlsx, scrive su SQLite + cache.xlsx       |
| `scripts/discover_expansions.py` | Popola expansions.json matchando fogli MAIN.xlsx vs API CT          |
| `expansions.json`       | Mappa nome_foglio → expansion_id CardTrader                                  |
| `requirements.txt`      | Dipendenze: `requests`, `xlwings`, `openpyxl`                                 |
| `prices.db`             | Generato — SQLite, append-only. Nella root del progetto.                     |
| `prices_cache.xlsx`     | Generato — accanto a MAIN.xlsx (= root del progetto, in questa config).       |
| `lookup_formulas.txt`   | Generato — formule XLOOKUP pronte da incollare in MAIN.xlsx.                  |

---

## Decisioni di design (perché così)

Queste sono le scelte che hanno richiesto iterazione e tempo. **Cambiarle senza motivo è regresso.**

1. **CardTrader come unica fonte**, non Cardmarket. Cardmarket ha l'API pubblica chiusa da ~2024, accesso solo manuale per partner approvati. Non è un problema tecnico aggirabile. CardTrader espone REST documentata, accesso ottenibile, struttura coerente.

2. **`xlwings` (non `openpyxl`) per MAIN.xlsx**. `openpyxl` riscrive l'XML da zero e perde immagini, hyperlink, formule. `xlwings` pilota Excel via COM, il file resta integro. Costo: serve Excel installato (siamo su Windows, OK).

3. **MAIN.xlsx read-only.** Lo script lo apre con `read_only=True, update_links=False`. Tutta la scrittura va in file separati. Anche con xlwings, ogni scrittura è un rischio: tagliarla via è la difesa definitiva.

4. **SQLite come fonte autoritativa, cache.xlsx come vista derivata.** SQLite è append-only: ogni run aggiunge righe, mai sovrascrive. `prices_cache.xlsx` è rigenerato da zero ogni volta dal database — usa-e-getta. Vantaggio: storico gratuito per grafici di andamento + zero rischio di "fonte di verità" che si corrompe in scrittura.

5. **Token via variabile d'ambiente, non hardcoded.** `CT_AUTH_TOKEN`. Nei vecchi script il JWT era nel codice — committare per errore una repo è banale, il token va isolato.

6. **`requests` (non `subprocess` su cURL).** I vecchi script usavano `subprocess.run(['curl', ...])` ed ereditavano tutti i problemi di PATH, encoding Windows, parsing stdout. `requests` è più affidabile e testabile.

7. **`pokemon_reverse` (non `foil`) come property booleana per Pokémon.** CardTrader nella categoria "Pokémon Singles" usa `pokemon_reverse: true/false` per indicare il reverse holo, e `first_edition` separato. Non c'è un campo `foil` generico per Pokémon.

8. **Chiave XLOOKUP concatenata: `sheet|collector|condition|language|reverse|first_edition`.** Univoca per la combinazione (carta posseduta, stato). Generata identica sia da Python (`build_lookup_key`) sia da Excel (formula). Se modifichi il formato in uno, modifica anche l'altro.

9. **`collector_number` normalizzato (`026/132` → `26/132`).** L'API risponde a volte con padding zero, a volte senza. La normalizzazione lato Python evita falsi negativi nel matching. Em-dash (`—`), `-`, vuoto → `None` (carte senza collector, tipiche Vending Machine).

10. **Filtro ownership di default (`only_owned=True`).** La colonna I "yes/no" di MAIN.xlsx contiene "Y" per le carte effettivamente possedute (formattazione condizionale lato utente). I fogli sono cataloghi del set completo, popolati con tutte le carte del set: senza filtro si farebbero chiamate API per migliaia di carte non possedute. Override con flag CLI `--all`.

11. **`prices_cache.xlsx` accanto a MAIN.xlsx, non in cwd.** La formula XLOOKUP fa riferimento al file con path relativo (`'[prices_cache.xlsx]Latest'!`), quindi deve stare nella stessa cartella di MAIN.xlsx. Lo script lo deduce da `Path(args.workbook).parent`. SQLite invece resta nella root del progetto (dato "operativo", non legato al workbook).

12. **Path centralizzati via `scripts/paths.py`.** Tutti i percorsi importanti (config, db, cache, default workbook) sono calcolati a partire da `BASE_DIR = Path(__file__).resolve().parent.parent`. Questo rende gli script eseguibili indipendentemente dalla cwd: `python scripts/update_prices.py` da root oppure `python update_prices.py` da scripts/ producono lo stesso risultato. Mai usare `Path(".cache")` o `Path("expansions.json")` (path relativi alla cwd, fragili).

13. **Secret via `.env` (project-scoped), non `setx` (system-wide).** Il `CT_AUTH_TOKEN` vive in `update_collection/.env`, caricato da `python-dotenv` dentro `paths.py`. Motivazione: il token appartiene al progetto, non alla workstation. `setx` accoppia macchina e progetto (rimane nel profilo Windows anche se cancelli il repo), espone il token a *ogni* script Python dell'utente, e non è portabile a CI/CD o container. `.env` confina il secret al progetto, è ignorato da Git, e l'env var di processo lo sovrascrive se serve (utile in pipeline automatizzate). Tradeoff: serve installare `python-dotenv` (dipendenza minima, già in `requirements.txt`).

---

## Configurazione

**`expansions.json`** — popolato automaticamente da `discover_expansions.py`. Mappa il nome esatto del foglio in MAIN.xlsx all'`expansion_id` di CardTrader. Esempio:

```json
{
  "GYM_HEROES": 1480,
  "BASE_SET": 1469,
  "JUNGLE": 1470,
  "FOSSIL": 1471
}
```

I valori `null` rappresentano fogli per cui il fuzzy match non ha trovato un'espansione: vanno risolti a mano (usa l'agente `cardtrader-api-explorer`).

**`.env`** (root del progetto, NON committato) — contiene `CT_AUTH_TOKEN`. Crealo copiando `.env.example`:

```cmd
copy .env.example .env
notepad .env
```

Contenuto:

```
CT_AUTH_TOKEN=eyJhbGciOiJSUzI1NiJ9...
```

Caricato automaticamente da `scripts/paths.py` via `python-dotenv` all'avvio di qualsiasi script. Override possibile via env var del processo (CI/CD): se `CT_AUTH_TOKEN` è già nell'ambiente, `.env` non la sovrascrive.

Verifica che il token sia leggibile:

```cmd
python -c "from dotenv import dotenv_values; print('OK' if dotenv_values().get('CT_AUTH_TOKEN','').startswith('eyJ') else 'MANCANTE')"
```

**`COLUMNS`** in `scripts/update_prices.py` — mappa nomi semantici → lettere colonne MAIN.xlsx. Vedi sezione successiva.

---

## Mappa colonne MAIN.xlsx

**Verificata sul file reale.** Layout uguale in tutti i 55 fogli set.

- **Riga 1**: vuota
- **Riga 2**: `B='HOME' C='Link' D=<URL Bulbapedia del set>`
- **Riga 3**: `C='Card set name' D=<nome ufficiale set>` ← usato da `discover_expansions.py` per matching API
- **Riga 4**: vuota
- **Riga 5**: header colonne
- **Riga 6+**: dati carte

| Lettera | Colonna           | Letta dallo script? | Note                                                        |
|---------|-------------------|---------------------|-------------------------------------------------------------|
| C       | Rarity            | No                  | Metadata utente                                              |
| D       | No.               | **Sì**              | Collector number (es. "26/132"). `—`/`-` → None (Vending Machine). |
| E       | Card name         | **Sì**              | Nome carta. Chiave per blueprint match.                      |
| F       | Type              | No                  | Metadata                                                     |
| G       | Promotion         | No                  | Metadata                                                     |
| H       | Quantity          | No                  | Quanti esemplari (di solito "-")                             |
| I       | yes/no            | ❓ Dipende           | **Probabile: "Y" = posseduta**. Decisione aperta sul filtraggio. |
| J       | Language          | **Sì**              | "English"/"Italian"/... → normalizzato a "en"/"it" via `normalize_language`. Default `en` se vuoto. |
| K       | 1st edition       | **Sì**              | Boolean. Filtro prezzo (`first_edition` su CT).              |
| L       | Stamped           | No                  | Ignorato (CardTrader non ha property dedicata).              |
| M       | Reverse Holo      | **Sì**              | Boolean. Filtro prezzo (`pokemon_reverse` su CT).            |
| N       | Condition         | **Sì**              | Sigla Cardmarket (NM/LP/EX/...). Mappata via `pricing.py`. Default "Near Mint" se vuoto. |
| O       | Estimated Value   | No (output)         | **Formula XLOOKUP** che pesca da `prices_cache.xlsx`.        |
| P       | Notes             | No                  | Metadata                                                     |

`START_ROW = 6`, `END_ROW = ultima riga con valore in colonna E (Card name)`.

### Domande aperte sui dati di MAIN.xlsx

1. **Colonna I "yes/no"**: cosa significa? Ipotesi attuale è "carta posseduta sì/no". Se confermata, vale la pena aggiungere un filtro `only_owned=True` per non interrogare l'API su 5000+ carte non possedute. Da chiedere all'utente prima del primo run.
2. **Colonna L "Stamped"**: c'è una property CardTrader per i timbri (Prerelease, Staff, World Championship)? Da verificare con `cardtrader-api-explorer`.
3. **Colonna H "Quantity"**: sempre "-" nei sample visti. Da capire se l'utente la usa o no.

### Fogli (55 totali)

**Set Pokémon WOTC era (mappabili 1:1 a CardTrader):**
BASE_SET, JUNGLE, FOSSIL, BASE_SET_2, TEAM_ROCKET, GYM_HEROES, GYM_CHALLENGE, NEO_GENESIS, NEO_DISCOVERY, NEO_REVELATION, NEO_DESTINY, LEGENDARY_COLLECTION, EXPEDITION, ACQUAPOLIS, SKYRIDGE.

**Set EX era (Nintendo, 2003-2007):**
EX_RUBY_SAPPHIRE, EX_SANDSTORM, EX_DRAGON, EX_TEAM_MAGMA_ACQUA, EX_HIDDEN_LEGENDS, EX_FIRERED_LEAFGREEN, EX_TEAM_ROCKET_RETURNS, EX_DEOXYS, EX_EMERALD, EX_UNSEEN_FORCES, EX_DELTA_SPECIES, EX_LEGEND_MAKER, EX_HOLON_PHANTOMS, EX_CRYSTAL_GUARDIANS, EX_DRAGON_FRONTIERS, EX_POWER_KEEPERS.

**Set/raccolte vintage giapponesi (mapping API incerto):**
SOUTHERN_ISLAND, WOC, WOC_PROMO, VENDING_MACHINE_CARDS, VMC_SERIES1_BLUE, VMC_SERIES2_RED, VMC_SERIES3_GREEN, VS, VS_PROMOTIONAL, THEATER_LTD_VS_PACK, MOVIE_VS_DEOXYS, MOVIE_VS_LUCARIO, BEST_OF_GAME, CARDASS, PM_CARDDASS, UNNUMBERED_PROMOTIONAL, ANA_CAMPAIGN, MASAKI, UNNUMBERED_OWNER.

**Fogli da escludere (non set):**
NOTES_AND_TO_DO_LIST, INDEX_TCG, PROMO (vuoto), EX-SERIES (raggruppamento), VS_ERA (raggruppamento). Gestiti in `discover_expansions.EXCLUDED_SHEETS`.

---

## Mapping condizioni Cardmarket → CardTrader

Implementato in `pricing.py`, funzione `normalize_condition`.

| Utente scrive (Cardmarket) | Lo script invia (CardTrader) |
|----------------------------|-------------------------------|
| M (Mint)                   | Near Mint                     |
| NM (Near Mint)             | Near Mint                     |
| EX (Excellent)             | Slightly Played               |
| LP (Light Played)          | Moderately Played             |
| GD (Good)                  | Played                        |
| PL (Played)                | Played                        |
| PO (Poor)                  | Poor                          |

"Mint" non esiste su CardTrader (considerato teorico). La perdita di precisione è limitata alla fascia EX/LP/GD compressa in SP/MP/PL. Per il documento di analisi: "non esiste corrispondenza uno-a-uno perfetta, la traduzione va lato ETL non in Excel".

---

## Comandi comuni

**Setup iniziale (una volta):**

```cmd
cd C:\percorso\pokecollecting\update_collection
.venv\Scripts\activate
pip install -r requirements.txt

copy .env.example .env
notepad .env
:: Sostituisci il placeholder con il tuo JWT CardTrader vero, salva e chiudi.
```

**Run quotidiano** (MAIN.xlsx nella root del progetto, default):

```cmd
python scripts\update_prices.py
```

**Run di test su un singolo foglio:**

```cmd
python scripts\update_prices.py --sheet GYM_HEROES
```

**Run includendo anche carte NON possedute** (per valutare prezzi di carte da comprare):

```cmd
python scripts\update_prices.py --all
```

**Prima configurazione di `expansions.json` (una volta):**

```cmd
python scripts\discover_expansions.py
```

Lo script propone un mapping automatico foglio → expansion_id via fuzzy match contro l'API CardTrader. I match incerti (60-80% di similarità) li indica con `[?]` per verifica manuale.

**MAIN.xlsx fuori dalla root** (override del default):

```cmd
python scripts\update_prices.py "C:\altro\percorso\MAIN.xlsx"
```

**Query SQLite (Python REPL):**

```python
from database import get_connection, latest_per_card, price_history
conn = get_connection()

# valore stimato totale collezione (ultimo snapshot per ogni carta)
total = sum(r['price_eur'] or 0 for r in latest_per_card(conn))
print(f"€ {total:.2f}")

# storia di una carta
for r in price_history(conn, blueprint_id=120778):
    print(r['ts'], r['condition_ct'], r['price_eur'])
```

**Discovery `expansion_id` manuale (in alternativa al subagente):**

```cmd
:: Estrai il token dal .env e usalo per la query
for /f "tokens=2 delims==" %t in ('findstr CT_AUTH_TOKEN .env') do set _TOK=%t
curl "https://api.cardtrader.com/api/v2/expansions" -H "Authorization: Bearer %_TOK%" | jq ".[] | select(.game_id==5)" > expansions_pokemon.json
set _TOK=
```

**Force refresh cache blueprint (raro, solo dopo nuovi rilasci):**

```python
from cardtrader_client import get_blueprints
get_blueprints(1480, force_refresh=True)
```

---

## Troubleshooting noto

Problemi storici risolti. **Sintomo → causa → fix.** Ricaderci è regresso.

1. **MAIN.xlsx corrotto, formule perse, immagini scomparse.**
   Causa: `openpyxl` riscrive l'XML da zero. Fix: usare `xlwings`. **Mai** scrivere su MAIN.xlsx con openpyxl.

2. **Lo script scrive lo stesso valore da M31 a M57 invece che solo M31.**
   Causa: iterazione su `DataFrame.values` con `enumerate(start=start_row)`. Fix: loop esplicito `for row in range(start_row, end_row + 1)` con `sheet.range(f"{col}{row}").value = scalar`. **Mai** usare DataFrame iteration per scrittura cella-singola.

3. **"Failed writing body" su `curl > output.json`.**
   Causa: path con spazi non quotato. Fix: quotare l'intero path di output. Nel codice attuale non rilevante (usiamo `requests`, non `curl`).

4. **`jq` non riconosciuto.**
   Causa: file chiamato `jq-windows-amd64.exe`. Fix: rinominarlo `jq.exe` o usare il nome completo. Nel codice attuale non rilevante.

5. **Errore "writing dictionary to cell".**
   Causa: si scriveva l'intero JSON di risposta API. Fix: estrarre solo il campo scalare (`price.cents/100`) prima della scrittura.

6. **Token nel codice committato per sbaglio.**
   Causa: hardcoded nelle prime versioni. Fix: `.env` letto via `python-dotenv` da `paths.py`, file ignorato da Git. **Mai** scrivere il token in chiaro nei file `.py` o committare `.env`.

7. **403 sull'API Cardmarket.**
   Causa strutturale: API pubblica chiusa da Cardmarket. Fix: abbandonare Cardmarket, restare su CardTrader. Non è un problema tecnico aggirabile.

---

## Roadmap

- **Done (step 1)** — Pipeline base funzionante con xlwings, scrittura diretta su MAIN.xlsx (poi abbandonata).
- **Done (step 2)** — Refactor read-only: SQLite + cache.xlsx + formula XLOOKUP.
- **Next** — Primo run end-to-end su MAIN.xlsx reale, verifica mappa colonne, popolamento `expansions.json`.
- **Possibile estensione** — Carte graded (PSA/Beckett): integrare PSA Sales History come secondo provider.
- **Step 4 (templatizzazione)** — Mini-webapp FastAPI + SQLite + HTML, deploy locale o su VPS economico. Solo dopo che il flusso Excel è stabile e l'utente vuole accesso multi-device.

---

## Cosa NON è supportato

- **Carte graded (PSA/Beckett/CGC):** CardTrader non espone il grade numerico. Le slab graduate vanno stimate a mano o aggiunte da un secondo provider in futuro.
- **"Stamped" (colonna L):** non c'è una property CardTrader dedicata. Ignorato dal filtro: il prezzo ritornato è della versione standard, non timbrata.
- **Match nome ambiguo senza collector_number:** se due carte stesso nome stesso set, il `collector_number` (colonna D) è obbligatorio. Senza, lo script prende il primo blueprint trovato (warning a video).
- **Lingue minoritarie:** lo script supporta tutte le lingue CardTrader (`en/fr/it/de/es/pt/nl/ru/pl/sv/kr/jp/zh-CN/zh-TW/id/th`), ma se nel marketplace non ci sono offerte in quella lingua + condizione → "n/d".

---

## Modalità operative (subagenti)

I subagenti in `.claude/agents/` sono "modalità" specializzate. Su Claude Code vengono caricati automaticamente. Su claude.ai puoi allegare il file relativo quando vuoi attivare quella modalità.

| Subagente                  | Quando usarlo                                                                |
|----------------------------|------------------------------------------------------------------------------|
| `cardtrader-api-explorer`  | Trovare `expansion_id`, debug blueprint mancanti, capire properties, esplorare risposte API |
| `sqlite-analyst`           | Query sullo storico, valore collezione, trend, top mover, generare report    |
| `pipeline-debugger`        | Errori dello script, righe non lette, formule XLOOKUP che ritornano "n/d", problemi di matching |

**Convenzione di attivazione su claude.ai:** "Entra in modalità `<nome-agente>` (file allegato)" + allega il `.md` corrispondente.
