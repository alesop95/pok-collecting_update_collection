"""
Path centralizzati del progetto. Tutti gli script importano da qui.

Logica: ogni file in scripts/ ha __file__ = .../update_collection/scripts/xxx.py.
La root del progetto e' quindi __file__.parent.parent. Questo rende gli
script eseguibili indipendentemente dalla cwd.

In aggiunta, qui carichiamo il .env del progetto. Poiche' tutti gli script
importano da paths.py (direttamente o transitivamente via cardtrader_client/
database/helpers), il caricamento avviene una sola volta all'inizio,
prima che qualsiasi codice provi a leggere os.environ['CT_AUTH_TOKEN'].
"""

from __future__ import annotations
from pathlib import Path
from dotenv import load_dotenv

# Root del progetto (update_collection/)
BASE_DIR = Path(__file__).resolve().parent.parent

# Carica .env dalla root del progetto. Se il file non esiste, ritorna False
# silenziosamente: il check vero su CT_AUTH_TOKEN avviene in cardtrader_client._auth_header.
load_dotenv(BASE_DIR / ".env")

# File di configurazione e dati persistenti — sempre nella root del progetto
EXPANSIONS_JSON  = BASE_DIR / "expansions.json"
PRICES_DB        = BASE_DIR / "prices.db"
CACHE_DIR        = BASE_DIR / ".cache"
LOOKUP_FORMULAS  = BASE_DIR / "lookup_formulas.txt"

# Default convenzionale per MAIN.xlsx (override-abile da CLI)
DEFAULT_WORKBOOK = BASE_DIR / "MAIN.xlsx"
