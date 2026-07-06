# Resume prompt

> Privato, ignorato da git. Prompt di ripresa della sessione: si aggiorna alla fine di ogni
> sessione con lo stato raggiunto e con un prompt pronto da incollare alla riapertura. Lo stato
> canonico del progetto resta in `.claude/memory/index.md`; questo file e la comodita per
> ripartire in fretta, non una seconda fonte di verita.

## Stato raggiunto

Data: <YYYY-MM-DD>
Branch / commit: <branch> / <hash>
Dove siamo: <una o due frasi sul punto raggiunto>
Prossimo passo: <azione concreta da cui ripartire>

## Da incollare a Claude alla riapertura

```
Riprendi il progetto <nome progetto>. Procedi cosi, senza leggere tutto:
1. Leggi .claude/memory/index.md (snapshot: branch, commit, stato schede, punto di ripresa).
2. Leggi .claude/context/current-work.md per la feature attiva.
3. Esegui la skill sync-context per misurare il drift delle schede rispetto a HEAD.
4. Dammi un recap conciso (dove siamo, cosa risulta fatto, prossimo passo) e fermati.
Vincoli: niente operazioni git (le faccio io); nessun valore segreto nei file tracciati.
```
